<?php

declare(strict_types=1);

class PreheatScheduler extends IPSModule
{
    private const STATUS_OK = 102;
    private const STATUS_URL_ERROR = 201;
    private const STATUS_AUTH_ERROR = 202;

    public function Create()
    {
        parent::Create();

        $this->RegisterPropertyString('CalendarURL', '');
        $this->RegisterPropertyString('CalUser', '');
        $this->RegisterPropertyString('CalPass', '');
        $this->RegisterPropertyFloat('SetpointWarm', 21.0);
        $this->RegisterPropertyFloat('HeatingRate', 1.0);
        $this->RegisterPropertyInteger('TempVarID', 0);
        $this->RegisterPropertyInteger('LookaheadHours', 36);
        $this->RegisterPropertyInteger('PreheatBufferMin', 0);
        $this->RegisterPropertyInteger('EvaluationIntervalSec', 60);
        $this->RegisterPropertyInteger('HoldStrategy', 0);
        $this->RegisterPropertyInteger('UpcomingEventsCount', 5);

        $heatingVarID = $this->RegisterVariableBoolean('HeatingDemand', $this->Translate('Heating Demand'));
        IPS_SetVariableCustomProfile($heatingVarID, '~Switch');
        $nextEventVarID = $this->RegisterVariableString('NextEventStartISO', $this->Translate('Next Event Start'));
        IPS_SetVariableCustomProfile($nextEventVarID, '~String');
        $nextPreheatVarID = $this->RegisterVariableString('NextPreheatStartISO', $this->Translate('Next Preheat Start'));
        IPS_SetVariableCustomProfile($nextPreheatVarID, '~String');
        $upcomingEventsVarID = $this->RegisterVariableString('UpcomingEventsHTML', $this->Translate('Upcoming Events'));
        $this->AssignHTMLProfile($upcomingEventsVarID);

        $this->SetValue('HeatingDemand', false);
        $this->SetValue('NextEventStartISO', '-');
        $this->SetValue('NextPreheatStartISO', '-');
        $this->SetValue('UpcomingEventsHTML', $this->BuildUpcomingEventsTable([], time(), 0.0, 0.0, null, 0));

        $this->RegisterTimer('Evaluate', 0, 'HEAT_Recalculate($_IPS[\'TARGET\']);');

        $this->RegisterAttributeInteger('RegisteredTempVarID', 0);
        $this->RegisterAttributeInteger('LastEventStart', 0);
        $this->RegisterAttributeInteger('LastEventEnd', 0);
        $this->RegisterAttributeInteger('LastPreheatStart', 0);
        $this->RegisterAttributeString('LastEventSummary', '');
    }

    private function AssignHTMLProfile(int $variableID): void
    {
        $profile = '~HTMLBox';
        if (!IPS_VariableProfileExists($profile)) {
            $fallback = '~TextBox';
            if (IPS_VariableProfileExists($fallback)) {
                $profile = $fallback;
                $this->Log('HTMLBox profile not available. Falling back to ~TextBox.');
            } else {
                $profile = '';
                $this->Log('HTMLBox profile not available and no fallback profile found.');
            }
        }

        if ($profile !== '') {
            IPS_SetVariableCustomProfile($variableID, $profile);
        }
    }

    public function ApplyChanges()
    {
        parent::ApplyChanges();

        $interval = max(15, $this->ReadPropertyInteger('EvaluationIntervalSec'));
        $this->SetTimerInterval('Evaluate', $interval * 1000);

        $tempVarID = $this->ReadPropertyInteger('TempVarID');
        $lastRegistered = $this->ReadAttributeInteger('RegisteredTempVarID');

        if ($lastRegistered > 0 && $lastRegistered !== $tempVarID) {
            $this->UnregisterMessage($lastRegistered, VM_UPDATE);
            $this->WriteAttributeInteger('RegisteredTempVarID', 0);
        }

        if ($tempVarID > 0 && IPS_VariableExists($tempVarID)) {
            $this->RegisterMessage($tempVarID, VM_UPDATE);
            $this->WriteAttributeInteger('RegisteredTempVarID', $tempVarID);
        }

        $this->Recalculate();
    }

    public function MessageSink($TimeStamp, $SenderID, $Message, $Data)
    {
        if ($Message === VM_UPDATE) {
            $tempVarID = $this->ReadPropertyInteger('TempVarID');
            if ($SenderID === $tempVarID) {
                $this->Recalculate();
            }
        }
    }

    public function Recalculate(): bool
    {
        $now = time();
        $upcomingEvents = [];
        $event = $this->DetermineNextEvent($now, $upcomingEvents);
        $holdStrategy = $this->ReadPropertyInteger('HoldStrategy');

        $tempVarID = $this->ReadPropertyInteger('TempVarID');
        $currentTemp = null;
        if ($tempVarID > 0 && IPS_VariableExists($tempVarID)) {
            $currentTemp = (float) GetValue($tempVarID);
        }

        if ($currentTemp === null) {
            $this->Log('Temperature variable not set or missing.');
        }

        $setpoint = $this->ReadPropertyFloat('SetpointWarm');
        $heatingRate = $this->ReadPropertyFloat('HeatingRate');
        if ($heatingRate <= 0.0) {
            $this->Log('Heating rate must be greater than zero.');
        }

        $bufferSeconds = max(0, $this->ReadPropertyInteger('PreheatBufferMin')) * 60;

        $heatingVarID = $this->GetIDForIdent('HeatingDemand');
        $currentlyOn = GetValueBoolean($heatingVarID);

        $shouldBeOn = false;
        $eventStartISO = '-';
        $preheatStartISO = '-';

        if ($event !== null) {
            $eventStart = $event['start'];
            $eventEnd = $event['end'];
            $eventStartISO = date('c', $eventStart);

            $preheatStart = $this->CalculatePreheatStart($eventStart, $setpoint, $heatingRate, $currentTemp, $bufferSeconds);
            if ($heatingRate <= 0.0) {
                $this->Log('Unable to calculate preheat start because heating rate is zero or negative.');
            }

            $this->WriteAttributeInteger('LastPreheatStart', $preheatStart);
            $preheatStartISO = $preheatStart > 0 ? date('c', $preheatStart) : '-';

            $windowEnd = $holdStrategy === 0 ? $eventEnd : $eventStart;
            if ($windowEnd < $eventStart) {
                $windowEnd = $eventStart;
            }

            if ($now >= $preheatStart && $now < $windowEnd) {
                $shouldBeOn = true;
            }

            if ($now >= $eventStart && $now < $eventEnd) {
                $shouldBeOn = true;
            }

            if (!$shouldBeOn && $currentlyOn) {
                if ($holdStrategy === 0 && $now < $eventEnd) {
                    $shouldBeOn = true;
                }
            }
        } else {
            $stored = $this->GetStoredEvent();
            if ($stored !== null) {
                if ($now < $stored['end']) {
                    if ($currentlyOn) {
                        $shouldBeOn = true;
                    }
                }
            }
        }

        $this->SetValue('NextEventStartISO', $eventStartISO);
        $this->SetValue('NextPreheatStartISO', $preheatStartISO);
        $this->SetValue('UpcomingEventsHTML', $this->BuildUpcomingEventsTable($upcomingEvents, $now, $setpoint, $heatingRate, $currentTemp, $bufferSeconds));

        if ($shouldBeOn !== $currentlyOn) {
            $this->SetValue('HeatingDemand', $shouldBeOn);
        }

        return $shouldBeOn;
    }

    public function GetHeatingDemand(): bool
    {
        $this->Recalculate();
        return GetValueBoolean($this->GetIDForIdent('HeatingDemand'));
    }

    public function ForceRefresh(): void
    {
        $this->Recalculate();
        $interval = max(15, $this->ReadPropertyInteger('EvaluationIntervalSec'));
        $this->SetTimerInterval('Evaluate', $interval * 1000);
    }

    private function DetermineNextEvent(int $now, ?array &$upcomingEvents = null): ?array
    {
        $calendarUrl = trim($this->ReadPropertyString('CalendarURL'));
        $collectedEvents = [];
        if ($calendarUrl === '') {
            $this->Log('Calendar URL is not configured.');
            $this->SetStatus(self::STATUS_URL_ERROR);
            $event = $this->GetStoredEventForFallback($now);
            if ($event !== null) {
                $collectedEvents[] = $event;
            }
            if ($upcomingEvents !== null) {
                $upcomingEvents = $collectedEvents;
            }
            return $event;
        }

        $content = $this->FetchCalendarContent($calendarUrl);
        if ($content === null) {
            $event = $this->GetStoredEventForFallback($now);
            if ($event !== null) {
                $collectedEvents[] = $event;
            }
            if ($upcomingEvents !== null) {
                $upcomingEvents = $collectedEvents;
            }
            return $event;
        }

        $events = $this->ParseICSEvents($content);
        if (empty($events)) {
            $this->Log('No events found in calendar export.');
            $this->WriteAttributeInteger('LastEventStart', 0);
            $this->WriteAttributeInteger('LastEventEnd', 0);
            $this->WriteAttributeInteger('LastPreheatStart', 0);
            $this->WriteAttributeString('LastEventSummary', '');
            $this->SetStatus(self::STATUS_OK);
            if ($upcomingEvents !== null) {
                $upcomingEvents = [];
            }
            return null;
        }

        $lookaheadSeconds = max(1, $this->ReadPropertyInteger('LookaheadHours')) * 3600;
        $horizon = $now + $lookaheadSeconds;

        $filtered = [];
        foreach ($events as $event) {
            if (!isset($event['start'], $event['end'])) {
                continue;
            }
            if ($event['end'] <= $now) {
                continue;
            }
            if ($event['start'] <= $now && $event['end'] > $now) {
                $filtered[] = $event;
                continue;
            }
            if ($event['start'] > $horizon) {
                continue;
            }
            if ($event['start'] >= $now) {
                $filtered[] = $event;
            }
        }

        usort($filtered, function (array $a, array $b): int {
            return $a['start'] <=> $b['start'];
        });

        $nextEvent = $filtered[0] ?? null;
        $limit = max(1, $this->ReadPropertyInteger('UpcomingEventsCount'));
        $collectedEvents = array_slice($filtered, 0, $limit);

        if ($nextEvent !== null) {
            $this->WriteAttributeInteger('LastEventStart', $nextEvent['start']);
            $this->WriteAttributeInteger('LastEventEnd', $nextEvent['end']);
            $this->WriteAttributeInteger('LastPreheatStart', 0);
            $this->WriteAttributeString('LastEventSummary', $nextEvent['summary'] ?? '');
            $this->SetStatus(self::STATUS_OK);
        } else {
            $this->WriteAttributeString('LastEventSummary', '');
            $this->SetStatus(self::STATUS_OK);
        }

        if ($upcomingEvents !== null) {
            $upcomingEvents = $collectedEvents;
        }

        return $nextEvent;
    }

    private function FetchCalendarContent(string $calendarUrl): ?string
    {
        $user = $this->ReadPropertyString('CalUser');
        $pass = $this->ReadPropertyString('CalPass');

        $urlsToTry = [];
        $trimmed = rtrim($calendarUrl);
        if (!preg_match('/\.ics($|\?)/i', $trimmed) && !str_contains($trimmed, '?export')) {
            $separator = str_contains($trimmed, '?') ? '&' : '?';
            $urlsToTry[] = $trimmed . $separator . 'export';
        }
        $urlsToTry[] = $trimmed;

        $auth = [];
        if ($user !== '' || $pass !== '') {
            $auth['AuthUser'] = $user;
            $auth['AuthPass'] = $pass;
        }

        $lastError = '';
        foreach ($urlsToTry as $url) {
            error_clear_last();
            $content = @Sys_GetURLContentEx($url, $auth);
            if ($content !== false && $content !== null) {
                if ($url !== $trimmed) {
                    $this->Log('Calendar fetched using export helper URL: ' . $url);
                }
                $this->SetStatus(self::STATUS_OK);
                return $content;
            }
            $error = error_get_last();
            if ($error !== null) {
                $lastError = $error['message'] ?? '';
            }
        }

        if ($lastError !== '') {
            $this->Log('Calendar fetch failed: ' . $lastError);
        } else {
            $this->Log('Calendar fetch failed: unknown error');
        }

        if ($lastError !== '' && (stripos($lastError, '401') !== false || stripos($lastError, '403') !== false)) {
            $this->SetStatus(self::STATUS_AUTH_ERROR);
        } else {
            $this->SetStatus(self::STATUS_URL_ERROR);
        }

        return null;
    }

    private function ParseICSEvents(string $content): array
    {
        $lines = preg_split('/\r\n|\n|\r/', $content);
        if ($lines === false) {
            return [];
        }

        $unfolded = [];
        foreach ($lines as $line) {
            if ($line === '') {
                $unfolded[] = '';
                continue;
            }
            $firstChar = $line[0];
            if (($firstChar === ' ' || $firstChar === '\t') && !empty($unfolded)) {
                $lastIndex = count($unfolded) - 1;
                $unfolded[$lastIndex] .= substr($line, 1);
            } else {
                $unfolded[] = $line;
            }
        }

        $events = [];
        $inEvent = false;
        $currentLines = [];
        foreach ($unfolded as $line) {
            if (strtoupper($line) === 'BEGIN:VEVENT') {
                $inEvent = true;
                $currentLines = [];
                continue;
            }
            if (strtoupper($line) === 'END:VEVENT') {
                if (!empty($currentLines)) {
                    $event = $this->BuildEventFromLines($currentLines);
                    if ($event !== null) {
                        $events[] = $event;
                    }
                }
                $inEvent = false;
                $currentLines = [];
                continue;
            }
            if ($inEvent) {
                $currentLines[] = $line;
            }
        }

        return $events;
    }

    private function BuildEventFromLines(array $lines): ?array
    {
        $start = null;
        $end = null;
        $summary = '';
        foreach ($lines as $line) {
            $upper = strtoupper($line);
            if (str_starts_with($upper, 'DTSTART')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $start = $this->ParseICSTime($meta, $value);
                continue;
            }

            if (str_starts_with($upper, 'DTEND')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $end = $this->ParseICSTime($meta, $value);
                continue;
            }

            if (str_starts_with($upper, 'SUMMARY')) {
                [, $value] = $this->SplitMetaValue($line);
                $summary = $this->UnescapeICSText($value);
            }
        }

        if ($start === null || $end === null) {
            return null;
        }

        if ($end <= $start) {
            return null;
        }

        return [
            'start' => $start,
            'end'   => $end,
            'summary' => $summary ?? ''
        ];
    }

    public function GetConfigurationForm()
    {
        $formPath = __DIR__ . '/form.json';
        $form = [];

        if (is_file($formPath)) {
            $content = file_get_contents($formPath);
            if ($content !== false) {
                $decoded = json_decode($content, true);
                if (is_array($decoded)) {
                    $form = $decoded;
                }
            }
        }

        if (!isset($form['actions']) || !is_array($form['actions'])) {
            $form['actions'] = [];
        }

        foreach ($form['actions'] as $index => $action) {
            if (($action['name'] ?? '') === 'UpcomingEventsPlaceholder') {
                $form['actions'][$index] = $this->BuildUpcomingEventsPreviewAction();
                break;
            }
        }

        $json = json_encode($form);
        if ($json === false) {
            $this->Log('Failed to encode configuration form.');
            return '{}';
        }

        return $json;
    }

    private function BuildUpcomingEventsPreviewAction(): array
    {
        if ($this->SupportsHtmlBoxInConfigurationForm()) {
            return [
                'type'    => 'HTMLBox',
                'caption' => $this->Translate('Upcoming events'),
                'html'    => '{{GetValueString(IPS_GetObjectIDByIdent("UpcomingEventsHTML", $id))}}'
            ];
        }

        return [
            'type'    => 'Label',
            'caption' => $this->BuildUpcomingEventsPreviewText()
        ];
    }

    private function SupportsHtmlBoxInConfigurationForm(): bool
    {
        $version = IPS_GetKernelVersion();
        if (!is_string($version)) {
            return false;
        }

        if (!preg_match('/^(\d+)\.(\d+)/', $version, $matches)) {
            $this->Log('Unable to determine HTMLBox support because kernel version "' . $version . '" does not match the expected pattern.');
            return false;
        }

        $major = (int) $matches[1];
        $minor = (int) $matches[2];

        if ($major > 7) {
            return true;
        }

        if ($major === 7 && $minor >= 1) {
            return true;
        }

        return false;
    }

    private function BuildUpcomingEventsPreviewText(): string
    {
        $intro = $this->Translate('Upcoming events preview:');
        $variableID = @IPS_GetObjectIDByIdent('UpcomingEventsHTML', $this->InstanceID);
        if ($variableID === false) {
            return $intro . PHP_EOL . $this->Translate('No upcoming events available.');
        }

        $html = GetValueString($variableID);
        $lines = $this->ExtractUpcomingEventLinesFromHtml($html);
        if (empty($lines)) {
            $lines[] = $this->Translate('No upcoming events');
        }

        return $intro . PHP_EOL . implode(PHP_EOL, $lines);
    }

    private function ExtractUpcomingEventLinesFromHtml(string $html): array
    {
        if ($html === '') {
            return [];
        }

        if (!preg_match_all('/<tr>(.*?)<\/tr>/is', $html, $rowMatches)) {
            return [];
        }

        $rows = [];
        foreach ($rowMatches[1] as $rowContent) {
            if (!preg_match_all('/<(?:td|th)[^>]*>(.*?)<\/(?:td|th)>/is', $rowContent, $cellMatches)) {
                continue;
            }

            $cells = [];
            foreach ($cellMatches[1] as $cellContent) {
                $text = trim(html_entity_decode(strip_tags($cellContent), ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8'));
                if ($text !== '') {
                    $cells[] = $text;
                }
            }

            if (!empty($cells)) {
                $rows[] = $cells;
            }
        }

        if (count($rows) <= 1) {
            return [];
        }

        $bodyRows = array_slice($rows, 1);
        $lines = [];
        foreach ($bodyRows as $cells) {
            if (count($cells) === 1) {
                $lines[] = $cells[0];
                continue;
            }

            $summary = $cells[0] ?? '';
            $start = $cells[1] ?? '';
            $status = $cells[2] ?? '';

            $parts = array_filter([$summary, $start], static function (string $value): bool {
                return $value !== '';
            });

            if (empty($parts)) {
                continue;
            }

            $line = implode(' — ', $parts);
            if ($status !== '') {
                $line .= ' (' . $status . ')';
            }

            $lines[] = $line;
        }

        return $lines;
    }

    private function SplitMetaValue(string $line): array
    {
        $parts = explode(':', $line, 2);
        $meta = $parts[0] ?? '';
        $value = $parts[1] ?? '';
        return [$meta, $value];
    }

    private function ParseICSTime(string $meta, string $value): ?int
    {
        $value = trim($value);
        if ($value === '') {
            return null;
        }

        if (str_contains($meta, 'VALUE=DATE')) {
            return null;
        }

        $timezone = null;
        $metaParts = explode(';', $meta);
        array_shift($metaParts); // remove DTSTART or DTEND
        foreach ($metaParts as $part) {
            $part = trim($part);
            if (stripos($part, 'TZID=') === 0) {
                $timezone = substr($part, 5);
            }
        }

        $valueNoZ = $value;
        $tz = null;
        if (str_ends_with($valueNoZ, 'Z')) {
            $valueNoZ = substr($valueNoZ, 0, -1);
            $tz = new DateTimeZone('UTC');
        } elseif ($timezone !== null && $timezone !== '') {
            try {
                $tz = new DateTimeZone($timezone);
            } catch (\Throwable $e) {
                $this->Log('Unknown timezone in calendar entry: ' . $timezone);
                $tz = null;
            }
        }

        if ($tz === null) {
            $tz = new DateTimeZone(date_default_timezone_get());
        }

        $formats = ['Ymd\THis', 'Ymd\THi', 'Ymd\TH', 'Ymd'];
        $dateTime = null;
        foreach ($formats as $format) {
            $dateTime = DateTime::createFromFormat($format, $valueNoZ, $tz);
            if ($dateTime instanceof DateTime) {
                break;
            }
        }

        if (!$dateTime instanceof DateTime) {
            return null;
        }

        if (strlen($valueNoZ) === 8) {
            return null;
        }

        if (str_ends_with($value, 'Z')) {
            $dateTime = new DateTime($dateTime->format('c'));
            $dateTime->setTimezone(new DateTimeZone('UTC'));
        }

        return $dateTime->getTimestamp();
    }

    private function GetStoredEvent(): ?array
    {
        $start = $this->ReadAttributeInteger('LastEventStart');
        $end = $this->ReadAttributeInteger('LastEventEnd');
        if ($start > 0 && $end > 0 && $end > $start) {
            $summary = $this->ReadAttributeString('LastEventSummary');
            return ['start' => $start, 'end' => $end, 'summary' => $summary];
        }
        return null;
    }

    private function GetStoredEventForFallback(int $now): ?array
    {
        $event = $this->GetStoredEvent();
        if ($event !== null) {
            if ($event['end'] > $now) {
                return $event;
            }
        }
        return null;
    }

    private function BuildUpcomingEventsTable(array $events, int $now, float $setpoint, float $heatingRate, ?float $currentTemp, int $bufferSeconds): string
    {
        $headerLabels = [
            $this->Translate('Event Name'),
            $this->Translate('Event Start'),
            $this->Translate('Heating Status')
        ];

        $header = '<tr>';
        foreach ($headerLabels as $label) {
            $header .= sprintf(
                "<td style='padding: 5px; font-weight: bold;'>%s</td>",
                htmlspecialchars($label, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8')
            );
        }
        $header .= '</tr>';

        $rows = '';
        foreach ($events as $event) {
            if (!isset($event['start'], $event['end'])) {
                continue;
            }

            $summary = trim((string) ($event['summary'] ?? ''));
            if ($summary === '') {
                $summary = $this->Translate('Unnamed event');
            }

            $summaryHtml = nl2br(htmlspecialchars($summary, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8'));
            $eventStart = (int) $event['start'];
            $eventEnd = (int) $event['end'];
            $eventStartHtml = htmlspecialchars(date('Y-m-d H:i', $eventStart), ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');

            $preheatStart = $this->CalculatePreheatStart($eventStart, $setpoint, $heatingRate, $currentTemp, $bufferSeconds);

            $status = '';
            if ($now >= $eventStart && $now < $eventEnd) {
                $status = $this->Translate('Event running');
            } elseif ($now >= $preheatStart && $now < $eventStart) {
                $status = $this->Translate('Preheating');
            } elseif ($now < $preheatStart) {
                $status = sprintf($this->Translate('Starts in %s'), $this->FormatInterval($preheatStart - $now));
            } else {
                $status = $this->Translate('Completed');
            }

            $statusHtml = htmlspecialchars($status, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');

            $rows .= sprintf(
                "<tr><td style='padding: 5px;'>%s</td><td style='padding: 5px;'>%s</td><td style='padding: 5px;'>%s</td></tr>",
                $summaryHtml,
                $eventStartHtml,
                $statusHtml
            );
        }

        if ($rows === '') {
            $rows = sprintf(
                "<tr><td colspan='3' style='padding: 5px;'>%s</td></tr>",
                htmlspecialchars($this->Translate('No upcoming events'), ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8')
            );
        }

        return "<table style='width: 100%; border-collapse: collapse;'>" . $header . $rows . '</table>';
    }

    private function FormatInterval(int $seconds): string
    {
        if ($seconds <= 0) {
            return '0s';
        }

        $units = [
            ['suffix' => 'd', 'seconds' => 86400],
            ['suffix' => 'h', 'seconds' => 3600],
            ['suffix' => 'm', 'seconds' => 60],
            ['suffix' => 's', 'seconds' => 1]
        ];

        $parts = [];
        foreach ($units as $unit) {
            if ($seconds >= $unit['seconds']) {
                $value = intdiv($seconds, $unit['seconds']);
                $seconds -= $value * $unit['seconds'];
                $parts[] = sprintf('%d%s', $value, $unit['suffix']);
            }

            if (count($parts) === 2) {
                break;
            }
        }

        if (empty($parts)) {
            $parts[] = '0s';
        }

        return implode(' ', $parts);
    }

    private function CalculatePreheatStart(int $eventStart, float $setpoint, float $heatingRate, ?float $currentTemp, int $bufferSeconds): int
    {
        $preheatStart = max(0, $eventStart - $bufferSeconds);

        if ($currentTemp !== null && $heatingRate > 0.0) {
            $delta = $setpoint - $currentTemp;
            if ($delta < 0.0) {
                $delta = 0.0;
            }

            $preheatDurationHours = $delta / $heatingRate;
            $preheatSeconds = (int) round($preheatDurationHours * 3600);
            $preheatStart = $eventStart - $preheatSeconds - $bufferSeconds;
        }

        if ($preheatStart < 0) {
            $preheatStart = 0;
        }

        return $preheatStart;
    }

    private function UnescapeICSText(string $value): string
    {
        $value = str_replace(
            ['\\n', '\\N', '\\,', '\\;', '\\\\'],
            ["\n", "\n", ',', ';', '\\'],
            $value
        );

        return trim($value);
    }

    private function Log(string $message): void
    {
        IPS_LogMessage('PreheatScheduler', sprintf('#%d %s', $this->InstanceID, $message));
    }
}
