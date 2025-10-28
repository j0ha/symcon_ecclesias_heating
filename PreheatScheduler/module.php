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

        $heatingVarID = $this->RegisterVariableBoolean('HeatingDemand', $this->Translate('Heating Demand'));
        IPS_SetVariableCustomProfile($heatingVarID, '~Switch');
        $nextEventVarID = $this->RegisterVariableString('NextEventStartISO', $this->Translate('Next Event Start'));
        IPS_SetVariableCustomProfile($nextEventVarID, '~String');
        $nextPreheatVarID = $this->RegisterVariableString('NextPreheatStartISO', $this->Translate('Next Preheat Start'));
        IPS_SetVariableCustomProfile($nextPreheatVarID, '~String');
        $overviewVarID = $this->RegisterVariableString('EventOverviewHTML', $this->Translate('Event Overview'));
        IPS_SetVariableCustomProfile($overviewVarID, '~HTMLBox');

        $this->SetValue('HeatingDemand', false);
        $this->SetValue('NextEventStartISO', '-');
        $this->SetValue('NextPreheatStartISO', '-');
        $this->SetValue('EventOverviewHTML', $this->BuildEventOverviewTable([], time(), $this->Translate('Keine Veranstaltungen verfügbar.')));

        $this->RegisterTimer('Evaluate', 0, 'HEAT_Recalculate($_IPS[\'TARGET\']);');

        $this->RegisterAttributeInteger('RegisteredTempVarID', 0);
        $this->RegisterAttributeInteger('LastEventStart', 0);
        $this->RegisterAttributeInteger('LastEventEnd', 0);
        $this->RegisterAttributeInteger('LastPreheatStart', 0);
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
        $event = $this->DetermineNextEvent($now);
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

            $preheatStart = max(0, $eventStart - $bufferSeconds);

            if ($currentTemp !== null && $heatingRate > 0.0) {
                $delta = $setpoint - $currentTemp;
                if ($delta < 0.0) {
                    $delta = 0.0;
                }
                $preheatDurationHours = $delta / $heatingRate;
                $preheatSeconds = (int) round($preheatDurationHours * 3600);
                $preheatStart = $eventStart - $preheatSeconds - $bufferSeconds;
            } elseif ($heatingRate <= 0.0) {
                $this->Log('Unable to calculate preheat start because heating rate is zero or negative.');
            }

            if ($preheatStart < 0) {
                $preheatStart = 0;
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

    private function DetermineNextEvent(int $now): ?array
    {
        $calendarUrl = trim($this->ReadPropertyString('CalendarURL'));
        if ($calendarUrl === '') {
            $this->Log('Calendar URL is not configured.');
            $this->SetStatus(self::STATUS_URL_ERROR);
            $this->UpdateEventOverview([], $now, $this->Translate('Kalender ist nicht konfiguriert.'));
            return $this->GetStoredEventForFallback($now);
        }

        $content = $this->FetchCalendarContent($calendarUrl);
        if ($content === null) {
            $this->UpdateEventOverview([], $now, $this->Translate('Kalender konnte nicht geladen werden.'));
            return $this->GetStoredEventForFallback($now);
        }

        $events = $this->ParseICSEvents($content);
        if (empty($events)) {
            $this->Log('No events found in calendar export.');
            $this->WriteAttributeInteger('LastEventStart', 0);
            $this->WriteAttributeInteger('LastEventEnd', 0);
            $this->WriteAttributeInteger('LastPreheatStart', 0);
            $this->SetStatus(self::STATUS_OK);
            $this->UpdateEventOverview([], $now, $this->Translate('Keine Veranstaltungen gefunden.'));
            return null;
        }

        $lookaheadSeconds = max(1, $this->ReadPropertyInteger('LookaheadHours')) * 3600;
        $horizon = $now + $lookaheadSeconds;

        $nextEvent = null;
        $upcomingEvents = [];
        foreach ($events as $event) {
            if (!isset($event['start'], $event['end'])) {
                continue;
            }
            if ($event['end'] <= $now) {
                continue;
            }
            if ($event['start'] <= $now && $event['end'] > $now) {
                if ($nextEvent === null || $event['start'] < $nextEvent['start']) {
                    $nextEvent = $event;
                }
                $upcomingEvents[] = $event;
                continue;
            }
            if ($event['start'] > $horizon) {
                continue;
            }
            if ($event['start'] >= $now) {
                if ($nextEvent === null || $event['start'] < $nextEvent['start']) {
                    $nextEvent = $event;
                }
                $upcomingEvents[] = $event;
            }
        }

        usort($upcomingEvents, static fn ($a, $b) => $a['start'] <=> $b['start']);
        $this->UpdateEventOverview($upcomingEvents, $now, null);

        if ($nextEvent !== null) {
            $this->WriteAttributeInteger('LastEventStart', $nextEvent['start']);
            $this->WriteAttributeInteger('LastEventEnd', $nextEvent['end']);
            $this->WriteAttributeInteger('LastPreheatStart', 0);
            $this->SetStatus(self::STATUS_OK);
        } else {
            $this->SetStatus(self::STATUS_OK);
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
        $summary = null;

        foreach ($lines as $line) {
            $upper = strtoupper($line);
            if (str_starts_with($upper, 'DTSTART')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $start = $this->ParseICSTime($meta, $value);
            } elseif (str_starts_with($upper, 'DTEND')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $end = $this->ParseICSTime($meta, $value);
            } elseif (str_starts_with($upper, 'SUMMARY')) {
                [, $value] = $this->SplitMetaValue($line);
                $summary = trim($value);
            } elseif ($start !== null && $end !== null) {
                break;
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
            return ['start' => $start, 'end' => $end];
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

    private function Log(string $message): void
    {
        IPS_LogMessage('PreheatScheduler', sprintf('#%d %s', $this->InstanceID, $message));
    }

    private function UpdateEventOverview(array $events, int $now, ?string $errorMessage): void
    {
        $html = $this->BuildEventOverviewTable($events, $now, $errorMessage);
        $this->SetValue('EventOverviewHTML', $html);
    }

    private function BuildEventOverviewTable(array $events, int $now, ?string $errorMessage): string
    {
        $table = "<table style='width: 100%; border-collapse: collapse;'><tr><td style='padding: 5px; font-weight: bold;'>Event Name</td><td style='padding: 5px; font-weight: bold;'>Event Start</td><td style='padding: 5px; font-weight: bold;'>Status</td></tr>";

        if ($errorMessage !== null) {
            $table .= "<tr><td colspan='3' style='padding: 5px;'>" . htmlspecialchars($errorMessage) . '</td></tr>';
            return $table . '</table>';
        }

        $tempVarID = $this->ReadPropertyInteger('TempVarID');
        $currentTemp = null;
        if ($tempVarID > 0 && IPS_VariableExists($tempVarID)) {
            $currentTemp = (float) GetValue($tempVarID);
        }

        $setpoint = $this->ReadPropertyFloat('SetpointWarm');
        $heatingRate = $this->ReadPropertyFloat('HeatingRate');
        $bufferSeconds = max(0, $this->ReadPropertyInteger('PreheatBufferMin')) * 60;

        $rowsAdded = false;
        foreach ($events as $event) {
            if (!isset($event['start'], $event['end'])) {
                continue;
            }
            if ($event['end'] <= $now) {
                continue;
            }
            $rowsAdded = true;
            $eventStart = (int) $event['start'];
            $eventEnd = (int) $event['end'];
            $summary = trim((string) ($event['summary'] ?? ''));
            if ($summary === '') {
                $summary = $this->Translate('Unbenannte Veranstaltung');
            }

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

            $status = '';
            if ($now >= $eventStart && $now < $eventEnd) {
                $status = 'Veranstaltung läuft // Temperatur wird gehalten';
            } elseif ($now >= $preheatStart && $now < $eventStart) {
                $status = 'Vorheizen Aktiv';
            } elseif ($now < $preheatStart) {
                $diff = $preheatStart - $now;
                $status = $this->FormatCountdown($diff);
            } else {
                $status = '-';
            }

            $table .= '<tr>'
                . "<td style='padding: 5px;'>" . htmlspecialchars($summary) . '</td>'
                . "<td style='padding: 5px;'>" . htmlspecialchars($this->FormatEventDate($eventStart)) . '</td>'
                . "<td style='padding: 5px;'>" . htmlspecialchars($status) . '</td>'
                . '</tr>';
        }

        if (!$rowsAdded) {
            $table .= "<tr><td colspan='3' style='padding: 5px;'>" . $this->Translate('Keine zukünftigen Veranstaltungen im Suchzeitraum.') . '</td></tr>';
        }

        return $table . '</table>';
    }

    private function FormatEventDate(int $timestamp): string
    {
        $days = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
        $months = [
            1 => 'Januar',
            2 => 'Februar',
            3 => 'März',
            4 => 'April',
            5 => 'Mai',
            6 => 'Juni',
            7 => 'Juli',
            8 => 'August',
            9 => 'September',
            10 => 'Oktober',
            11 => 'November',
            12 => 'Dezember'
        ];

        $dayName = $days[(int) date('w', $timestamp)];
        $day = (int) date('j', $timestamp);
        $month = $months[(int) date('n', $timestamp)] ?? '';
        $hour = (int) date('G', $timestamp);
        $minute = (int) date('i', $timestamp);

        return sprintf('%s %d. %s %02d:%02d', $dayName, $day, $month, $hour, $minute);
    }

    private function FormatCountdown(int $seconds): string
    {
        if ($seconds < 0) {
            $seconds = 0;
        }
        $hours = intdiv($seconds, 3600);
        $minutes = intdiv($seconds % 3600, 60);
        return sprintf('Heizstart in %d Std %d Min', $hours, $minutes);
    }
}
