<?php

declare(strict_types=1);

class PreheatScheduler extends IPSModule
{
    private const STATUS_OK = 102;
    private const STATUS_URL_ERROR = 201;
    private const STATUS_AUTH_ERROR = 202;

    private bool $legacyBlacklistWarningIssued = false;

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
        $this->RegisterPropertyString('EventBlacklist', '[]');

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
        $this->RegisterAttributeInteger('DemandHoldUntil', 0);
    }

    public function ApplyChanges()
    {
        parent::ApplyChanges();

        $this->Debug('ApplyChanges', 'Starting apply changes');

        $this->invalidBlacklistPatterns = [];

        $interval = max(15, $this->ReadPropertyInteger('EvaluationIntervalSec'));
        $this->SetTimerInterval('Evaluate', $interval * 1000);

        $this->Debug('ApplyChanges', sprintf('Timer interval set to %d seconds', $interval));

        $tempVarID = $this->ReadPropertyInteger('TempVarID');
        $lastRegistered = $this->ReadAttributeInteger('RegisteredTempVarID');

        $this->Debug('ApplyChanges', sprintf('Temperature variable configured: %d (last registered %d)', $tempVarID, $lastRegistered));

        if ($lastRegistered > 0 && $lastRegistered !== $tempVarID) {
            $this->Debug('ApplyChanges', sprintf('Unregistering previous temp var ID %d', $lastRegistered));
            $this->UnregisterMessage($lastRegistered, VM_UPDATE);
            $this->WriteAttributeInteger('RegisteredTempVarID', 0);
        }

        if ($tempVarID > 0 && IPS_VariableExists($tempVarID)) {
            $this->Debug('ApplyChanges', sprintf('Registering update message for temp var ID %d', $tempVarID));
            $this->RegisterMessage($tempVarID, VM_UPDATE);
            $this->WriteAttributeInteger('RegisteredTempVarID', $tempVarID);
        } else {
            $this->Debug('ApplyChanges', sprintf('Temperature variable %d not registered or missing', $tempVarID));
        }

        $this->Recalculate();
    }

    public function MessageSink($TimeStamp, $SenderID, $Message, $Data)
    {
        $this->Debug('MessageSink', sprintf('Message received: sender=%d message=%d timestamp=%d', $SenderID, $Message, $TimeStamp));
        if ($Message === VM_UPDATE) {
            $tempVarID = $this->ReadPropertyInteger('TempVarID');
            if ($SenderID === $tempVarID) {
                $this->Debug('MessageSink', 'Temperature variable update detected, recalculating');
                $this->Recalculate();
            }
        }
    }

    public function Recalculate(): bool
    {
        $this->Debug('Recalculate', 'Recalculation started');
        $now = time();
        $this->Debug('Recalculate', sprintf('Current timestamp: %d', $now));
        $event = $this->DetermineNextEvent($now);
        $this->Debug('Recalculate', $event === null ? 'No upcoming event detected' : 'Upcoming event detected');

        $holdStrategy = $this->ReadPropertyInteger('HoldStrategy');
        $this->Debug('Recalculate', sprintf('Hold strategy: %d', $holdStrategy));

        $tempVarID = $this->ReadPropertyInteger('TempVarID');
        $currentTemp = null;
        if ($tempVarID > 0 && IPS_VariableExists($tempVarID)) {
            $currentTemp = (float) GetValue($tempVarID);
            $this->Debug('Recalculate', sprintf('Current temperature: %.2f', $currentTemp));
        }

        if ($currentTemp === null) {
            $this->Log('Temperature variable not set or missing.');
            $this->Debug('Recalculate', 'Temperature variable not set or missing');
        }

        $setpoint = $this->ReadPropertyFloat('SetpointWarm');
        $heatingRate = $this->ReadPropertyFloat('HeatingRate');
        $this->Debug('Recalculate', sprintf('Setpoint: %.2f Heating rate: %.2f', $setpoint, $heatingRate));
        if ($heatingRate <= 0.0) {
            $this->Log('Heating rate must be greater than zero.');
            $this->Debug('Recalculate', 'Heating rate must be greater than zero');
        }

        $bufferSeconds = max(0, $this->ReadPropertyInteger('PreheatBufferMin')) * 60;
        $this->Debug('Recalculate', sprintf('Preheat buffer (seconds): %d', $bufferSeconds));

        $heatingVarID = $this->GetIDForIdent('HeatingDemand');
        $currentlyOn = GetValueBoolean($heatingVarID);
        $demandHoldUntil = $this->ReadAttributeInteger('DemandHoldUntil');

        $this->Debug('Recalculate', sprintf('Heating currently on: %s, demand hold until: %d', $currentlyOn ? 'true' : 'false', $demandHoldUntil));

        $shouldBeOn = false;
        $eventStartISO = '-';
        $preheatStartISO = '-';
        $activeEventEnd = null;

        if ($event !== null) {
            $eventStart = $event['start'];
            $eventEnd = $event['end'];
            $eventStartISO = date('c', $eventStart);
            $activeEventEnd = $eventEnd;

            $this->Debug('Recalculate', sprintf('Processing event: start=%d end=%d', $eventStart, $eventEnd));

            $preheatStart = max(0, $eventStart - $bufferSeconds);

            if ($currentTemp !== null && $heatingRate > 0.0) {
                $delta = $setpoint - $currentTemp;
                if ($delta < 0.0) {
                    $delta = 0.0;
                }
                $preheatDurationHours = $delta / $heatingRate;
                $preheatSeconds = (int) round($preheatDurationHours * 3600);
                $preheatStart = $eventStart - $preheatSeconds - $bufferSeconds;
                $this->Debug('Recalculate', sprintf('Calculated preheat start: %d (delta=%.2f, duration=%d)', $preheatStart, $delta, $preheatSeconds));
            } elseif ($heatingRate <= 0.0) {
                $this->Log('Unable to calculate preheat start because heating rate is zero or negative.');
                $this->Debug('Recalculate', 'Unable to calculate preheat start due to heating rate');
            }

            if ($preheatStart < 0) {
                $preheatStart = 0;
            }

            $this->WriteAttributeInteger('LastPreheatStart', $preheatStart);
            $preheatStartISO = $preheatStart > 0 ? date('c', $preheatStart) : '-';

            $this->Debug('Recalculate', sprintf('Preheat start ISO: %s', $preheatStartISO));

            $windowEnd = $holdStrategy === 0 ? $eventEnd : $eventStart;
            if ($windowEnd < $eventStart) {
                $windowEnd = $eventStart;
            }

            if ($now >= $preheatStart && $now < $windowEnd) {
                $shouldBeOn = true;
                $this->Debug('Recalculate', 'Heating should be on due to preheat window');
            }

            if ($now >= $eventStart && $now < $eventEnd) {
                $shouldBeOn = true;
                $this->Debug('Recalculate', 'Heating should be on due to active event');
            }

            if (!$shouldBeOn && $currentlyOn) {
                if ($holdStrategy === 0 && $now >= $eventStart && $now < $eventEnd) {
                    $shouldBeOn = true;
                    $this->Debug('Recalculate', 'Maintaining heating due to hold strategy');
                }
            }
        } else {
            $stored = $this->GetStoredEvent();
            if ($stored !== null) {
                $activeEventEnd = $stored['end'];
                if ($now < $stored['end']) {
                    if ($currentlyOn) {
                        $shouldBeOn = true;
                        $this->Debug('Recalculate', 'Continuing heating due to stored event');
                    }
                }
            }
        }

        if ($currentlyOn && !$shouldBeOn && $demandHoldUntil > $now) {
            $shouldBeOn = true;
            $this->Debug('Recalculate', 'Demand hold is keeping heating on');
        }

        $this->SetValue('NextEventStartISO', $eventStartISO);
        $this->SetValue('NextPreheatStartISO', $preheatStartISO);

        if ($shouldBeOn !== $currentlyOn) {
            $this->SetValue('HeatingDemand', $shouldBeOn);
            $this->Debug('Recalculate', sprintf('Heating demand toggled to %s', $shouldBeOn ? 'true' : 'false'));
        }

        if ($shouldBeOn) {
            if ($activeEventEnd !== null && $activeEventEnd !== $demandHoldUntil) {
                $this->WriteAttributeInteger('DemandHoldUntil', $activeEventEnd);
                $this->Debug('Recalculate', sprintf('Demand hold updated to %d', $activeEventEnd));
            }
        } else {
            if ($demandHoldUntil !== 0 && $demandHoldUntil <= $now) {
                $this->WriteAttributeInteger('DemandHoldUntil', 0);
                $this->Debug('Recalculate', 'Demand hold reset');
            }
        }

        $this->Debug('Recalculate', sprintf('Recalculation finished, heating demand: %s', $shouldBeOn ? 'true' : 'false'));
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
            $this->Debug('DetermineNextEvent', 'Calendar URL missing');
            $this->SetStatus(self::STATUS_URL_ERROR);
            $this->UpdateEventOverview([], $now, $this->Translate('Kalender ist nicht konfiguriert.'));
            return $this->GetStoredEventForFallback($now);
        }

        $content = $this->FetchCalendarContent($calendarUrl);
        if ($content === null) {
            $this->Debug('DetermineNextEvent', 'Calendar content fetch failed');
            $this->UpdateEventOverview([], $now, $this->Translate('Kalender konnte nicht geladen werden.'));
            return $this->GetStoredEventForFallback($now);
        }

        $events = $this->ParseICSEvents($content);
        $this->Debug('DetermineNextEvent', sprintf('Parsed %d events', count($events)));
        $events = $this->FilterBlacklistedEvents($events);
        $this->Debug('DetermineNextEvent', sprintf('Events after blacklist filter: %d', count($events)));
        if (empty($events)) {
            $this->Log('No events found in calendar export.');
            $this->Debug('DetermineNextEvent', 'No events after filtering');
            $this->WriteAttributeInteger('LastEventStart', 0);
            $this->WriteAttributeInteger('LastEventEnd', 0);
            $this->WriteAttributeInteger('LastPreheatStart', 0);
            $this->SetStatus(self::STATUS_OK);
            $this->UpdateEventOverview([], $now, $this->Translate('Keine Veranstaltungen gefunden.'));
            return null;
        }

        $lookaheadSeconds = max(1, $this->ReadPropertyInteger('LookaheadHours')) * 3600;
        $horizon = $now + $lookaheadSeconds;

        $events = $this->ExpandRecurringEvents($events, $now, $horizon);
        $this->Debug('DetermineNextEvent', sprintf('Events after expansion: %d', count($events)));

        $nextEvent = null;
        $upcomingEvents = [];
        foreach ($events as $event) {
            if (!isset($event['start'], $event['end'])) {
                continue;
            }
            if ($event['end'] <= $now) {
                continue;
            }
            $isCancelled = isset($event['status']) && strtoupper((string) $event['status']) === 'CANCELLED';
            if ($isCancelled) {
                if ($event['start'] <= $horizon) {
                    $upcomingEvents[] = $event;
                }
                continue;
            }
            if ($event['start'] <= $now && $event['end'] > $now) {
                if ($nextEvent === null || $event['start'] < $nextEvent['start']) {
                    $nextEvent = $event;
                }
                $upcomingEvents[] = $event;
                $this->Debug('DetermineNextEvent', sprintf('Active event detected: start=%d end=%d', $event['start'], $event['end']));
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
                $this->Debug('DetermineNextEvent', sprintf('Upcoming event queued: start=%d end=%d', $event['start'], $event['end']));
            }
        }

        usort($upcomingEvents, static fn ($a, $b) => $a['start'] <=> $b['start']);
        $this->UpdateEventOverview($upcomingEvents, $now, null);

        if ($nextEvent !== null) {
            $this->WriteAttributeInteger('LastEventStart', $nextEvent['start']);
            $this->WriteAttributeInteger('LastEventEnd', $nextEvent['end']);
            $this->WriteAttributeInteger('LastPreheatStart', 0);
            $this->SetStatus(self::STATUS_OK);
            $this->Debug('DetermineNextEvent', sprintf('Next event selected: start=%d end=%d', $nextEvent['start'], $nextEvent['end']));
        } else {
            $this->SetStatus(self::STATUS_OK);
            $this->Debug('DetermineNextEvent', 'No next event found');
        }

        return $nextEvent;
    }

    private function FetchCalendarContent(string $calendarUrl): ?string
    {
        $user = $this->ReadPropertyString('CalUser');
        $pass = $this->ReadPropertyString('CalPass');

        $this->Debug('FetchCalendar', sprintf('Fetching calendar from %s', $calendarUrl));

        $urlsToTry = [];
        $trimmed = rtrim($calendarUrl);
        if (!preg_match('/\.ics($|\?)/i', $trimmed) && !str_contains($trimmed, '?export')) {
            $separator = str_contains($trimmed, '?') ? '&' : '?';
            $urlsToTry[] = $trimmed . $separator . 'export';
            $this->Debug('FetchCalendar', sprintf('Added export helper URL: %s', $trimmed . $separator . 'export'));
        }
        $urlsToTry[] = $trimmed;

        $auth = [];
        if ($user !== '' || $pass !== '') {
            $auth['AuthUser'] = $user;
            $auth['AuthPass'] = $pass;
            $this->Debug('FetchCalendar', 'Authentication configured for calendar fetch');
        }

        $lastError = '';
        foreach ($urlsToTry as $url) {
            error_clear_last();
            $content = @Sys_GetURLContentEx($url, $auth);
            if ($content !== false && $content !== null) {
                if ($url !== $trimmed) {
                    $this->Log('Calendar fetched using export helper URL: ' . $url);
                    $this->Debug('FetchCalendar', sprintf('Calendar fetched via helper URL: %s', $url));
                }
                $this->SetStatus(self::STATUS_OK);
                $this->Debug('FetchCalendar', 'Calendar content fetched successfully');
                return $content;
            }
            $error = error_get_last();
            if ($error !== null) {
                $lastError = $error['message'] ?? '';
                $this->Debug('FetchCalendar', sprintf('Fetch attempt failed for %s: %s', $url, $lastError));
            }
        }

        if ($lastError !== '') {
            $this->Log('Calendar fetch failed: ' . $lastError);
            $this->Debug('FetchCalendar', sprintf('Calendar fetch failed with error: %s', $lastError));
        } else {
            $this->Log('Calendar fetch failed: unknown error');
            $this->Debug('FetchCalendar', 'Calendar fetch failed with unknown error');
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
        $this->Debug('ParseICSEvents', 'Parsing ICS content');
        $lines = preg_split('/\r\n|\n|\r/', $content);
        if ($lines === false) {
            $this->Debug('ParseICSEvents', 'Failed to split ICS lines');
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
                        $this->Debug('ParseICSEvents', sprintf('Event parsed: start=%s end=%s', $event['start'] ?? 'n/a', $event['end'] ?? 'n/a'));
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

        $this->Debug('ParseICSEvents', sprintf('Total events parsed: %d', count($events)));

        return $events;
    }

    private function BuildEventFromLines(array $lines): ?array
    {
        $start = null;
        $end = null;
        $summary = '';
        $status = '';
        $timezone = null;
        $rrule = '';
        $rdates = [];
        $exdates = [];
        $uid = '';
        $recurrenceId = null;

        foreach ($lines as $line) {
            $upper = strtoupper($line);
            if (str_starts_with($upper, 'DTSTART')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $start = $this->ParseICSTime($meta, $value, $timezone);
                continue;
            }

            if (str_starts_with($upper, 'DTEND')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $end = $this->ParseICSTime($meta, $value);
                continue;
            }

            if (str_starts_with($upper, 'SUMMARY')) {
                [, $value] = $this->SplitMetaValue($line);
                $summary = trim($this->UnescapeICSText($value));
                continue;
            }

            if (str_starts_with($upper, 'STATUS')) {
                [, $value] = $this->SplitMetaValue($line);
                $status = strtoupper(trim($this->UnescapeICSText($value)));
                continue;
            }

            if (str_starts_with($upper, 'UID')) {
                [, $value] = $this->SplitMetaValue($line);
                $uid = trim($this->UnescapeICSText($value));
                continue;
            }

            if (str_starts_with($upper, 'RECURRENCE-ID')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $recurrenceId = $this->ParseICSTime($meta, $value);
                continue;
            }

            if (str_starts_with($upper, 'RRULE')) {
                [, $value] = $this->SplitMetaValue($line);
                $rrule = trim($value);
                continue;
            }

            if (str_starts_with($upper, 'RDATE')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $rdates = array_merge($rdates, $this->ParseICSMultipleTimes($meta, $value));
                continue;
            }

            if (str_starts_with($upper, 'EXDATE')) {
                [$meta, $value] = $this->SplitMetaValue($line);
                $exdates = array_merge($exdates, $this->ParseICSMultipleTimes($meta, $value));
                continue;
            }
        }

        if ($start === null || $end === null) {
            $this->Debug('BuildEventFromLines', 'Missing start or end in event definition');
            return null;
        }

        if ($end <= $start) {
            $this->Debug('BuildEventFromLines', sprintf('Invalid event duration start=%d end=%d', $start, $end));
            return null;
        }

        return [
            'start' => $start,
            'end'   => $end,
            'summary' => $summary,
            'status' => $status,
            'timezone' => $timezone,
            'rrule' => $rrule,
            'rdates' => $rdates,
            'exdates' => $exdates,
            'uid' => $uid,
            'recurrence_id' => $recurrenceId
        ];
    }

    private function SplitMetaValue(string $line): array
    {
        $parts = explode(':', $line, 2);
        $meta = $parts[0] ?? '';
        $value = $parts[1] ?? '';
        return [$meta, $value];
    }

    private function UnescapeICSText(string $value): string
    {
        $value = str_replace(['\\n', '\\N'], "\n", $value);
        $value = str_replace('\\,', ',', $value);
        $value = str_replace('\\;', ';', $value);
        return str_replace('\\\\', '\\', $value);
    }

    private function ParseICSMultipleTimes(string $meta, string $value): array
    {
        $timestamps = [];
        $parts = explode(',', $value);
        foreach ($parts as $part) {
            $part = trim($part);
            if ($part === '') {
                continue;
            }
            $timestamp = $this->ParseICSTime($meta, $part);
            if ($timestamp !== null) {
                $timestamps[] = $timestamp;
            }
        }

        return $timestamps;
    }

    private function ParseICSTime(string $meta, string $value, ?string &$timezoneName = null): ?int
    {
        $timezoneName = null;
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
        $timezoneName = $tz->getName();

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
            $timezoneName = 'UTC';
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
        $formatted = sprintf('#%d %s', $this->InstanceID, $message);
        IPS_LogMessage('PreheatScheduler', $formatted);
        $this->Debug('Log', $formatted);
    }

    private function Debug(string $messageName, $data, int $format = 0): void
    {
        if (is_array($data) || is_object($data)) {
            $encoded = json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
            if ($encoded !== false) {
                $data = $encoded;
                $format = 0;
            } else {
                $data = print_r($data, true);
                $format = 0;
            }
        }

        if (is_bool($data)) {
            $data = $data ? 'true' : 'false';
        }

        if (!is_string($data)) {
            $data = (string) $data;
        }

        $this->SendDebug($messageName, $data, $format);
    }

    private function UpdateEventOverview(array $events, int $now, ?string $errorMessage): void
    {
        $html = $this->BuildEventOverviewTable($events, $now, $errorMessage);
        $this->SetValue('EventOverviewHTML', $html);
    }

    private function BuildEventOverviewTable(array $events, int $now, ?string $errorMessage): string
    {
        $table = "<table style='width: 100%; border-collapse: collapse;'><tr><td style='padding: 5px; font-weight: bold;'>Event</td><td style='padding: 5px; font-weight: bold;'>Zeit</td><td style='padding: 5px; font-weight: bold;'>Status</td></tr>";

        if ($errorMessage !== null) {
            $table .= "<tr><td colspan='3' style='padding: 5px;'>" . htmlspecialchars($errorMessage) . '</td></tr>';
            return $table . '</table>';
        }

        $events = $this->FilterBlacklistedEvents($events);

        $tempVarID = $this->ReadPropertyInteger('TempVarID');
        $currentTemp = null;
        if ($tempVarID > 0 && IPS_VariableExists($tempVarID)) {
            $currentTemp = (float) GetValue($tempVarID);
        }

        $setpoint = $this->ReadPropertyFloat('SetpointWarm');
        $heatingRate = $this->ReadPropertyFloat('HeatingRate');
        $bufferSeconds = max(0, $this->ReadPropertyInteger('PreheatBufferMin')) * 60;

        $rowsAdded = false;

        $heatingVarID = $this->GetIDForIdent('HeatingDemand');
        $heatingActive = GetValueBoolean($heatingVarID);
        $trackedEventStart = $this->ReadAttributeInteger('LastEventStart');
        $demandHoldUntil = $this->ReadAttributeInteger('DemandHoldUntil');
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

            $isCancelled = isset($event['status']) && strtoupper((string) $event['status']) === 'CANCELLED';

            $status = '';
            if ($isCancelled) {
                $status = 'Abgesagt - Keine Heizung';
            } elseif ($now >= $eventStart && $now < $eventEnd) {
                $status = 'Veranstaltung läuft // Temperatur wird gehalten';
            } elseif (
                $heatingActive
                && $trackedEventStart > 0
                && $trackedEventStart === $eventStart
                && $demandHoldUntil > $now
                && !$isCancelled
                && $now < $eventStart
            ) {
                $status = 'Vorheizen Aktiv';
            } elseif ($now >= $preheatStart && $now < $eventStart) {
                $status = 'Vorheizen Aktiv';
            } elseif ($now < $preheatStart) {
                $diff = $preheatStart - $now;
                $status = $this->FormatCountdown($diff);
            } else {
                $status = '-';
            }

            $rowStyle = $isCancelled ? " style='color: #b30000;'" : '';

            $table .= '<tr' . $rowStyle . '>'
                . "<td style='padding: 5px;'>" . htmlspecialchars($summary) . '</td>'
                . "<td style='padding: 5px;'>" . htmlspecialchars($this->FormatEventDate($eventStart, $eventEnd)) . '</td>'
                . "<td style='padding: 5px;'>" . htmlspecialchars($status) . '</td>'
                . '</tr>';
        }

        if (!$rowsAdded) {
            $table .= "<tr><td colspan='3' style='padding: 5px;'>" . $this->Translate('Keine zukünftigen Veranstaltungen im Suchzeitraum.') . '</td></tr>';
        }

        return $table . '</table>';
    }

    private function FilterBlacklistedEvents(array $events): array
    {
        $rules = $this->GetBlacklistRules();
        if (empty($rules)) {
            $this->Debug('FilterBlacklist', 'No blacklist rules configured');
            return $events;
        }

        $filtered = [];
        foreach ($events as $event) {
            if (!is_array($event)) {
                continue;
            }
            if ($this->IsEventBlacklisted($event, $rules)) {
                $summary = is_string($event['summary'] ?? null) ? $event['summary'] : '';
                $this->Debug('FilterBlacklist', sprintf('Event filtered: %s', $summary));
                continue;
            }
            $filtered[] = $event;
        }

        $this->Debug('FilterBlacklist', sprintf('Events kept after filtering: %d', count($filtered)));

        return $filtered;
    }

    private function ExpandRecurringEvents(array $events, int $windowStart, int $windowEnd): array
    {
        $this->Debug('ExpandRecurring', sprintf('Expanding events between %d and %d', $windowStart, $windowEnd));
        $expanded = [];

        $cancelledOccurrences = [];
        $cancelledSeries = [];

        foreach ($events as $event) {
            if (!is_array($event)) {
                continue;
            }

            $uid = is_string($event['uid'] ?? null) ? trim($event['uid']) : '';
            $recurrenceId = isset($event['recurrence_id']) && is_int($event['recurrence_id']) ? $event['recurrence_id'] : null;
            $isCancelled = isset($event['status']) && strtoupper((string) $event['status']) === 'CANCELLED';

            if ($isCancelled && $uid !== '') {
                if ($recurrenceId !== null) {
                    if (!isset($cancelledOccurrences[$uid])) {
                        $cancelledOccurrences[$uid] = [];
                    }
                    $cancelledOccurrences[$uid][$recurrenceId] = true;
                    $this->Debug('ExpandRecurring', sprintf('Cancelled occurrence recorded for %s at %d', $uid, $recurrenceId));
                } else {
                    $cancelledSeries[$uid] = true;
                    $this->Debug('ExpandRecurring', sprintf('Cancelled series recorded for %s', $uid));
                }
            }
        }

        foreach ($events as $event) {
            if (!isset($event['start'], $event['end'])) {
                continue;
            }

            $duration = (int) $event['end'] - (int) $event['start'];
            if ($duration <= 0) {
                $this->Debug('ExpandRecurring', 'Skipping event with non-positive duration');
                continue;
            }

            $uid = is_string($event['uid'] ?? null) ? trim($event['uid']) : '';
            $isCancelled = isset($event['status']) && strtoupper((string) $event['status']) === 'CANCELLED';

            if ($isCancelled) {
                if ($event['end'] > $windowStart && $event['start'] <= $windowEnd) {
                    $expanded[] = $event;
                    $this->Debug('ExpandRecurring', sprintf('Keeping cancelled event within window: %s', $uid));
                }
                continue;
            }

            if ($uid !== '' && isset($cancelledSeries[$uid])) {
                $this->Debug('ExpandRecurring', sprintf('Skipping event due to cancelled series: %s', $uid));
                continue;
            }

            $rrule = '';
            if (isset($event['rrule']) && is_string($event['rrule'])) {
                $rrule = trim($event['rrule']);
            }

            $rdates = [];
            if (isset($event['rdates']) && is_array($event['rdates'])) {
                $rdates = array_values(array_filter($event['rdates'], 'is_int'));
            }

            $exdates = [];
            if (isset($event['exdates']) && is_array($event['exdates'])) {
                $exdates = array_values(array_filter($event['exdates'], 'is_int'));
            }

            $timezone = null;
            if (isset($event['timezone']) && is_string($event['timezone']) && $event['timezone'] !== '') {
                $timezone = $event['timezone'];
            }

            $occurrenceStarts = [(int) $event['start']];
            foreach ($rdates as $timestamp) {
                $occurrenceStarts[] = $timestamp;
            }

            if ($rrule !== '') {
                $generated = $this->GenerateRRuleOccurrences((int) $event['start'], $rrule, $timezone, $windowStart, $windowEnd);
                foreach ($generated as $timestamp) {
                    $occurrenceStarts[] = $timestamp;
                }
            }

            $occurrenceStarts = array_values(array_unique(array_filter($occurrenceStarts, 'is_int')));
            sort($occurrenceStarts);

            foreach ($occurrenceStarts as $startTimestamp) {
                if (in_array($startTimestamp, $exdates, true)) {
                    $this->Debug('ExpandRecurring', sprintf('Skipping excluded occurrence %s at %d', $uid, $startTimestamp));
                    continue;
                }

                if ($uid !== '' && isset($cancelledOccurrences[$uid][$startTimestamp])) {
                    $this->Debug('ExpandRecurring', sprintf('Skipping cancelled occurrence %s at %d', $uid, $startTimestamp));
                    continue;
                }

                $endTimestamp = $startTimestamp + $duration;
                if ($endTimestamp <= 0) {
                    $this->Debug('ExpandRecurring', 'Skipping occurrence with non-positive end time');
                    continue;
                }

                if ($startTimestamp > $windowEnd) {
                    continue;
                }

                $occurrence = $event;
                $occurrence['start'] = $startTimestamp;
                $occurrence['end'] = $endTimestamp;
                $occurrence['rrule'] = '';
                $occurrence['rdates'] = [];
                $occurrence['exdates'] = [];
                $expanded[] = $occurrence;
                $this->Debug('ExpandRecurring', sprintf('Occurrence added %s start=%d end=%d', $uid, $startTimestamp, $endTimestamp));
            }
        }

        usort($expanded, static fn($a, $b) => ($a['start'] ?? 0) <=> ($b['start'] ?? 0));

        $this->Debug('ExpandRecurring', sprintf('Total expanded events: %d', count($expanded)));

        return $expanded;
    }

    private function GenerateRRuleOccurrences(int $baseStart, string $rrule, ?string $timezoneName, int $windowStart, int $windowEnd): array
    {
        $this->Debug('GenerateRRule', sprintf('Generating occurrences for rule %s', $rrule));
        $rule = $this->ParseRRule($rrule);
        if (empty($rule)) {
            $this->Debug('GenerateRRule', 'Parsed rule is empty');
            return [];
        }

        $freq = strtoupper($rule['FREQ'] ?? '');
        if ($freq === '') {
            $this->Debug('GenerateRRule', 'Frequency missing in rule');
            return [];
        }

        $interval = max(1, (int) ($rule['INTERVAL'] ?? 1));
        $remaining = null;
        if (array_key_exists('COUNT', $rule)) {
            $count = max(0, (int) $rule['COUNT']);
            if ($count <= 1) {
                $this->Debug('GenerateRRule', 'Count exhausted or not sufficient');
                return [];
            }
            $remaining = $count - 1;
        }

        $until = null;
        if (array_key_exists('UNTIL', $rule)) {
            $until = $this->ParseRRuleUntil($rule['UNTIL'], $timezoneName, $baseStart);
            if ($until !== null && $until < $baseStart) {
                $this->Debug('GenerateRRule', 'Until date precedes base start');
                return [];
            }
        }

        $tzName = $timezoneName;
        if ($tzName === null || $tzName === '') {
            $tzName = date_default_timezone_get();
        }

        try {
            $tz = new DateTimeZone($tzName);
        } catch (\Throwable $e) {
            $tz = new DateTimeZone(date_default_timezone_get());
        }

        $baseDate = (new DateTimeImmutable('@' . $baseStart))->setTimezone($tz);

        $occurrences = [];

        if ($freq === 'DAILY') {
            $intervalSeconds = 86400 * $interval;
            $targetIndex = 0;
            if ($intervalSeconds > 0 && $windowStart > $baseStart) {
                $targetIndex = (int) floor(($windowStart - $baseStart) / $intervalSeconds);
            }
            $startIndex = max(1, $targetIndex);

            if ($remaining !== null) {
                $skipped = $startIndex - 1;
                if ($skipped >= $remaining) {
                    return [];
                }
                $remaining -= $skipped;
            }

            $current = $baseDate->add(new DateInterval('P' . ($startIndex * $interval) . 'D'));
            $iterations = 0;
            while ($remaining === null || $remaining > 0) {
                $iterations++;
                if ($iterations > 2000) {
                    break;
                }

                $timestamp = $current->getTimestamp();
                if ($until !== null && $timestamp > $until) {
                    break;
                }
                if ($timestamp > $windowEnd) {
                    break;
                }

                $occurrences[] = $timestamp;

                if ($remaining !== null) {
                    $remaining--;
                    if ($remaining <= 0) {
                        break;
                    }
                }

                $current = $current->add(new DateInterval('P' . $interval . 'D'));
            }

            $this->Debug('GenerateRRule', sprintf('Daily rule generated %d occurrences', count($occurrences)));
            return $occurrences;
        }

        if ($freq === 'WEEKLY') {
            $bydayTokens = [];
            if (array_key_exists('BYDAY', $rule)) {
                $bydayTokens = array_filter(array_map('trim', explode(',', (string) $rule['BYDAY'])));
            }

            $weekDays = [];
            foreach ($bydayTokens as $token) {
                $weekday = $this->MapWeekday($token);
                if ($weekday !== null) {
                    $weekDays[] = $weekday;
                }
            }
            if (empty($weekDays)) {
                $weekDays[] = (int) $baseDate->format('w');
            }
            $weekDays = array_values(array_unique($weekDays));
            sort($weekDays);

            $occurrencesPerInterval = count($weekDays);
            $baseWeekday = (int) $baseDate->format('w');

            $secondsDiff = max(0, $windowStart - $baseStart);
            $weeksDiff = (int) floor($secondsDiff / 604800);
            $intervalPeriods = (int) floor($weeksDiff / $interval);
            $startIntervalIndex = $intervalPeriods > 0 ? $intervalPeriods - 1 : 0;
            if ($startIntervalIndex < 0) {
                $startIntervalIndex = 0;
            }

            if ($remaining !== null) {
                $skippedOccurrences = $startIntervalIndex * $occurrencesPerInterval;
                if ($skippedOccurrences >= $remaining) {
                    return [];
                }
                $remaining -= $skippedOccurrences;
            }

            $iteration = 0;
            for ($intervalIndex = $startIntervalIndex; $remaining === null || $remaining > 0; $intervalIndex++) {
                $iteration++;
                if ($iteration > 520) {
                    break;
                }

                $weekOffset = $intervalIndex * $interval;
                foreach ($weekDays as $weekday) {
                    $dayOffset = ($weekday - $baseWeekday + 7) % 7 + 7 * $weekOffset;
                    if ($weekOffset === 0 && $dayOffset === 0) {
                        continue;
                    }

                    $candidate = $baseDate->modify('+' . $dayOffset . ' days');
                    $timestamp = $candidate->getTimestamp();

                    if ($timestamp <= $baseStart) {
                        continue;
                    }

                    if ($until !== null && $timestamp > $until) {
                        $remaining = 0;
                        break 2;
                    }

                    if ($timestamp < $windowStart) {
                        if ($remaining !== null) {
                            if ($remaining === 0) {
                                break 2;
                            }
                            $remaining--;
                            continue;
                        }
                        continue;
                    }

                    if ($timestamp > $windowEnd) {
                        $remaining = 0;
                        break 2;
                    }

                    $occurrences[] = $timestamp;

                    if ($remaining !== null) {
                        $remaining--;
                        if ($remaining <= 0) {
                            break 2;
                        }
                    }
                }
            }

            $this->Debug('GenerateRRule', sprintf('Weekly rule generated %d occurrences', count($occurrences)));
            return $occurrences;
        }

        $this->Debug('GenerateRRule', 'Frequency not supported');
        return [];
    }

    private function ParseRRule(string $rrule): array
    {
        $result = [];
        $parts = explode(';', $rrule);
        foreach ($parts as $part) {
            $part = trim($part);
            if ($part === '') {
                continue;
            }
            $pair = explode('=', $part, 2);
            if (count($pair) !== 2) {
                continue;
            }
            $result[strtoupper($pair[0])] = trim($pair[1]);
        }

        return $result;
    }

    private function ParseRRuleUntil(string $value, ?string $timezoneName, int $referenceTimestamp): ?int
    {
        $value = trim($value);
        if ($value === '') {
            return null;
        }

        $valueNoZ = $value;
        $tz = null;
        if (str_ends_with($valueNoZ, 'Z')) {
            $valueNoZ = substr($valueNoZ, 0, -1);
            $tz = new DateTimeZone('UTC');
        } else {
            $tzName = $timezoneName;
            if ($tzName === null || $tzName === '') {
                $tzName = date_default_timezone_get();
            }
            try {
                $tz = new DateTimeZone($tzName);
            } catch (\Throwable $e) {
                $tz = new DateTimeZone(date_default_timezone_get());
            }
        }

        $formats = ['Ymd\THis', 'Ymd\THi', 'Ymd'];
        foreach ($formats as $format) {
            $dateTime = DateTimeImmutable::createFromFormat($format, $valueNoZ, $tz);
            if ($dateTime instanceof DateTimeImmutable) {
                if ($format === 'Ymd') {
                    $dateTime = $dateTime->setTime(
                        (int) date('H', $referenceTimestamp),
                        (int) date('i', $referenceTimestamp),
                        (int) date('s', $referenceTimestamp)
                    );
                }

                return $dateTime->getTimestamp();
            }
        }

        return null;
    }

    private function MapWeekday(string $token): ?int
    {
        $token = strtoupper(trim($token));
        if ($token === '') {
            return null;
        }

        if (strlen($token) > 2) {
            $token = substr($token, -2);
        }

        $map = [
            'SU' => 0,
            'MO' => 1,
            'TU' => 2,
            'WE' => 3,
            'TH' => 4,
            'FR' => 5,
            'SA' => 6
        ];

        return $map[$token] ?? null;
    }

    private function GetBlacklistRules(): array
    {
        $raw = $this->ReadPropertyString('EventBlacklist');
        if ($raw === '') {
            $this->Debug('BlacklistRules', 'No blacklist configured');
            return [];
        }

        $decoded = json_decode($raw, true);
        if ($decoded === null && json_last_error() !== JSON_ERROR_NONE) {
            $this->Log('Unable to decode event blacklist: ' . json_last_error_msg());
            $this->Debug('BlacklistRules', 'Blacklist JSON decode failed');
            return [];
        }

        if (!is_array($decoded)) {
            $this->Debug('BlacklistRules', 'Blacklist configuration not an array');
            return [];
        }

        $rules = [];
        foreach ($decoded as $entry) {
            if (!is_array($entry)) {
                continue;
            }

            if (array_key_exists('pattern', $entry)) {
                if (!$this->legacyBlacklistWarningIssued) {
                    $this->legacyBlacklistWarningIssued = true;
                    $this->Log('Regex blacklist entries are no longer supported. Please update the blacklist configuration.');
                    $this->Debug('BlacklistRules', 'Legacy blacklist entry encountered');
                }
                continue;
            }

            $startsWith = trim((string) ($entry['starts_with'] ?? ''));
            $endsWith = trim((string) ($entry['ends_with'] ?? ''));

            if ($startsWith === '' && $endsWith === '') {
                continue;
            }

            $rules[] = [
                'starts_with' => $startsWith,
                'ends_with'   => $endsWith
            ];
        }

        $this->Debug('BlacklistRules', sprintf('Loaded blacklist rules: %d', count($rules)));

        return $rules;
    }

    private function IsEventBlacklisted(array $event, array $rules): bool
    {
        $summary = (string) ($event['summary'] ?? '');
        $summaryLower = mb_strtolower($summary);

        foreach ($rules as $rule) {
            $startsWith = mb_strtolower($rule['starts_with']);
            $endsWith = mb_strtolower($rule['ends_with']);

            $matchesStart = $startsWith === '' || str_starts_with($summaryLower, $startsWith);
            $matchesEnd = $endsWith === '' || str_ends_with($summaryLower, $endsWith);

            if ($matchesStart && $matchesEnd) {
                return true;
            }
        }

        return false;
    }

    private function FormatEventDate(int $startTimestamp, int $endTimestamp): string
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

        $dayName = $days[(int) date('w', $startTimestamp)];
        $day = (int) date('j', $startTimestamp);
        $month = $months[(int) date('n', $startTimestamp)] ?? '';

        $startHour = (int) date('G', $startTimestamp);
        $startMinute = (int) date('i', $startTimestamp);
        $endHour = (int) date('G', $endTimestamp);
        $endMinute = (int) date('i', $endTimestamp);

        return sprintf('%s %d. %s %02d:%02d-%02d:%02d', $dayName, $day, $month, $startHour, $startMinute, $endHour, $endMinute);
    }

    private function FormatCountdown(int $seconds): string
    {
        if ($seconds < 0) {
            $seconds = 0;
        }
        if ($seconds >= 86400) {
            $days = (int) ceil($seconds / 86400);
            return sprintf('Heizstart in %d Tagen', $days);
        }
        $hours = intdiv($seconds, 3600);
        $minutes = intdiv($seconds % 3600, 60);
        return sprintf('Heizstart in %d Std %d Min', $hours, $minutes);
    }
}
