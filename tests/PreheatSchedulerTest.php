<?php

declare(strict_types=1);

date_default_timezone_set('Europe/Berlin');

$GLOBALS['IPS_VARIABLES'] = [];

class IPSModule
{
    protected int $InstanceID;
    protected array $properties = [];
    protected array $attributes = [];
    protected array $idents = [];
    private int $nextId = 1;

    public function __construct(int $InstanceID = 0)
    {
        $this->InstanceID = $InstanceID;
    }

    public function Create()
    {
    }

    public function ApplyChanges()
    {
    }

    protected function RegisterPropertyString(string $name, string $default): void
    {
        $this->properties[$name] = $default;
    }

    protected function RegisterPropertyInteger(string $name, int $default): void
    {
        $this->properties[$name] = $default;
    }

    protected function RegisterPropertyFloat(string $name, float $default): void
    {
        $this->properties[$name] = $default;
    }

    protected function RegisterVariableBoolean(string $ident, string $name): int
    {
        $id = $this->nextId++;
        $this->idents[$ident] = $id;
        $GLOBALS['IPS_VARIABLES'][$id] = ['value' => false];
        return $id;
    }

    protected function RegisterVariableString(string $ident, string $name): int
    {
        $id = $this->nextId++;
        $this->idents[$ident] = $id;
        $GLOBALS['IPS_VARIABLES'][$id] = ['value' => ''];
        return $id;
    }

    protected function RegisterTimer(string $ident, int $interval, string $script): void
    {
    }

    protected function RegisterAttributeInteger(string $name, int $default): void
    {
        $this->attributes[$name] = $default;
    }

    protected function WriteAttributeInteger(string $name, int $value): void
    {
        $this->attributes[$name] = $value;
    }

    protected function ReadAttributeInteger(string $name): int
    {
        return (int) ($this->attributes[$name] ?? 0);
    }

    protected function SetTimerInterval(string $ident, int $milliseconds): void
    {
    }

    protected function RegisterMessage(int $SenderID, int $Message): void
    {
    }

    protected function UnregisterMessage(int $SenderID, int $Message): void
    {
    }

    protected function ReadPropertyInteger(string $name): int
    {
        return (int) ($this->properties[$name] ?? 0);
    }

    protected function ReadPropertyString(string $name): string
    {
        return (string) ($this->properties[$name] ?? '');
    }

    protected function ReadPropertyFloat(string $name): float
    {
        return (float) ($this->properties[$name] ?? 0.0);
    }

    protected function GetIDForIdent(string $ident): int
    {
        return $this->idents[$ident] ?? 0;
    }

    protected function SetValue(string $ident, $value): void
    {
        $id = $this->GetIDForIdent($ident);
        if ($id > 0) {
            $GLOBALS['IPS_VARIABLES'][$id]['value'] = $value;
        }
    }

    protected function SendDebug(string $Message, string $Data, int $Format): void
    {
    }

    protected function SetStatus(int $status): void
    {
    }

    protected function Translate(string $text): string
    {
        return $text;
    }
}

function IPS_SetVariableCustomProfile(int $id, string $profile): void
{
}

function IPS_VariableExists(int $id): bool
{
    return array_key_exists($id, $GLOBALS['IPS_VARIABLES']);
}

function GetValueBoolean(int $id): bool
{
    return (bool) ($GLOBALS['IPS_VARIABLES'][$id]['value'] ?? false);
}

function GetValue(int $id)
{
    return $GLOBALS['IPS_VARIABLES'][$id]['value'] ?? null;
}

function IPS_LogMessage(string $module, string $message): void
{
}

require_once __DIR__ . '/../PreheatScheduler/module.php';

class PreheatSchedulerTestable extends PreheatScheduler
{
    private ?array $nextEvent = null;

    public function setProperty(string $name, $value): void
    {
        $this->properties[$name] = $value;
    }

    public function setNextEvent(?array $event): void
    {
        $this->nextEvent = $event;
    }

    public function setAttribute(string $name, int $value): void
    {
        $this->WriteAttributeInteger($name, $value);
    }

    public function getIdentId(string $ident): int
    {
        return $this->GetIDForIdent($ident);
    }

    protected function DetermineNextEvent(int $now): ?array
    {
        return $this->nextEvent;
    }
}

$module = new PreheatSchedulerTestable(1);
$module->Create();
$module->setProperty('SetpointWarm', 21.0);
$module->setProperty('HeatingRate', 2.0);
$module->setProperty('PreheatBufferMin', 0);
$module->setProperty('HoldStrategy', 0);

$tempVarID = 99;
$module->setProperty('TempVarID', $tempVarID);
$GLOBALS['IPS_VARIABLES'][$tempVarID] = ['value' => 18.0];

$eventStart = (new DateTime('2025-12-17 19:30:00', new DateTimeZone('Europe/Berlin')))->getTimestamp();
$eventEnd = (new DateTime('2025-12-17 21:30:00', new DateTimeZone('Europe/Berlin')))->getTimestamp();
$module->setNextEvent(['start' => $eventStart, 'end' => $eventEnd]);

$heatingVarID = $module->getIdentId('HeatingDemand');
$GLOBALS['IPS_VARIABLES'][$heatingVarID]['value'] = true;

$module->setAttribute('DemandHoldUntil', $eventEnd);

$now = (new DateTime('2025-12-17 12:25:00', new DateTimeZone('Europe/Berlin')))->getTimestamp();
$result = $module->Recalculate($now);

assert($result === false, 'Heating demand should be false before preheat window.');
assert(GetValueBoolean($heatingVarID) === false, 'Heating variable must be turned off before preheat window.');

echo "Test passed\n";
