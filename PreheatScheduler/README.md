# PreheatScheduler

PreheatScheduler fetches usage events from a CalDAV/ICS calendar and raises a heating demand flag early enough so the configured space reaches the warm setpoint when an event begins.

## Usage

1. Import the library into IP-Symcon and create an instance of **PreheatScheduler**.
2. Provide the calendar URL. Use a direct `.ics` export link when possible. For CalDAV collections append `?export` (Nextcloud/Owncloud) or rely on the module to try it automatically. Supply credentials if required.
3. Select the temperature status variable that reflects the current room temperature.
4. Configure the warm setpoint, heating rate (°C per hour), optional preheat buffer and evaluation interval.
5. Choose the hold strategy (keep heating demand until the event ends or only until it starts).
6. Add optional blacklist rules to ignore events whose titles start and/or end with specific text.

The module exposes one boolean variable `Heating Demand`. Link this variable to the controller that switches the real heating system. Optional string variables show the next event start and calculated preheat start in ISO format.

## Limitations

* Complex recurring events (RRULE) are not expanded; ensure the calendar export provides discrete events or pre-expanded series.
* The module controls only the demand signal; actual heating control must be implemented separately in Symcon.
* Authentication credentials are stored in Symcon configuration. Prefer using tokenised URLs when supported by the calendar server.
