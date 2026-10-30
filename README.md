# ICS Import Card

A Home Assistant Lovelace card to import the events of an `.ics` file into an existing calendar.

1. Choose (or drop) an `.ics` file.
2. Check the events you want in the preview.
3. Pick the target calendar and click **Import**.

Works with any calendar entity that supports creating events (Local Calendar, Google Calendar, CalDAV, ...).
Standalone: it does not depend on any other card.

## Installation

**HACS** – add this repository as a custom repository (category *Dashboard*), then install *ICS Import Card*.

**Manual** – copy `dist/ics-import-card.js` to `config/www/`, then add the resource
`/local/ics-import-card.js` (type *JavaScript module*) in *Settings → Dashboards → Resources*.

## Configuration

```yaml
type: custom:ics-import-card
title: Import an .ics file   # optional
calendars:                   # optional – defaults to every calendar that accepts new events
  - calendar.family
  - calendar.work
```

## What is imported

| iCalendar | Imported as |
|-----------|-------------|
| `SUMMARY`, `LOCATION`, `DESCRIPTION` | title, location, description |
| `DTSTART` / `DTEND` or `DURATION` | start / end – all-day dates, UTC, floating times and `TZID` (IANA or Outlook/Windows names) |
| `RRULE` | repeating event |

Not imported: alarms, attendees, to-dos. For repeating events, changed or skipped occurrences
(`EXDATE`, `RECURRENCE-ID`) are lost – the preview flags them, and single-occurrence changes and
cancelled events are unchecked by default. An event without an end lasts one hour (one day if all-day).

The interface is in English or French, following the Home Assistant language.
