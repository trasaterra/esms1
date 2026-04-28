from __future__ import annotations

import json
import re
from datetime import date, datetime, time, timedelta, timezone
from email.utils import formatdate
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Lock
from typing import Any
from urllib.error import URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo


HOST = "127.0.0.1"
PORT = 8000
CALENDAR_ICS_URL = "https://ical.echalk.com/tQKq5cnYG4EIDSuQaJurZXbuXu57067msoD3c8QBAt81"
CACHE_TTL_SECONDS = 300
LOCAL_TIMEZONE = ZoneInfo("America/New_York")
ROOT_DIR = Path(__file__).resolve().parent

_CACHE_LOCK = Lock()
_CACHE_PAYLOAD: dict[str, Any] | None = None
_CACHE_EXPIRES_AT = 0.0


def unfold_ical_lines(raw_text: str) -> list[str]:
    unfolded: list[str] = []
    for line in raw_text.splitlines():
        if line.startswith((" ", "\t")) and unfolded:
            unfolded[-1] += line[1:]
        else:
            unfolded.append(line)
    return unfolded


def unescape_ical_text(value: str) -> str:
    return (
        value.replace("\\n", " ")
        .replace("\\N", " ")
        .replace("\\,", ",")
        .replace("\\;", ";")
        .replace("\\\\", "\\")
        .strip()
    )


def clean_summary(summary: str) -> str:
    cleaned = re.sub(r"\s*\[(?:MS114.*?|Public calendar)\]", "", summary)
    return re.sub(r"\s+", " ", cleaned).strip()


def parse_ical_datetime(raw_value: str, parameters: str) -> tuple[datetime, bool]:
    raw_value = raw_value.strip()
    params_upper = parameters.upper()

    if "VALUE=DATE" in params_upper or re.fullmatch(r"\d{8}", raw_value):
        parsed_date = datetime.strptime(raw_value[:8], "%Y%m%d").date()
        return datetime.combine(parsed_date, time.min, tzinfo=LOCAL_TIMEZONE), True

    if raw_value.endswith("Z"):
        parsed_datetime = datetime.strptime(raw_value, "%Y%m%dT%H%M%SZ").replace(tzinfo=timezone.utc)
        return parsed_datetime.astimezone(LOCAL_TIMEZONE), False

    parsed_datetime = datetime.strptime(raw_value, "%Y%m%dT%H%M%S")
    return parsed_datetime.replace(tzinfo=LOCAL_TIMEZONE), False


def parse_ical_events(ics_text: str) -> list[dict[str, Any]]:
    events: list[dict[str, Any]] = []
    current_event: dict[str, str] | None = None

    for line in unfold_ical_lines(ics_text):
        if line == "BEGIN:VEVENT":
            current_event = {}
            continue

        if line == "END:VEVENT":
            if current_event:
                summary = clean_summary(unescape_ical_text(current_event.get("SUMMARY", "")))
                if summary and "DTSTART" in current_event and "DTEND" in current_event:
                    start_value, start_params = current_event["DTSTART"].split("||", 1)
                    end_value, end_params = current_event["DTEND"].split("||", 1)
                    start_dt, all_day = parse_ical_datetime(start_value, start_params)
                    end_dt, _ = parse_ical_datetime(end_value, end_params)
                    description = unescape_ical_text(current_event.get("DESCRIPTION", ""))
                    events.append({
                        "title": summary,
                        "description": description,
                        "start": start_dt,
                        "end": end_dt,
                        "allDay": all_day,
                    })
            current_event = None
            continue

        if current_event is None or ":" not in line:
            continue

        name_and_params, value = line.split(":", 1)
        property_name, _, parameters = name_and_params.partition(";")
        property_name = property_name.upper()

        if property_name in {"SUMMARY", "DESCRIPTION"}:
            current_event[property_name] = value
        elif property_name in {"DTSTART", "DTEND"}:
            current_event[property_name] = f"{value}||{parameters}"

    normalized_events: list[dict[str, Any]] = []
    for event in events:
        normalized_events.append({
            **event,
            "description": event["description"].split("Link to the calendar with original event:", 1)[0].strip(),
        })
    return normalized_events


def format_day_label(event_start: datetime, all_day: bool) -> str:
    day_label = event_start.strftime("%a %b ").upper() + str(event_start.day)
    if all_day:
        return day_label

    time_label = event_start.strftime("%I:%M %p").lstrip("0")
    return f"{day_label} | {time_label}"


def build_weekly_events_payload() -> dict[str, Any]:
    request = Request(
        CALENDAR_ICS_URL,
        headers={
            "User-Agent": "ESMS-Screen/1.0",
            "Accept": "text/calendar,text/plain;q=0.9,*/*;q=0.8",
        },
    )

    with urlopen(request, timeout=20) as response:
        ics_text = response.read().decode("utf-8")

    now = datetime.now(LOCAL_TIMEZONE)
    week_start = datetime.combine(now.date(), time.min, tzinfo=LOCAL_TIMEZONE)
    days_until_next_monday = (7 - week_start.weekday()) % 7 or 7
    week_end = week_start + timedelta(days=days_until_next_monday)

    weekly_events = []
    for event in parse_ical_events(ics_text):
        if event["end"] <= week_start or event["start"] >= week_end:
            continue

        weekly_events.append({
            "title": event["title"],
            "description": event["description"],
            "allDay": event["allDay"],
            "start": event["start"].isoformat(),
            "end": event["end"].isoformat(),
            "label": format_day_label(event["start"], event["allDay"]),
        })

    weekly_events.sort(key=lambda event: event["start"])
    week_end_display = (week_end - timedelta(days=1)).date()
    heading = f"This Week | {week_start.strftime('%b')} {week_start.day} - {week_end_display.strftime('%b')} {week_end_display.day}"

    return {
        "heading": heading,
        "generatedAt": now.isoformat(),
        "events": weekly_events,
    }


def get_cached_weekly_events_payload() -> dict[str, Any]:
    global _CACHE_EXPIRES_AT, _CACHE_PAYLOAD

    now_ts = datetime.now(timezone.utc).timestamp()
    with _CACHE_LOCK:
        if _CACHE_PAYLOAD and now_ts < _CACHE_EXPIRES_AT:
            return _CACHE_PAYLOAD

    payload = build_weekly_events_payload()

    with _CACHE_LOCK:
        _CACHE_PAYLOAD = payload
        _CACHE_EXPIRES_AT = now_ts + CACHE_TTL_SECONDS
        return _CACHE_PAYLOAD


class ScreenRequestHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, directory=str(ROOT_DIR), **kwargs)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/weekly-events":
            self.serve_weekly_events()
            return
        super().do_GET()

    def serve_weekly_events(self) -> None:
        try:
            payload = get_cached_weekly_events_payload()
            body = json.dumps(payload).encode("utf-8")
            self.send_response(HTTPStatus.OK)
        except (TimeoutError, URLError, ValueError) as error:
            body = json.dumps({"error": str(error), "events": []}).encode("utf-8")
            self.send_response(HTTPStatus.BAD_GATEWAY)

        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Last-Modified", formatdate(usegmt=True))
        self.end_headers()
        self.wfile.write(body)


def main() -> None:
    with ThreadingHTTPServer((HOST, PORT), ScreenRequestHandler) as httpd:
        print(f"Serving animated screen on http://{HOST}:{PORT}")
        httpd.serve_forever()


if __name__ == "__main__":
    main()