"""
Email automation pipeline.

Background scheduler that:
  1. Polls a configured mailbox (Outlook or Gmail) on a fixed interval.
  2. For each new message with supported attachments, hands them to a
     processor callback that runs the existing extraction + report
     generation pipeline.
  3. Replies to the sender with the generated report attached.
  4. Marks the source message as read.

Settings live in `data/email_settings.json` so they can be edited from
the web UI without restarting the server. Secrets are stored on disk in
plain text — protect the host filesystem accordingly, or move them to a
secret manager in production.
"""

from __future__ import annotations

import json
import os
import threading
import time
import traceback
from datetime import datetime
from typing import Callable, Optional

from email_service import (
    EmailProviderError,
    FetchedMessage,
    build_provider,
)

SETTINGS_FILE = os.path.join('data', 'email_settings.json')

DEFAULT_SETTINGS: dict = {
    'provider': 'outlook',          # 'outlook' | 'gmail'
    'mailbox': '',                  # mailbox UPN (Outlook) or address (Gmail)
    'tenant_id': '',                # Outlook only
    'client_id': '',                # Outlook only
    'client_secret': '',            # Outlook only
    'app_password': '',             # Gmail only
    'poll_interval_seconds': 120,
}

_SECRET_KEYS = ('client_secret', 'app_password')


def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                return {**DEFAULT_SETTINGS, **(json.load(f) or {})}
        except Exception as e:
            print(f"[email_pipeline] failed to load settings: {e}")
    return dict(DEFAULT_SETTINGS)


def save_settings(updates: dict) -> dict:
    """Merge `updates` into existing settings and persist.

    Empty / placeholder secret values are ignored so the form can show a
    masked value without overwriting the stored secret on save.
    """
    current = load_settings()
    for k, v in (updates or {}).items():
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == '' and k in _SECRET_KEYS:
            continue
        if isinstance(v, str) and v.strip() in ('••••••••', '••••', '*****'):
            continue
        current[k] = v
    try:
        current['poll_interval_seconds'] = max(15, int(current.get('poll_interval_seconds') or 120))
    except (TypeError, ValueError):
        current['poll_interval_seconds'] = 120
    os.makedirs(os.path.dirname(SETTINGS_FILE) or '.', exist_ok=True)
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(current, f, indent=2)
    return current


def classify_attachment(file_name: str) -> str:
    """Best-effort filename → document slot mapping for the existing pipeline."""
    n = (file_name or '').lower()
    if 'iauditor' in n or 'safety' in n or 'inspection' in n or 'survey' in n:
        return 'iauditor_report'
    if 'pack' in n:
        return 'packing_list'
    if 'invoice' in n or 'commercial' in n:
        return 'commercial_invoice'
    if 'lading' in n or n.startswith('bl') or '_bl' in n or 'bill' in n:
        return 'bill_of_lading'
    # Default unknown PDFs/Docs to iAuditor — that slot accepts all formats
    # and drives the bulk of the report content.
    return 'iauditor_report'


# ---------------------------------------------------------------------------
# Scheduler
# ---------------------------------------------------------------------------

ProcessorFn = Callable[[FetchedMessage], Optional[str]]


class EmailAutomationScheduler:
    """Single background poller. Thread-safe start/stop."""

    def __init__(self) -> None:
        self._thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._lock = threading.Lock()
        self._processor: Optional[ProcessorFn] = None
        self.last_status: str = 'idle'
        self.last_run_at: Optional[str] = None
        self.last_error: Optional[str] = None
        self.processed_count: int = 0

    def configure(self, processor: ProcessorFn) -> None:
        self._processor = processor

    # -- lifecycle --------------------------------------------------------

    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self) -> dict:
        with self._lock:
            if self.is_running():
                return self.status()
            if self._processor is None:
                self.last_status = 'error'
                self.last_error = 'Scheduler has no processor configured.'
                return self.status()
            self._stop_event.clear()
            self._thread = threading.Thread(
                target=self._loop, name='email-automation', daemon=True
            )
            self._thread.start()
            self.last_status = 'running'
            self.last_error = None
        return self.status()

    def stop(self) -> dict:
        with self._lock:
            self._stop_event.set()
            self.last_status = 'stopping'
        return self.status()

    def status(self) -> dict:
        return {
            'running': self.is_running(),
            'last_status': self.last_status,
            'last_run_at': self.last_run_at,
            'last_error': self.last_error,
            'processed_count': self.processed_count,
        }

    # -- main loop --------------------------------------------------------

    def _loop(self) -> None:
        while not self._stop_event.is_set():
            settings = load_settings()
            try:
                provider = build_provider(settings)
                messages = provider.fetch_new_messages()
                self.last_run_at = datetime.now().isoformat()

                for msg in messages:
                    if self._stop_event.is_set():
                        break
                    try:
                        report_path = self._processor(msg)  # type: ignore[misc]
                        if report_path and msg.sender:
                            provider.send_reply(
                                to=msg.sender,
                                subject=(
                                    f"Inspection Report — "
                                    f"{msg.subject or 'Your submission'}"
                                ),
                                body=(
                                    "Hello,\n\n"
                                    "Please find attached the automatically "
                                    "generated inspection report for the "
                                    "documents you sent.\n\n"
                                    "Quest Marine — Automated Reporting"
                                ),
                                attachment_path=report_path,
                            )
                        provider.mark_read(msg.message_id)
                        self.processed_count += 1
                        self.last_error = None
                    except Exception as e:
                        traceback.print_exc()
                        self.last_error = f"Message {msg.message_id}: {e}"

                self.last_status = 'running'
            except EmailProviderError as e:
                self.last_status = 'error'
                self.last_error = str(e)
            except Exception as e:
                traceback.print_exc()
                self.last_status = 'error'
                self.last_error = str(e)

            interval = max(15, int(load_settings().get('poll_interval_seconds') or 120))
            for _ in range(interval):
                if self._stop_event.is_set():
                    break
                time.sleep(1)

        self.last_status = 'stopped'


scheduler = EmailAutomationScheduler()
