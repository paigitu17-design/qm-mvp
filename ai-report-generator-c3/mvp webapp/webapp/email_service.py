"""
Multi-provider email service for the Quest Marine inspection pipeline.

Providers (on-demand pull, NO webhooks):
    * GraphEmailProvider  — Microsoft 365 / Outlook via Microsoft Graph
                            (MSAL client-credentials, app-only).
    * GmailProvider       — Gmail via IMAP (read) + SMTP (send), using an
                            app password.

Common surface:
    fetch_new_messages() -> list[FetchedMessage]
    mark_read(message_id)
    send_reply(to, subject, body, attachment_path)

`build_provider(settings)` returns the right provider for a settings dict.
"""

from __future__ import annotations

import base64
import email
import email.utils
import imaplib
import os
import smtplib
import ssl
from dataclasses import dataclass, field
from email.message import EmailMessage
from typing import List, Optional

import msal
import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SUPPORTED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.png', '.jpg', '.jpeg')


def is_supported_attachment(name: str) -> bool:
    return bool(name) and name.lower().endswith(SUPPORTED_EXTENSIONS)


@dataclass
class FetchedAttachment:
    file_name: str
    content_type: str
    content_bytes: bytes


@dataclass
class FetchedMessage:
    message_id: str
    subject: str
    sender: str
    received_at: str
    attachments: List[FetchedAttachment] = field(default_factory=list)


class EmailProviderError(Exception):
    pass


# ---------------------------------------------------------------------------
# Microsoft 365 / Outlook (Microsoft Graph, app-only)
# ---------------------------------------------------------------------------

class GraphEmailProvider:
    name = 'outlook'

    def __init__(self, tenant_id: str, client_id: str, client_secret: str, mailbox: str):
        if not all([tenant_id, client_id, client_secret, mailbox]):
            raise EmailProviderError(
                "Microsoft 365 credentials are incomplete. "
                "Required: tenant_id, client_id, client_secret, mailbox."
            )
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.mailbox = mailbox

    def _token(self) -> str:
        app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if 'access_token' not in result:
            raise EmailProviderError(
                f"Token acquisition failed: {result.get('error_description', result)}"
            )
        return result['access_token']

    def _headers(self) -> dict:
        return {'Authorization': f'Bearer {self._token()}', 'Accept': 'application/json'}

    def fetch_new_messages(self) -> List[FetchedMessage]:
        headers = self._headers()
        # Graph rejects (isRead eq false AND hasAttachments eq true) + orderby
        # as "InefficientFilter". Filter only on isRead and check
        # hasAttachments client-side.
        list_url = (
            f"{GRAPH_BASE}/users/{self.mailbox}/mailFolders/Inbox/messages"
            "?$filter=isRead eq false"
            "&$orderby=receivedDateTime desc"
            "&$top=25"
            "&$select=id,subject,from,receivedDateTime,hasAttachments"
        )
        resp = requests.get(list_url, headers=headers, timeout=30)
        if resp.status_code != 200:
            raise EmailProviderError(
                f"Graph list messages failed: {resp.status_code} {resp.text}"
            )

        out: List[FetchedMessage] = []
        for msg in resp.json().get('value', []):
            if not msg.get('hasAttachments'):
                continue
            att_resp = requests.get(
                f"{GRAPH_BASE}/users/{self.mailbox}/messages/{msg['id']}/attachments",
                headers=headers, timeout=30,
            )
            if att_resp.status_code != 200:
                continue
            attachments: List[FetchedAttachment] = []
            for att in att_resp.json().get('value', []):
                if att.get('@odata.type') != '#microsoft.graph.fileAttachment':
                    continue
                name = att.get('name') or ''
                if not is_supported_attachment(name):
                    continue
                content_b64 = att.get('contentBytes')
                if not content_b64:
                    continue
                attachments.append(FetchedAttachment(
                    file_name=name,
                    content_type=att.get('contentType') or 'application/octet-stream',
                    content_bytes=base64.b64decode(content_b64),
                ))
            if not attachments:
                continue
            out.append(FetchedMessage(
                message_id=msg['id'],
                subject=msg.get('subject', '') or '',
                sender=(msg.get('from', {}) or {}).get('emailAddress', {}).get('address', ''),
                received_at=msg.get('receivedDateTime', '') or '',
                attachments=attachments,
            ))
        return out

    def mark_read(self, message_id: str) -> None:
        headers = {**self._headers(), 'Content-Type': 'application/json'}
        try:
            requests.patch(
                f"{GRAPH_BASE}/users/{self.mailbox}/messages/{message_id}",
                headers=headers, json={'isRead': True}, timeout=30,
            )
        except Exception:
            pass

    def send_reply(self, to: str, subject: str, body: str, attachment_path: str) -> None:
        headers = {**self._headers(), 'Content-Type': 'application/json'}
        with open(attachment_path, 'rb') as f:
            content_b64 = base64.b64encode(f.read()).decode('ascii')
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": to}}],
                "attachments": [{
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(attachment_path),
                    "contentBytes": content_b64,
                }],
            },
            "saveToSentItems": True,
        }
        r = requests.post(
            f"{GRAPH_BASE}/users/{self.mailbox}/sendMail",
            headers=headers, json=payload, timeout=60,
        )
        if r.status_code >= 300:
            raise EmailProviderError(f"Graph sendMail failed: {r.status_code} {r.text}")


# ---------------------------------------------------------------------------
# Gmail (IMAP + SMTP, app password)
# ---------------------------------------------------------------------------

class GmailProvider:
    name = 'gmail'

    IMAP_HOST = 'imap.gmail.com'
    SMTP_HOST = 'smtp.gmail.com'
    SMTP_PORT = 465

    def __init__(self, address: str, app_password: str):
        if not address or not app_password:
            raise EmailProviderError(
                "Gmail credentials are incomplete. Required: mailbox, app_password."
            )
        self.address = address
        self.password = app_password

    def _imap(self) -> imaplib.IMAP4_SSL:
        try:
            m = imaplib.IMAP4_SSL(self.IMAP_HOST)
            m.login(self.address, self.password)
            return m
        except imaplib.IMAP4.error as e:
            raise EmailProviderError(f"Gmail IMAP login failed: {e}")

    def fetch_new_messages(self) -> List[FetchedMessage]:
        m = self._imap()
        try:
            m.select('INBOX')
            typ, data = m.search(None, '(UNSEEN)')
            if typ != 'OK' or not data or not data[0]:
                return []
            ids = data[0].split()
            out: List[FetchedMessage] = []
            for mid in ids[-25:]:
                typ, msg_data = m.fetch(mid, '(RFC822)')
                if typ != 'OK' or not msg_data or not msg_data[0]:
                    continue
                msg = email.message_from_bytes(msg_data[0][1])
                attachments: List[FetchedAttachment] = []
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    fn = part.get_filename()
                    if not fn or not is_supported_attachment(fn):
                        continue
                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue
                    attachments.append(FetchedAttachment(
                        file_name=fn,
                        content_type=part.get_content_type(),
                        content_bytes=payload,
                    ))
                if not attachments:
                    # No supported attachment — mark seen so we don't reprocess
                    m.store(mid, '+FLAGS', '\\Seen')
                    continue
                out.append(FetchedMessage(
                    message_id=mid.decode(),
                    subject=msg.get('Subject', '') or '',
                    sender=email.utils.parseaddr(msg.get('From', ''))[1],
                    received_at=msg.get('Date', '') or '',
                    attachments=attachments,
                ))
            return out
        finally:
            try:
                m.logout()
            except Exception:
                pass

    def mark_read(self, message_id: str) -> None:
        m = self._imap()
        try:
            m.select('INBOX')
            m.store(message_id.encode(), '+FLAGS', '\\Seen')
        finally:
            try:
                m.logout()
            except Exception:
                pass

    def send_reply(self, to: str, subject: str, body: str, attachment_path: str) -> None:
        em = EmailMessage()
        em['From'] = self.address
        em['To'] = to
        em['Subject'] = subject
        em.set_content(body)
        with open(attachment_path, 'rb') as f:
            data = f.read()
        em.add_attachment(
            data,
            maintype='application',
            subtype='octet-stream',
            filename=os.path.basename(attachment_path),
        )
        ctx = ssl.create_default_context()
        try:
            with smtplib.SMTP_SSL(self.SMTP_HOST, self.SMTP_PORT, context=ctx) as s:
                s.login(self.address, self.password)
                s.send_message(em)
        except Exception as e:
            raise EmailProviderError(f"Gmail SMTP send failed: {e}")


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

def build_provider(settings: dict):
    provider = (settings.get('provider') or '').lower()
    if provider == 'outlook':
        return GraphEmailProvider(
            tenant_id=settings.get('tenant_id', ''),
            client_id=settings.get('client_id', ''),
            client_secret=settings.get('client_secret', ''),
            mailbox=settings.get('mailbox', ''),
        )
    if provider == 'gmail':
        return GmailProvider(
            address=settings.get('mailbox', ''),
            app_password=settings.get('app_password', ''),
        )
    raise EmailProviderError(f"Unknown email provider: {provider!r}")


# ---------------------------------------------------------------------------
# Backwards-compat shim for the old single-shot fetch endpoint
# ---------------------------------------------------------------------------

@dataclass
class FetchedEmailDocument:
    """Legacy single-attachment record used by /email/fetch-latest-document."""
    file_name: str
    content_type: str
    content_bytes: bytes
    message_id: str
    subject: str
    sender: str
    received_at: str


class EmailServiceError(EmailProviderError):
    pass


class GraphEmailService:
    """Thin wrapper kept for the manual 'Fetch from M365' button.

    Loads settings from the JSON settings file first, then falls back to
    environment variables.
    """

    def __init__(self):
        from email_pipeline import load_settings  # local import to avoid cycle
        settings = load_settings()
        self.provider = GraphEmailProvider(
            tenant_id=settings.get('tenant_id') or os.environ.get('GRAPH_TENANT_ID', ''),
            client_id=settings.get('client_id') or os.environ.get('GRAPH_CLIENT_ID', ''),
            client_secret=settings.get('client_secret') or os.environ.get('GRAPH_CLIENT_SECRET', ''),
            mailbox=settings.get('mailbox') or os.environ.get('GRAPH_MAILBOX_USER_ID', ''),
        )

    def fetch_latest_supported_attachment(self) -> Optional[FetchedEmailDocument]:
        msgs = self.provider.fetch_new_messages()
        if not msgs:
            return None
        msg = msgs[0]
        att = msg.attachments[0]
        self.provider.mark_read(msg.message_id)
        return FetchedEmailDocument(
            file_name=att.file_name,
            content_type=att.content_type,
            content_bytes=att.content_bytes,
            message_id=msg.message_id,
            subject=msg.subject,
            sender=msg.sender,
            received_at=msg.received_at,
        )
