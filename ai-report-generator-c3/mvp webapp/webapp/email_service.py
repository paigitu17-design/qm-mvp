"""
Microsoft 365 Graph email service.

On-demand pull (NO webhooks): authenticates via MSAL client-credentials,
finds the latest UNREAD message in the configured mailbox that has a
supported attachment (PDF / DOCX / XLSX / PNG / JPG), downloads the first
matching attachment, marks the message as read, and returns the bytes +
metadata so the existing upload pipeline can store/process it.

Required environment variables:
    GRAPH_TENANT_ID
    GRAPH_CLIENT_ID
    GRAPH_CLIENT_SECRET
    GRAPH_MAILBOX_USER_ID   (UPN of mailbox, e.g. reports@tenant.onmicrosoft.com)

Azure AD app registration must have application permission
`Mail.ReadWrite` (or `Mail.Read` if you skip the mark-as-read patch)
with admin consent. Recommended: scope to a single mailbox via
ApplicationAccessPolicy.
"""

import os
from dataclasses import dataclass
from typing import Optional

import msal
import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SUPPORTED_EXTENSIONS = ('.pdf', '.docx', '.xlsx', '.png', '.jpg', '.jpeg')


@dataclass
class FetchedEmailDocument:
    file_name: str
    content_type: str
    content_bytes: bytes
    message_id: str
    subject: str
    sender: str
    received_at: str


class EmailServiceError(Exception):
    pass


class GraphEmailService:
    def __init__(self):
        self.tenant_id = os.environ.get('GRAPH_TENANT_ID', '')
        self.client_id = os.environ.get('GRAPH_CLIENT_ID', '')
        self.client_secret = os.environ.get('GRAPH_CLIENT_SECRET', '')
        self.mailbox = os.environ.get('GRAPH_MAILBOX_USER_ID', '')
        if not all([self.tenant_id, self.client_id, self.client_secret, self.mailbox]):
            raise EmailServiceError(
                "Microsoft Graph credentials are not configured. "
                "Set GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, "
                "GRAPH_MAILBOX_USER_ID."
            )

    def _acquire_token(self) -> str:
        app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )
        result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if 'access_token' not in result:
            raise EmailServiceError(
                f"Token acquisition failed: {result.get('error_description', result)}"
            )
        return result['access_token']

    def _headers(self, token: str) -> dict:
        return {'Authorization': f'Bearer {token}', 'Accept': 'application/json'}

    def fetch_latest_supported_attachment(self) -> Optional[FetchedEmailDocument]:
        token = self._acquire_token()
        headers = self._headers(token)

        # Search unread inbox messages with attachments, newest first
        list_url = (
            f"{GRAPH_BASE}/users/{self.mailbox}/mailFolders/Inbox/messages"
            "?$filter=isRead eq false and hasAttachments eq true"
            "&$orderby=receivedDateTime desc"
            "&$top=25"
            "&$select=id,subject,from,receivedDateTime,hasAttachments"
        )
        resp = requests.get(list_url, headers=headers, timeout=30)
        if resp.status_code != 200:
            raise EmailServiceError(
                f"Graph list messages failed: {resp.status_code} {resp.text}"
            )

        for msg in resp.json().get('value', []):
            msg_id = msg['id']
            att_url = (
                f"{GRAPH_BASE}/users/{self.mailbox}/messages/{msg_id}/attachments"
            )
            att_resp = requests.get(att_url, headers=headers, timeout=30)
            if att_resp.status_code != 200:
                continue

            for att in att_resp.json().get('value', []):
                if att.get('@odata.type') != '#microsoft.graph.fileAttachment':
                    continue
                name = att.get('name') or ''
                if not name.lower().endswith(SUPPORTED_EXTENSIONS):
                    continue

                import base64
                content_b64 = att.get('contentBytes')
                if not content_b64:
                    continue
                content_bytes = base64.b64decode(content_b64)

                # Mark message as read so the next click skips it
                try:
                    requests.patch(
                        f"{GRAPH_BASE}/users/{self.mailbox}/messages/{msg_id}",
                        headers={**headers, 'Content-Type': 'application/json'},
                        json={'isRead': True},
                        timeout=30,
                    )
                except Exception:
                    pass

                from_addr = (
                    msg.get('from', {})
                    .get('emailAddress', {})
                    .get('address', '')
                )
                return FetchedEmailDocument(
                    file_name=name,
                    content_type=att.get('contentType') or 'application/octet-stream',
                    content_bytes=content_bytes,
                    message_id=msg_id,
                    subject=msg.get('subject', ''),
                    sender=from_addr,
                    received_at=msg.get('receivedDateTime', ''),
                )

        return None
