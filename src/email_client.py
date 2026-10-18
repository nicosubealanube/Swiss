"""Cliente de correo para leer la bandeja de entrada de ABM.Convenios@swissmedical.com.ar."""

import base64
import imaplib
import json
from typing import List, Optional
import urllib.parse
import urllib.request

import config
from src.email_parser import parse_email_body, parse_eml_file
from src.models import EmailExtractedData


def get_graph_access_token(
    tenant_id: str = config.AZURE_TENANT_ID,
    client_id: str = config.AZURE_CLIENT_ID,
    client_secret: str = config.AZURE_CLIENT_SECRET,
) -> str:
    """Obtiene un token OAuth2 de Azure AD para Microsoft Graph API (Client Credentials Flow)."""
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = urllib.parse.urlencode(
        {
            "client_id": client_id,
            "scope": "https://graph.microsoft.com/.default",
            "client_secret": client_secret,
            "grant_type": "client_credentials",
        }
    ).encode("utf-8")

    req = urllib.request.Request(
        token_url,
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urllib.request.urlopen(req) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    return payload["access_token"]


class GraphMailboxClient:
    """
    Lee correos no leídos con adjuntos de la casilla ABM.Convenios@swissmedical.com.ar
    usando Microsoft Graph API.
    """

    def __init__(
        self,
        access_token: str,
        mailbox: str = config.TARGET_MAILBOX,
    ) -> None:
        self.access_token = access_token
        self.mailbox = mailbox

    def fetch_unread_exception_emails(self, top: int = 25) -> List[EmailExtractedData]:
        """Busca correos no leídos en Inbox que tengan archivos adjuntos."""
        filter_q = urllib.parse.quote("isRead eq false and hasAttachments eq true")
        url = (
            f"https://graph.microsoft.com/v1.0/users/{self.mailbox}"
            f"/mailFolders/Inbox/messages?$filter={filter_q}&$top={top}"
            f"&$select=id,subject,from,receivedDateTime,body,hasAttachments"
        )
        req = urllib.request.Request(
            url,
            headers={"Authorization": f"Bearer {self.access_token}"},
            method="GET",
        )
        with urllib.request.urlopen(req) as resp:
            messages = json.loads(resp.read().decode("utf-8")).get("value", [])

        results: List[EmailExtractedData] = []
        for msg in messages:
            msg_id = msg["id"]
            subject = msg.get("subject", "")
            from_email_obj = msg.get("from", {}).get("emailAddress", {})
            sender_name = from_email_obj.get("name", "")
            sender_email = from_email_obj.get("address", "")
            received_dt = msg.get("receivedDateTime", "")
            body_content = msg.get("body", {}).get("content", "")

            extracted = parse_email_body(
                body_content=body_content,
                subject=subject,
                sender_name=sender_name,
                sender_email=sender_email,
                received_date=received_dt,
            )
            extracted.message_id = msg_id
            extracted.pdf_attachments = self._fetch_pdf_attachments(msg_id)

            if extracted.pdf_attachments:
                results.append(extracted)

        return results

    def _fetch_pdf_attachments(self, message_id: str) -> List[dict]:
        """Descarga los adjuntos .pdf de un mensaje específico."""
        url = (
            f"https://graph.microsoft.com/v1.0/users/{self.mailbox}"
            f"/messages/{message_id}/attachments"
        )
        req = urllib.request.Request(
            url,
            headers={"Authorization": f"Bearer {self.access_token}"},
            method="GET",
        )
        with urllib.request.urlopen(req) as resp:
            attachments = json.loads(resp.read().decode("utf-8")).get("value", [])

        pdfs = []
        for att in attachments:
            name = att.get("name", "")
            content_type = att.get("contentType", "")
            if name.lower().endswith(".pdf") or content_type == "application/pdf":
                content_b64 = att.get("contentBytes", "")
                if content_b64:
                    pdfs.append(
                        {
                            "filename": name,
                            "content": base64.b64decode(content_b64),
                        }
                    )
        return pdfs

    def mark_as_read(self, message_id: str) -> None:
        """Marca el correo como leído una vez cargado exitosamente en el Excel."""
        url = f"https://graph.microsoft.com/v1.0/users/{self.mailbox}/messages/{message_id}"
        body = json.dumps({"isRead": True}).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=body,
            headers={
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            },
            method="PATCH",
        )
        with urllib.request.urlopen(req):
            pass


class ImapMailboxClient:
    """Alternativa por protocolo IMAP estándar para leer correos no leídos con PDF."""

    def __init__(
        self,
        server: str = config.IMAP_SERVER,
        user: str = config.IMAP_USER,
        password: str = config.IMAP_PASSWORD,
    ) -> None:
        self.server = server
        self.user = user
        self.password = password

    def fetch_unread_exception_emails(self, mark_seen: bool = True) -> List[EmailExtractedData]:
        import email as email_pkg

        results: List[EmailExtractedData] = []
        with imaplib.IMAP4_SSL(self.server) as mail:
            mail.login(self.user, self.password)
            mail.select("INBOX")
            status, messages = mail.search(None, "(UNSEEN)")
            if status != "OK" or not messages[0]:
                return results

            for num in messages[0].split():
                status, data = mail.fetch(num, "(RFC822)")
                if status != "OK" or not data or not data[0]:
                    continue
                raw_email = data[0][1]
                msg = email_pkg.message_from_bytes(raw_email)

                subject = msg.get("Subject", "")
                sender_name, sender_email = email_pkg.utils.parseaddr(msg.get("From", ""))
                date_header = msg.get("Date", "")

                html_body = ""
                text_body = ""
                pdf_attachments = []

                if msg.is_multipart():
                    for part in msg.walk():
                        ctype = part.get_content_type()
                        fname = part.get_filename() or ""
                        if fname.lower().endswith(".pdf") or ctype == "application/pdf":
                            payload = part.get_payload(decode=True)
                            if payload:
                                pdf_attachments.append({"filename": fname, "content": payload})
                        elif ctype == "text/html" and not html_body:
                            payload = part.get_payload(decode=True)
                            if payload:
                                html_body = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
                        elif ctype == "text/plain" and not text_body:
                            payload = part.get_payload(decode=True)
                            if payload:
                                text_body = payload.decode(part.get_content_charset() or "utf-8", errors="replace")

                if pdf_attachments:
                    extracted = parse_email_body(
                        body_content=html_body or text_body,
                        subject=subject,
                        sender_name=sender_name,
                        sender_email=sender_email,
                        received_date=date_header,
                    )
                    extracted.message_id = num.decode("utf-8")
                    extracted.pdf_attachments = pdf_attachments
                    results.append(extracted)
                    if mark_seen:
                        mail.store(num, "+FLAGS", "\\Seen")

        return results
