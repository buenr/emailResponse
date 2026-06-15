#!/usr/bin/env python3
import argparse
import base64
import io
import json
import os
import re
import logging
import time
import threading
from typing import Optional
from dataclasses import dataclass, field
from contextlib import contextmanager

from bs4 import BeautifulSoup
from docx import Document
from openai import APIConnectionError, APIError, RateLimitError, OpenAI
from pypdf import PdfReader
import pandas as pd
from PIL import Image
from tenacity import retry, stop_after_attempt, wait_fixed, retry_if_exception_type
import msal
import requests
import structlog
from prometheus_client import Counter, generate_latest, CONTENT_TYPE_LATEST

logger = structlog.get_logger()

ORDER_ETA_DB = {
    "A-1042": {"status": "Packed and ready to ship", "eta": "2 business days", "carrier": "UPS"},
    "B-7711": {"status": "In transit", "eta": "5 business days", "carrier": "DHL"},
    "C-2209": {"status": "Awaiting carrier scan", "eta": "3 business days", "carrier": "FedEx"},
}

# Metrics counters
METRICS_PROCESSED = Counter("email_requests_processed_total", "Total number of email requests processed")
METRICS_REPLIED = Counter("email_replies_created_total", "Total number of replies/drafts created")
METRICS_FAILED = Counter("email_requests_failed_total", "Total number of failed email processing attempts")
METRICS_API_CALLS = Counter("msgraph_api_calls_total", "Total number of MS Graph API calls")

@dataclass
class RateLimiter:
    """Simple token bucket rate limiter for MS Graph API calls."""
    max_tokens: int
    refill_rate: float  # tokens per second
    tokens: float = field(default=None)
    last_refill: float = field(default=None)
    _condition: threading.Condition = field(default_factory=threading.Condition)

    def __post_init__(self):
        if self.tokens is None:
            self.tokens = float(self.max_tokens)
        if self.last_refill is None:
            self.last_refill = time.monotonic()
        if self.refill_rate <= 0:
            self.refill_rate = 1.0

    def _refill(self) -> None:
        now = time.monotonic()
        elapsed = now - self.last_refill
        tokens_to_add = elapsed * self.refill_rate
        self.tokens = min(self.max_tokens, self.tokens + tokens_to_add)
        self.last_refill = now

    def _wait_time_for_next_token(self) -> float:
        if self.tokens >= 1:
            return 0
        return max(0.01, (1 - self.tokens) / self.refill_rate)

    @contextmanager
    def acquire(self):
        """Acquire a token, waiting if necessary. Must be used as context manager."""
        with self._condition:
            self._refill()
            while self.tokens < 1:
                self._condition.wait(timeout=self._wait_time_for_next_token())
                self._refill()
            self.tokens -= 1
            METRICS_API_CALLS.inc()
            yield

def setup_logging() -> None:
    """Configure structured JSON logging."""
    structlog.configure(
        processors=[
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.add_log_level,
            structlog.processors.JSONRenderer()
        ],
        logger_factory=structlog.PrintLoggerFactory(),
        wrapper_class=structlog.make_filtering_bound_logger(logging.INFO),
    )
    logging.basicConfig(level=logging.INFO)

def load_rate_limiter_config() -> RateLimiter:
    """Load rate limiter configuration from environment."""
    return RateLimiter(
        max_tokens=int(os.getenv("MSGRAPH_RATE_LIMIT_TOKENS", "10")),
        refill_rate=float(os.getenv("MSGRAPH_RATE_LIMIT_REFILL_PER_SEC", "1.0")),
    )

# Global rate limiter instance
_rate_limiter: Optional[RateLimiter] = None

def get_rate_limiter() -> RateLimiter:
    """Get or create the global rate limiter."""
    global _rate_limiter
    if _rate_limiter is None:
        _rate_limiter = load_rate_limiter_config()
    return _rate_limiter

def rate_limited_request(method: str, url: str, **kwargs) -> requests.Response:
    """Wrapper for requests that enforces rate limiting on MS Graph API calls."""
    limiter = get_rate_limiter()
    with limiter.acquire():
        return requests.request(method, url, **kwargs)

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "lookup_eta",
            "description": "Look up the current ETA and delivery status for shipment identifiers such as orders, trucks, trailers, BOLs, or POs.",
            "parameters": {
                "type": "object",
                "properties": {
                    "orders": {
                        "type": "array",
                        "description": "List of order numbers to check.",
                        "items": {"type": "string"},
                    },
                    "trucks": {
                        "type": "array",
                        "description": "List of truck codes to check.",
                        "items": {"type": "string"},
                    },
                    "trailers": {
                        "type": "array",
                        "description": "List of trailer codes to check.",
                        "items": {"type": "string"},
                    },
                    "bols": {
                        "type": "array",
                        "description": "List of BOL numbers to check.",
                        "items": {"type": "string"},
                    },
                    "pos": {
                        "type": "array",
                        "description": "List of purchase order numbers to check.",
                        "items": {"type": "string"},
                    },
                },
                "required": [],
            },
        },
    }
]


def strip_html_tags(html_content: str) -> str:
    """Strip HTML tags and extract visible text, optimized for Outlook emails."""
    if not html_content:
        return ""
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Remove script and style elements
    for script in soup(["script", "style"]):
        script.decompose()
    
    # Remove hidden elements (common in Outlook)
    # Check for inline display:none styles
    for element in soup.find_all(style=re.compile(r"display\s*:\s*none", re.IGNORECASE)):
        element.decompose()
    
    # Also check for visibility:hidden
    for element in soup.find_all(style=re.compile(r"visibility\s*:\s*hidden", re.IGNORECASE)):
        element.decompose()
    
    # Remove elements with common hidden class names
    for element in soup.find_all(class_=re.compile(r"hidden", re.IGNORECASE)):
        element.decompose()
    
    # Get text
    text = soup.get_text()
    
    # Clean up whitespace
    lines = (line.strip() for line in text.splitlines())
    chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
    text = '\n'.join(chunk for chunk in chunks if chunk)
    
    return text


def extract_text_from_attachment(file_path: str, file_content: Optional[bytes] = None) -> str:
    """Extract text from various attachment formats."""
    if file_content:
        # Handle from bytes
        file_ext = os.path.splitext(file_path)[1].lower()
    else:
        # Handle from file path
        file_ext = os.path.splitext(file_path)[1].lower()
        try:
            with open(file_path, 'rb') as f:
                file_content = f.read()
        except Exception as e:
            logger.warning(f"Could not read attachment {file_path}: {e}")
            return ""
    
    try:
        if file_ext == '.pdf':
            return extract_text_from_pdf(file_content)
        elif file_ext in ['.docx', '.doc']:
            return extract_text_from_docx(file_content)
        elif file_ext in ['.xlsx', '.xls']:
            return extract_text_from_excel(file_content)
        elif file_ext == '.csv':
            return extract_text_from_csv(file_content)
        else:
            logger.warning(f"Unsupported attachment format: {file_ext}")
            return ""
    except Exception as e:
        logger.warning(f"Error extracting text from {file_path}: {e}")
        return ""


def extract_text_from_pdf(content: bytes) -> str:
    """Extract text from PDF using pypdf."""
    try:
        pdf_file = io.BytesIO(content)
        reader = PdfReader(pdf_file)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text.strip()
    except Exception as e:
        logger.warning(f"PDF extraction error: {e}")
        return ""


def extract_text_from_docx(content: bytes) -> str:
    """Extract text from DOCX using python-docx."""
    try:
        doc_file = io.BytesIO(content)
        doc = Document(doc_file)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text.strip()
    except Exception as e:
        logger.warning(f"DOCX extraction error: {e}")
        return ""


def extract_text_from_excel(content: bytes) -> str:
    """Extract text from Excel using pandas."""
    try:
        excel_file = io.BytesIO(content)
        df = pd.read_excel(excel_file)
        return df.to_string()
    except Exception as e:
        logger.warning(f"Excel extraction error: {e}")
        return ""


def extract_text_from_csv(content: bytes) -> str:
    """Extract text from CSV using pandas."""
    try:
        csv_file = io.BytesIO(content)
        df = pd.read_csv(csv_file)
        return df.to_string()
    except Exception as e:
        logger.warning(f"CSV extraction error: {e}")
        return ""


def parse_embedded_images(html_content: str) -> str:
    """Parse embedded images using Qwen 3.6 model for analysis."""
    if not html_content:
        return ""
    
    soup = BeautifulSoup(html_content, 'html.parser')
    images = soup.find_all('img')
    
    if not images:
        return ""
    
    image_descriptions = []
    client = OpenAI(
        api_key=os.getenv("OPENAI_API_KEY", "EMPTY"),
        base_url=os.getenv("VLLM_BASE_URL", "http://localhost:8000/v1"),
    )
    
    for img in images:
        try:
            src = img.get('src', '')
            if src.startswith('data:image'):
                header, encoded = src.split(',', 1)
                image_data = base64.b64decode(encoded)
                
                image = Image.open(io.BytesIO(image_data))
                
                if image.mode != 'RGB':
                    image = image.convert('RGB')
                
                max_size = 1024
                if max(image.size) > max_size:
                    image.thumbnail((max_size, max_size))
                
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='JPEG')
                img_byte_arr = img_byte_arr.getvalue()
                
                response = client.chat.completions.create(
                    model=os.getenv("VLLM_MODEL", "Qwen/Qwen3.6-27B"),
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": "Describe this image in detail, focusing on any text, numbers, order information, or shipping details visible."},
                                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64.b64encode(img_byte_arr).decode()}"}}
                            ]
                        }
                    ],
                    max_tokens=500
                )
                
                description = response.choices[0].message.content
                image_descriptions.append(f"[Image description: {description}]")
                
        except Exception as e:
            logger.warning(f"Error parsing image: {e}")
            continue
    
    return " ".join(image_descriptions) if image_descriptions else ""


def extract_latest_email_html(email_text: str) -> str:
    """Isolate the HTML of the latest email in a thread, decomposing older parts."""
    if not email_text or not ('<' in email_text and '>' in email_text):
        return email_text
        
    soup = BeautifulSoup(email_text, 'html.parser')
    
    # Decompose common Outlook reply headers / forward blocks
    div_rply = soup.find(id=re.compile(r"divRplyFwdMsg", re.IGNORECASE))
    if div_rply:
        for sibling in list(div_rply.find_next_siblings()) + [div_rply]:
            sibling.decompose()
            
    # Outlook border-top separator
    for div in soup.find_all('div', style=True):
        style = div['style'].lower()
        if 'border-top:solid' in style or 'border-top: solid' in style:
            for sibling in list(div.find_next_siblings()) + [div]:
                sibling.decompose()
            break
            
    # Standard horizontal rule separator
    hr = soup.find('hr')
    if hr:
        for sibling in list(hr.find_next_siblings()) + [hr]:
            sibling.decompose()
            
    return str(soup)


def extract_latest_email_text(text: str) -> str:
    """Isolate the latest message in plain text by cutting off at thread headers."""
    if not text:
        return ""
        
    lines = text.splitlines()
    cleaned_lines = []
    
    # Regex to match common headers
    original_message_regex = re.compile(
        r"(-----Original Message-----|________________________________|On\s+.*\s+wrote:|^From:\s+|^Sent:\s+|^To:\s+|^Subject:\s+)",
        re.IGNORECASE
    )
    
    for line in lines:
        if original_message_regex.search(line):
            logger.info(f"Thread separator matched in plain text: '{line.strip()}'. Truncating thread.")
            break
        cleaned_lines.append(line)
        
    return "\n".join(cleaned_lines).strip()


def build_system_prompt() -> str:
    return (
        "You are an order ETA assistant for a shared inbox. You will receive HTML Outlook emails. "
        "Only answer when the email is clearly an ETA or delivery timing request. "
        "Use your judgment to identify shipment references such as order numbers, truck codes, trailer codes, "
        "Bill of Lading (BOL), Purchase Orders (PO), and invoice references from the text. "
        "The exact format can vary, so inspect the visible text carefully and extract any shipment identifiers you can find. "
        "When you call the lookup_eta tool, pass the identifiers in the appropriate list fields: orders, trucks, trailers, bols, and pos. "
        "If the email is not an ETA request, do nothing and return an empty response. "
        "Write a friendly, concise email reply that references the matched identifiers and ETA details."
    )


def lookup_eta(order_id: str | None = None, *, orders=None, trucks=None, trailers=None, bols=None, pos=None) -> dict:
    identifiers = [order_id] if order_id else []
    identifiers.extend(orders or [])
    identifiers.extend(trucks or [])
    identifiers.extend(trailers or [])
    identifiers.extend(bols or [])
    identifiers.extend(pos or [])

    if not identifiers:
        return {"status": "No shipment identifiers provided", "eta": "Unknown", "carrier": "Unknown"}

    normalized = [value.upper().strip() for value in identifiers if value]
    for value in normalized:
        if value in ORDER_ETA_DB:
            return ORDER_ETA_DB[value]

    return {"status": "Identifier not found in the demo catalog", "eta": "3-5 business days", "carrier": "Carrier update pending"}


def get_retry_config():
    """Get retry configuration from environment variables."""
    return {
        "max_retries": int(os.getenv("MAX_RETRIES", "2")),
        "retry_delay": int(os.getenv("RETRY_DELAY_SECONDS", "60"))
    }


def call_llm_api(client, email_text):
    """Make the LLM API call with retry logic."""
    config = get_retry_config()
    
    @retry(
        stop=stop_after_attempt(config["max_retries"] + 1),
        wait=wait_fixed(config["retry_delay"]),
        retry=retry_if_exception_type((APIConnectionError, APIError, RateLimitError)),
        before_sleep=lambda retry_state: logger.warning("Retry attempt after error", attempt=retry_state.attempt_number, error=str(retry_state.outcome.exception()))
    )
    def _make_api_call():
        response = client.chat.completions.create(
            model=os.getenv("VLLM_MODEL", "Qwen/Qwen3.6-27B"),
            messages=[
                {
                    "role": "system",
                    "content": build_system_prompt(),
                },
                {
                    "role": "user",
                    "content": f"Please answer this shared-inbox ETA request. Email text: {email_text}",
                },
            ],
            tools=TOOLS,
            tool_choice="auto",
        )
        return response.choices[0].message
    
    return _make_api_call()


def answer_email(email_text: str, attachments: Optional[list] = None) -> str:
    # Handle HTML vs Plain Text email thread processing
    if '<' in email_text and '>' in email_text:
        latest_html = extract_latest_email_html(email_text)
        processed_text = strip_html_tags(latest_html)
        # Apply secondary text-based thread truncation to clean up any leftover text separators
        processed_text = extract_latest_email_text(processed_text)
        # Analyze embedded images only from the latest message
        image_descriptions = parse_embedded_images(latest_html)
    else:
        processed_text = extract_latest_email_text(email_text)
        image_descriptions = ""


    # Extract text from attachments if provided
    attachment_texts = []
    if attachments:
        for attachment in attachments:
            if isinstance(attachment, dict):
                file_path = attachment.get("file_path", "")
                file_content = attachment.get("file_content")
            else:
                file_path = attachment
                file_content = None
            extracted = extract_text_from_attachment(file_path, file_content)
            if extracted:
                attachment_texts.append(extracted)

    # Combine all text sources
    combined_text = processed_text
    if image_descriptions:
        combined_text += f"\n\n{image_descriptions}"
    if attachment_texts:
        combined_text += f"\n\nAttachment contents:\n" + "\n".join(attachment_texts)

    client = OpenAI(
        api_key=os.getenv("OPENAI_API_KEY", "EMPTY"),
        base_url=os.getenv("VLLM_BASE_URL", "http://localhost:8000/v1"),
    )

    try:
        message = call_llm_api(client, combined_text)

        if getattr(message, "tool_calls", None):
            reply_parts = []
            for tool_call in message.tool_calls:
                if tool_call.function.name == "lookup_eta":
                    args = json.loads(tool_call.function.arguments)
                    eta_info = lookup_eta(
                        orders=args.get("orders"),
                        trucks=args.get("trucks"),
                        trailers=args.get("trailers"),
                        bols=args.get("bols"),
                        pos=args.get("pos"),
                    )
                    reply_parts.append(
                        f"I checked the referenced shipment details and found "
                        f"status is {eta_info['status']}; ETA is {eta_info['eta']}; carrier is {eta_info['carrier']}."
                    )
            return " ".join(reply_parts) if reply_parts else "I could not verify the ETA from the model response."

        return message.content or "I could not generate a reply."

    except (APIConnectionError, APIError, RateLimitError) as e:
        logger.error(f"API error after retries: {e}")
        return ""


def get_ms_graph_token(
    tenant_id: str,
    client_id: str,
    client_secret: str,
) -> str:
    """Acquire an access token for Microsoft Graph using MSAL with Client Credentials flow."""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    
    logger.info("Authenticating via ConfidentialClientApplication (Client Credentials flow)...")
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )
    # For client credentials, the scope is always https://graph.microsoft.com/.default
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    else:
        error_msg = result.get("error_description") or result.get("error", "Unknown error")
        raise ValueError(f"Failed to acquire token via client credentials: {error_msg}")


def process_msgraph_inbox(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    user_email: str,
    auto_reply: bool = False,
    mark_read: bool = False,
    create_draft: bool = True,
    limit: int = 10,
    time_window_minutes: int = 5,
) -> None:
    """Read unread messages from MS Graph inbox, run them through LLM, and reply/draft if configured."""
    setup_logging()
    
    try:
        token = get_ms_graph_token(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )
    except Exception as e:
        logger.error("Authentication failed", error=str(e))
        METRICS_FAILED.inc()
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    base_url = f"https://graph.microsoft.com/v1.0/users/{user_email}"
    
    if mark_read:
        filter_query = "isRead eq false"
    else:
        filter_query = "isRead eq false and not(categories/any(c:c eq 'AgentDrafted'))"
    
    # Add time window filter for emails received in the last X minutes
    if time_window_minutes and time_window_minutes > 0:
        from datetime import datetime, timezone, timedelta
        cutoff_time = datetime.now(timezone.utc) - timedelta(minutes=time_window_minutes)
        cutoff_iso = cutoff_time.strftime("%Y-%m-%dT%H:%M:%SZ")
        filter_query = f"{filter_query} and receivedDateTime ge {cutoff_iso}"

    params = {
        "$filter": filter_query,
        "$orderby": "receivedDateTime desc",
        "$top": limit
    }

    inbox_url = f"{base_url}/mailFolders/inbox/messages"

    logger.info("Fetching unread messages", url=inbox_url)
    try:
        response = rate_limited_request("GET", inbox_url, headers=headers, params=params)
        if response.status_code != 200:
            logger.error("Failed to fetch messages", status_code=response.status_code, response_text=response.text)
            METRICS_FAILED.inc()
            return
        
        messages = response.json().get("value", [])
        logger.info("Found unread messages", count=len(messages))
        
        for msg in messages:
            msg_id = msg.get("id")
            subject = msg.get("subject", "(No Subject)")
            sender = msg.get("from", {}).get("emailAddress", {}).get("address", "Unknown Sender")
            received = msg.get("receivedDateTime", "Unknown Time")
            
            logger.info("Processing message", msg_id=msg_id, subject=subject, sender=sender, received=received)
            METRICS_PROCESSED.inc()

            body_content = msg.get("body", {}).get("content", "")
            
            attachments = []
            if msg.get("hasAttachments"):
                attachments_url = f"{base_url}/messages/{msg_id}/attachments"
                logger.info("Fetching attachments", msg_id=msg_id)
                att_resp = rate_limited_request("GET", attachments_url, headers=headers)
                if att_resp.status_code == 200:
                    for att in att_resp.json().get("value", []):
                        if att.get("@odata.type") == "#microsoft.graph.fileAttachment" and "contentBytes" in att:
                            try:
                                content_bytes = base64.b64decode(att["contentBytes"])
                                attachments.append({
                                    "file_path": att.get("name", "attachment"),
                                    "file_content": content_bytes
                                })
                                logger.info("Loaded attachment", name=att.get('name'))
                            except Exception as e:
                                logger.warning("Failed to decode attachment", name=att.get('name'), error=str(e))
                else:
                    logger.warning("Failed to retrieve attachments", status_code=att_resp.status_code, response_text=att_resp.text)

            logger.info("Passing email content to LLM agent", msg_id=msg_id)
            answer = answer_email(body_content, attachments=attachments)
            
            if answer:
                logger.info("LLM Reply generated", msg_id=msg_id, reply_preview=answer[:100])
                print(f"--- REPLY START ---\n{answer}\n--- REPLY END ---")
                
                if auto_reply:
                    if create_draft:
                        logger.info("Creating draft reply", msg_id=msg_id)
                        reply_url = f"{base_url}/messages/{msg_id}/createReply"
                        reply_payload = {"comment": answer}
                        reply_resp = rate_limited_request("POST", reply_url, headers=headers, json=reply_payload)
                        if reply_resp.status_code in [200, 201, 202]:
                            logger.info("Draft reply created successfully", msg_id=msg_id)
                            METRICS_REPLIED.inc()
                        else:
                            logger.error("Failed to create draft reply", msg_id=msg_id, status_code=reply_resp.status_code, response_text=reply_resp.text)
                            METRICS_FAILED.inc()
                    else:
                        logger.info("Sending reply directly", msg_id=msg_id)
                        reply_url = f"{base_url}/messages/{msg_id}/reply"
                        reply_payload = {"comment": answer}
                        reply_resp = rate_limited_request("POST", reply_url, headers=headers, json=reply_payload)
                        if reply_resp.status_code in [200, 201, 202]:
                            logger.info("Reply sent successfully", msg_id=msg_id)
                            METRICS_REPLIED.inc()
                        else:
                            logger.error("Failed to send reply", msg_id=msg_id, status_code=reply_resp.status_code, response_text=reply_resp.text)
                            METRICS_FAILED.inc()
                else:
                    logger.info("Auto-reply disabled, skipping reply/draft creation", msg_id=msg_id)
            else:
                logger.info("No reply generated - email did not match ETA criteria or returned empty response", msg_id=msg_id)

            if mark_read:
                logger.info("Marking message as read", msg_id=msg_id)
                update_url = f"{base_url}/messages/{msg_id}"
                update_payload = {"isRead": True}
                update_resp = rate_limited_request("PATCH", update_url, headers=headers, json=update_payload)
                if update_resp.status_code == 200:
                    logger.info("Message marked as read", msg_id=msg_id)
                else:
                    logger.error("Failed to mark message as read", msg_id=msg_id, status_code=update_resp.status_code, response_text=update_resp.text)
            else:
                logger.info("Categorizing message as AgentDrafted", msg_id=msg_id)
                existing_categories = msg.get("categories", [])
                if "AgentDrafted" not in existing_categories:
                    new_categories = existing_categories + ["AgentDrafted"]
                    update_url = f"{base_url}/messages/{msg_id}"
                    update_payload = {"categories": new_categories}
                    update_resp = rate_limited_request("PATCH", update_url, headers=headers, json=update_payload)
                    if update_resp.status_code == 200:
                        logger.info("Message successfully categorized as AgentDrafted", msg_id=msg_id)
                    else:
                        logger.error("Failed to add category", msg_id=msg_id, status_code=update_resp.status_code, response_text=update_resp.text)

    except Exception as e:
        logger.error("Error processing inbox", error=str(e), exc_info=True)
        METRICS_FAILED.inc()


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate an order ETA email reply or process MS Graph inbox.")
    
    parser.add_argument("email_text", nargs="?", default=None, help="Email text to answer (ignored if using MS Graph mode).")
    
    parser.add_argument("--msgraph", action="store_true", help="Enable Microsoft Graph inbox processing mode.")
    parser.add_argument("--tenant-id", default=os.getenv("MSGRAPH_TENANT_ID"), help="Microsoft Entra tenant ID.")
    parser.add_argument("--client-id", default=os.getenv("MSGRAPH_CLIENT_ID"), help="Client (application) ID.")
    parser.add_argument("--client-secret", default=os.getenv("MSGRAPH_CLIENT_SECRET"), help="Client secret (required).")
    parser.add_argument("--user-email", default=os.getenv("MSGRAPH_USER_EMAIL"), help="Email/UPN of the inbox to process (required).")
    parser.add_argument("--auto-reply", action="store_true", default=os.getenv("MSGRAPH_AUTO_REPLY", "False").lower() in ("true", "1", "yes"), help="Automatically generate draft or send replies.")
    parser.add_argument("--create-draft", action="store_true", default=os.getenv("MSGRAPH_CREATE_DRAFT", "True").lower() in ("true", "1", "yes"), help="Create a draft reply instead of sending directly. (Default: True).")
    parser.add_argument("--no-draft", action="store_false", dest="create_draft", help="Send replies directly instead of creating a draft.")
    parser.add_argument("--mark-read", action="store_true", default=os.getenv("MSGRAPH_MARK_AS_READ", "False").lower() in ("true", "1", "yes"), help="Mark processed emails as read. (Default: False, keeps them unread and tags as 'AgentDrafted').")
    parser.add_argument("--limit", type=int, default=10, help="Maximum number of messages to process.")
    parser.add_argument("--time-window-minutes", type=int, default=int(os.getenv("MSGRAPH_TIME_WINDOW_MINUTES", "5")), help="Only process emails received in the last X minutes (default: 5, 0 for no time filter).")
    parser.add_argument("--health-port", type=int, default=int(os.getenv("HEALTH_PORT", "8080")), help="Port for health check HTTP endpoint.")
    parser.add_argument("--daemon", action="store_true", help="Run in daemon mode, continuously processing inbox.")
    parser.add_argument("--poll-interval", type=int, default=int(os.getenv("POLL_INTERVAL_SECONDS", "60")), help="Polling interval in daemon mode (seconds).")
    
    args = parser.parse_args()
    
    if args.msgraph:
        if not args.tenant_id or not args.client_id or not args.client_secret or not args.user_email:
            parser.error("--tenant-id, --client-id, --client-secret, and --user-email are required for MS Graph mode (can also be set via env variables).")
        
        setup_logging()
        logger.info("Starting MS Graph email agent")
        
        from flask import Flask, Response
        
        app = Flask(__name__)
        
        @app.route("/health")
        def health_check():
            return Response(generate_latest(), mimetype=CONTENT_TYPE_LATEST)
        
        def run_daemon():
            while True:
                try:
                    process_msgraph_inbox(
                        tenant_id=args.tenant_id,
                        client_id=args.client_id,
                        client_secret=args.client_secret,
                        user_email=args.user_email,
                        auto_reply=args.auto_reply,
                        mark_read=args.mark_read,
                        create_draft=args.create_draft,
                        limit=args.limit,
                        time_window_minutes=args.time_window_minutes
                    )
                except Exception as e:
                    logger.error("Daemon iteration failed", error=str(e))
                time.sleep(args.poll_interval)
        
        daemon_thread = threading.Thread(target=run_daemon, daemon=True)
        daemon_thread.start()
        
        logger.info(f"Starting health endpoint on port {args.health_port}")
        app.run(host="0.0.0.0", port=args.health_port)
    else:
        setup_logging()
        email_text = args.email_text or "Hi, when will order A-1042 arrive?"
        answer = answer_email(email_text)
        print("\nReply:\n")
        print(answer)


if __name__ == "__main__":
    main()

