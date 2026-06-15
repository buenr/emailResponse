# Agent Guide

This agent is designed for a shared inbox and should handle HTML Outlook emails.

## What it must understand

1. Incoming email content may be HTML, not plain text.
2. It should strip tags, decode visible text, and ignore marketing banners or boilerplate where possible.
3. It should only answer when the email is clearly an ETA or delivery-status request.
4. It should identify shipment references from the email, including:
   - Order numbers: 3 letters + 4 digits (for example ABC1234)
   - Truck codes: 6 alphanumeric chars (for example ABC123)
   - Trailer codes: 6 alphanumeric chars (for example DEF456)
   - BOLs: 8 to 25 alphanumeric chars, often prefixed BOL-
   - POs: 8 to 25 alphanumeric chars, often prefixed PO-
   - Invoices: common forms such as INV-55 or INV123
5. It should call the ETA lookup tool with a list-based payload:
   - orders
   - trucks
   - trailers
   - bols
   - pos
6. It should extract text from email attachments if present, supporting:
   - PDF files (using pypdf)
   - DOCX files (using python-docx)
   - Excel files (using pandas and openpyxl)
   - CSV files (using pandas)
7. It should parse embedded images in HTML emails and analyze them using the vision model to extract any visible text, numbers, or shipping details.

## Tool contract

The ETA lookup tool expects a JSON object with any of the following arrays:

```json
{
  "orders": ["ABC1234"],
  "trucks": ["ABC123"],
  "trailers": ["DEF456"],
  "bols": ["BOL-ABC12345"],
  "pos": ["PO-ABC12345"]
}
```

The agent should prefer the identifiers actually present in the email and use them to build the reply.

## Retry Logic

The agent implements retry logic for API calls to handle transient errors:
- Default: 2 retry attempts with 60-second fixed delay
- Configurable via environment variables: `MAX_RETRIES` and `RETRY_DELAY_SECONDS`

## Production Hardening

The MS Graph workflow includes production hardening:
- A token-bucket rate limiter protects Graph API calls from 429 throttling.
- Prometheus counters track processed, replied, and failed requests.
- Structured JSON logging is enabled for debugging runs.
- Daemon mode exposes `/health` with Prometheus metrics.
- Retries on: APIConnectionError, APIError, RateLimitError

## MS Graph Integration

The agent supports reading, processing, and replying to messages in an Outlook inbox via Microsoft Graph:
- **Scope Permissions:**
  - Client Credentials (Application permissions): `Mail.ReadWrite`, `Mail.Send` (admin consent required).
  - Device Code (Delegated permissions): `Mail.ReadWrite`, `Mail.Send` (user consent).
- **Inbox Processing Workflow:**
  - Fetch unread messages: `$filter=isRead eq false` (standard) or `$filter=isRead eq false and not(categories/any(c:c eq 'AgentDrafted'))` (if mark_read is disabled).
  - Isolate the latest reply in email threads (both HTML and plain text separators) to ignore historical conversation details.
  - Fetch attachments for emails flagged with `hasAttachments eq true`
  - Pass the message to the LLM to determine if it is an ETA request and generate a response.
  - Draft/Reply Action (if auto_reply is enabled):
    - If `create_draft` is True (default): Create a draft reply in Outlook (POST to `/createReply` endpoint).
    - If `create_draft` is False: Send reply directly back to the sender (POST to `/reply` endpoint).
  - Mark as processed:
    - If `mark_read` is True: Mark the message as read (PATCH message with `{"isRead": true}`).
    - If `mark_read` is False (default): Tag message with the `AgentDrafted` category to keep it unread in the inbox but avoid duplicate processing.



