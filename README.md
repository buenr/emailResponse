# Email ETA Agent

This repo now contains a small order-ETA reply agent.

## Run it

```powershell
"EmailResponse.py" "Hi, when will my order A-1042 arrive?"
```

## Notes

- The current implementation uses a demo order ETA table for A-1042, B-7711, and C-2209.
- If a local vLLM server is available, the script will try to use it first.
- The agent should rely on the LLM path for all ETA responses; no deterministic fallback reply logic is used.
- The agent handles HTML Outlook emails by stripping tags and extracting visible text.
- The agent isolates the latest reply in email threads (supporting both HTML and plain text separators) to ignore historical messages and prevent false-positive ETA triggers.
- The agent can extract text from email attachments (PDF, DOCX, Excel, CSV).

- The agent can parse embedded images in HTML emails using the vision model.
- The agent implements retry logic for API calls to handle transient errors.
- MS Graph calls are protected by a configurable token-bucket rate limiter to reduce 429 throttling.
- The daemon exposes Prometheus metrics through `/health` and logs structured JSON events.

## Environment Variables

To point the script at a different model endpoint, set:

```powershell
$env:VLLM_BASE_URL="http://localhost:8000/v1"
$env:VLLM_MODEL="Qwen/Qwen3.6-27B"
```

To configure retry behavior, set:

```powershell
$env:MAX_RETRIES="2"
$env:RETRY_DELAY_SECONDS="60"
```

- `MAX_RETRIES`: Number of retry attempts for API calls (default: 2)
- `RETRY_DELAY_SECONDS`: Delay between retries in seconds (default: 60)
- `MSGRAPH_RATE_LIMIT_TOKENS`: Maximum MS Graph token-bucket capacity (default: 10)
- `MSGRAPH_RATE_LIMIT_REFILL_PER_SEC`: MS Graph token refill rate per second (default: 1.0)
- `HEALTH_PORT`: Port for the `/health` Prometheus metrics endpoint (default: 8080)
- `POLL_INTERVAL_SECONDS`: Daemon polling interval in seconds (default: 60)

## Microsoft Graph Mode

The agent can read unread emails from an Outlook inbox using the Microsoft Graph API, pass them through the LLM, print/log the generated reply, and optionally reply to the sender and mark the email as read.

### Execution

```bash
# Run in MS Graph mode (interactive Device Code Flow, default)
python3 EmailResponse.py --msgraph --tenant-id "your-tenant-id" --client-id "your-client-id"

# Run in MS Graph mode (unattended/daemon Client Credentials Flow)
python3 EmailResponse.py --msgraph --tenant-id "your-tenant-id" --client-id "your-client-id" --client-secret "your-secret" --user-email "shared-inbox@company.com" --auto-reply --daemon --health-port 8080
```

### Authentication Flows Supported

1. **Client Credentials Flow (Daemon App):** If you provide `--client-secret` (or set `MSGRAPH_CLIENT_SECRET`), the script will authenticate as a daemon app/service principal. Note: `--user-email` is required in this flow to specify whose mailbox to access.
2. **Device Code Flow (Interactive User):** If no client secret is provided (or if `--device-code` is passed), the script will prompt with a device code link for authentication. Serialized tokens are cached in `token_cache.bin` to prevent repeated prompts.

### Command-line Arguments

- `--msgraph`: Enables Microsoft Graph inbox processing.
- `--tenant-id`: Microsoft Entra Tenant ID (can be set via `MSGRAPH_TENANT_ID`).
- `--client-id`: Client/Application ID (can be set via `MSGRAPH_CLIENT_ID`).
- `--client-secret`: Client Secret (can be set via `MSGRAPH_CLIENT_SECRET`).
- `--user-email`: Email/UPN of the inbox to process (can be set via `MSGRAPH_USER_EMAIL`).
- `--device-code`: Force MSAL Device Code Flow.
- `--auto-reply`: Automatically generate draft or send replies (can be set via `MSGRAPH_AUTO_REPLY`).
- `--create-draft`: Create a draft reply instead of sending directly. (Default: True, can be set via `MSGRAPH_CREATE_DRAFT`).
- `--no-draft`: Send replies directly instead of creating drafts.
- `--mark-read`: Mark processed emails as read. (Default: False, keeps them unread and tags them with the 'AgentDrafted' category to avoid duplicate processing. Can be set via `MSGRAPH_MARK_AS_READ`).
- `--limit`: Max number of unread emails to process in a single run (default: 10).
- `--health-port`: Port for the `/health` Prometheus metrics endpoint (default: 8080).
- `--daemon`: Run continuously and poll the inbox on a fixed interval.
- `--poll-interval`: Polling interval for daemon mode in seconds (default: 60).


