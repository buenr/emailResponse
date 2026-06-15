# Email ETA Agent

Email ETA Agent is a Python automation tool for shared Outlook inboxes. It reads incoming email requests, isolates the latest message in a thread, extracts text from attachments, uses an OpenAI-compatible LLM endpoint to determine whether the message is an order ETA request, and optionally creates a draft or sends a reply through Microsoft Graph.

The project is designed for teams that need a lightweight, script-based email assistant with configurable authentication, retry behavior, rate limiting, structured logging, and Prometheus-compatible health metrics.

## Features

- Processes HTML and plain-text Outlook emails
- Isolates the latest message in email threads
- Extracts text from PDF, DOCX, Excel, and CSV attachments
- Parses embedded HTML images through an OpenAI-compatible vision model
- Uses a configurable LLM tool-calling workflow for ETA lookups
- Supports Microsoft Graph mailbox access through Client Credentials Flow for daemon/service-principal authentication
- Creates draft replies or sends replies directly
- Marks messages as read or tags them with `AgentDrafted` to avoid duplicate processing
- Applies token-bucket rate limiting to Microsoft Graph API calls
- Exposes Prometheus-compatible metrics through `/health` in daemon mode
- Provides a Dockerfile and GitLab CI template for OpenShift container deployments
- Uses structured JSON logging for production debugging

## Requirements

- Python 3.11+
- Microsoft Entra app registration with Microsoft Graph permissions
- OpenAI-compatible API endpoint, such as a local vLLM server

## Installation

```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # macOS/Linux

pip install -r requirements.txt
```

## Quick Test

Run the local ETA demo without Microsoft Graph:

```bash
python EmailResponse.py "Hi, when will my order A-1042 arrive?"
```

## Microsoft Graph Setup

Create an app registration in Microsoft Entra ID and grant Microsoft Graph permissions:

### Client Credentials Flow

Use this mode for unattended daemon operation.

Required application permissions:

- `Mail.ReadWrite`
- `Mail.Send`

Admin consent is required.

## Usage

### Generate a Reply Locally

```bash
python EmailResponse.py "Hi, when will my order A-1042 arrive?"
```

### Process Microsoft Graph Inbox Once

```bash
python EmailResponse.py \
  --msgraph \
  --tenant-id "your-tenant-id" \
  --client-id "your-client-id" \
  --client-secret "your-secret" \
  --user-email "shared-inbox@company.com"
```

### Run in Daemon Mode

```bash
python EmailResponse.py \
  --msgraph \
  --tenant-id "your-tenant-id" \
  --client-id "your-client-id" \
  --client-secret "your-secret" \
  --user-email "shared-inbox@company.com" \
  --auto-reply \
  --daemon \
  --health-port 8080
```

In daemon mode, the process polls the inbox on a fixed interval and serves Prometheus-compatible metrics at:

```text
http://localhost:8080/health
```

### Container/OpenShift Mode

The included Dockerfile runs the application with Gunicorn on port `8080`. Set `EMAIL_AGENT_DAEMON=true` plus the Microsoft Graph environment variables to run inbox polling from the container health endpoint process:

```bash
docker build -t email-response-agent .
docker run --rm -p 8080:8080 --env-file .env email-response-agent
```

The GitLab CI template builds the image, pushes it to the GitLab registry, and includes a manual OpenShift deployment job that updates or creates `deployment/email-response-agent`.

## Configuration

Copy `.env.example` to `.env` and update the values for your environment:

```bash
cp .env.example .env
```

### LLM Configuration

| Variable | Default | Description |
|---|---:|---|
| `VLLM_BASE_URL` | `http://localhost:8000/v1` | OpenAI-compatible API base URL |
| `VLLM_MODEL` | `Qwen/Qwen3.6-27B` | Model name used for email processing |
| `PORT` | `8080` | Container/OpenShift health endpoint port |
| `OPENAI_API_KEY` | `EMPTY` | API key for the OpenAI-compatible endpoint |
| `MAX_RETRIES` | `2` | Number of retry attempts for LLM API failures |
| `RETRY_DELAY_SECONDS` | `60` | Delay between LLM retry attempts |

### Microsoft Graph Configuration

| Variable | Default | Description |
|---|---:|---|
| `MSGRAPH_TENANT_ID` | Required | Microsoft Entra tenant ID |
| `MSGRAPH_CLIENT_ID` | Required | Application/client ID |
| `MSGRAPH_CLIENT_SECRET` | Required | Client secret for Client Credentials flow |
| `MSGRAPH_USER_EMAIL` | Required | Mailbox UPN or email address to process |
| `MSGRAPH_AUTO_REPLY` | `False` | Create drafts or send replies automatically |
| `MSGRAPH_CREATE_DRAFT` | `True` | Create draft replies instead of sending directly |
| `MSGRAPH_MARK_AS_READ` | `False` | Mark processed messages as read |
| `MSGRAPH_TIME_WINDOW_MINUTES` | `5` | Only process emails received in last X minutes (0 for no filter) |

### Production Hardening

| Variable | Default | Description |
|---|---:|---|
| `MSGRAPH_RATE_LIMIT_TOKENS` | `10` | Token-bucket capacity for Microsoft Graph API calls |
| `MSGRAPH_RATE_LIMIT_REFILL_PER_SEC` | `1.0` | Token refill rate per second |
| `HEALTH_PORT` | `8080` | Port for the `/health` Prometheus metrics endpoint |
| `POLL_INTERVAL_SECONDS` | `60` | Daemon polling interval in seconds |
| `EMAIL_AGENT_DAEMON` | `False` | Start background inbox polling when the container/Wsgi app starts |
| `EMAIL_AGENT_AUTO_REPLY` | `False` | Container override for automatic draft/reply creation |
| `EMAIL_AGENT_CREATE_DRAFT` | `True` | Container override for draft-vs-send behavior |
| `EMAIL_AGENT_MARK_READ` | `False` | Container override for read-vs-category processing |
| `EMAIL_AGENT_LIMIT` | `10` | Container override for messages processed per polling iteration |
| `EMAIL_AGENT_TIME_WINDOW_MINUTES` | `5` | Container override for received-time filter |
| `EMAIL_AGENT_POLL_INTERVAL_SECONDS` | `60` | Container override for polling interval |

## Command-Line Options

| Option | Description |
|---|---|
| `--msgraph` | Enable Microsoft Graph inbox processing |
| `--tenant-id` | Microsoft Entra tenant ID |
| `--client-id` | Microsoft Graph application/client ID |
| `--client-secret` | Client secret for Client Credentials flow |
| `--user-email` | Mailbox UPN or email address to process |
| `--auto-reply` | Create draft replies or send replies automatically |
| `--create-draft` | Create draft replies instead of sending directly |
| `--no-draft` | Send replies directly instead of creating drafts |
| `--mark-read` | Mark processed emails as read |
| `--time-window-minutes` | Only process emails received in last X minutes (default: 5) |
| `--limit` | Maximum number of unread messages to process per run |
| `--health-port` | Port for the `/health` endpoint |
| `--daemon` | Continuously poll the inbox |
| `--poll-interval` | Polling interval for daemon mode |

## Processing Behavior

When Microsoft Graph processing is enabled, the agent:

1. Authenticates with Microsoft Graph using Client Credentials flow.
2. Fetches unread inbox messages received in the last 5 minutes (configurable).
2. Fetches unread inbox messages.
3. Filters out messages already tagged with `AgentDrafted` when `mark_read` is disabled.
4. Fetches attachments for messages that contain them.
5. Extracts the latest message from the email thread.
6. Passes the cleaned email content and attachment text to the LLM.
7. Generates a reply when the email is identified as an ETA or delivery-status request.
8. Creates a draft, sends a reply, marks the message as read, or tags it as `AgentDrafted`, depending on configuration.

## Observability

The daemon exposes Prometheus-compatible counters through `/health`:

- `email_requests_processed_total`
- `email_replies_created_total`
- `email_requests_failed_total`
- `msgraph_api_calls_total`

Logs are emitted as structured JSON events with timestamps, severity, and contextual fields.

## Testing

Run the test suite:

```bash
python -m pytest test_email_agent.py -v
```

Run a single test:

```bash
python -m pytest test_email_agent.py::EmailAgentTests::test_strip_html_tags -v
```

Tests mock Microsoft Graph, MSAL, requests, and OpenAI client calls. No external network access is required.

## Security Notes

- Do not commit `.env`, credentials, or client secrets.
- `.gitignore` excludes local virtual environments, Python caches, pytest caches, and token cache files.
- Use client secrets and tokens only through environment variables or a secure secret manager.
- Prefer least-privilege Microsoft Graph permissions for production deployments.

## License

This repository does not declare a license. Add an explicit license before distributing the project outside your organization.
