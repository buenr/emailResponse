# AGENTS.md

## Run Commands

```powershell
# Quick test with email text argument
python EmailResponse.py "Hi, when will my order ABC1234 arrive?"

# MS Graph mode (Client Credentials Flow - production)
python EmailResponse.py --msgraph --tenant-id "your-tenant-id" --client-id "your-client-id" --client-secret "your-secret" --user-email "shared-inbox@company.com" --auto-reply --daemon --health-port 8080
```

## Test

```bash
python -m pytest test_email_agent.py -v
```

## Environment Variables

| Variable | Default | Purpose |
|----------|---------|---------|
| `VLLM_BASE_URL` | `http://localhost:8000/v1` | LLM endpoint |
| `VLLM_MODEL` | `Qwen/Qwen3.6-27B` | Model name |
| `OPENAI_API_KEY` | `EMPTY` | API key |
| `MAX_RETRIES` | `2` | Retry attempts |
| `RETRY_DELAY_SECONDS` | `60` | Delay between retries |
| `MSGRAPH_TENANT_ID` | - | Graph tenant |
| `MSGRAPH_CLIENT_ID` | - | Graph client ID |
| `MSGRAPH_CLIENT_SECRET` | - | Graph client secret (required for Client Credentials flow) |
| `MSGRAPH_USER_EMAIL` | - | Mailbox to access (required) |
| `MSGRAPH_AUTO_REPLY` | `False` | Enable reply/draft |
| `MSGRAPH_CREATE_DRAFT` | `True` | Create draft vs send |
| `MSGRAPH_MARK_AS_READ` | `False` | Mark read vs tag `AgentDrafted` |
| `MSGRAPH_RATE_LIMIT_TOKENS` | `10` | Max concurrent Graph API requests |
| `MSGRAPH_RATE_LIMIT_REFILL_PER_SEC` | `1.0` | Token refill rate per second |
| `HEALTH_PORT` | `8080` | Port for /health endpoint |
| `POLL_INTERVAL_SECONDS` | `60` | Polling interval for daemon mode |
| `MSGRAPH_TIME_WINDOW_MINUTES` | `5` | Only process emails received in last X minutes (0 for no filter) |

## MS Graph Quirks

- **Time filter**: Default filters to emails received in the last 5 minutes (`MSGRAPH_TIME_WINDOW_MINUTES=5`). Set to 0 to disable.
- **Client Credentials flow**: Requires `--user-email` and `--client-secret`; uses `Mail.ReadWrite`, `Mail.Send` (admin consent)
- **Filter**: If `mark_read=False` (default), filters out messages with `AgentDrafted` category to avoid reprocessing unread mail
- **Reply endpoints**: `createReply` (draft) vs `reply` (send)

## Testing Notes

- Tests mock MSAL, requests, and OpenAI client
- Run single test: `python -m pytest test_email_agent.py::EmailAgentTests::test_strip_html_tags -v`
- All external calls mocked; no network required