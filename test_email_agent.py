import unittest
from unittest.mock import patch, MagicMock
import io

import EmailResponse


class EmailAgentTests(unittest.TestCase):

    def test_system_prompt_describes_email_and_identifier_rules(self):
        prompt = EmailResponse.build_system_prompt()
        self.assertIn("HTML Outlook emails", prompt)
        self.assertIn("order numbers", prompt)
        self.assertIn("Bill of Lading", prompt)
        self.assertIn("lookup_eta tool", prompt)
        self.assertIn("orders, trucks, trailers, bols, and pos", prompt)

    def test_strip_html_tags(self):
        html = "<html><body><p>Hi, when will order A-1042 arrive?</p></body></html>"
        text = EmailResponse.strip_html_tags(html)
        self.assertIn("order A-1042", text)
        self.assertNotIn("<html>", text)
        self.assertNotIn("<p>", text)

    def test_strip_html_tags_with_outlook_structure(self):
        html = """
        <html>
        <head><style>.hidden {display: none;}</style></head>
        <body>
        <table><tr><td>Hi, when will order A-1042 arrive?</td></tr></table>
        <div class="hidden">Marketing banner</div>
        </body>
        </html>
        """
        text = EmailResponse.strip_html_tags(html)
        self.assertIn("order A-1042", text)
        self.assertNotIn("Marketing banner", text)

    def test_extract_text_from_pdf(self):
        pdf_content = b"%PDF-1.4\nmock pdf content"
        result = EmailResponse.extract_text_from_pdf(pdf_content)
        self.assertIsInstance(result, str)

    def test_extract_text_from_docx(self):
        docx_content = b"PK\x03\x04"
        result = EmailResponse.extract_text_from_docx(docx_content)
        self.assertIsInstance(result, str)

    def test_extract_text_from_csv(self):
        csv_content = b"order,status\nA-1042,shipped\nB-7711,in transit"
        result = EmailResponse.extract_text_from_csv(csv_content)
        self.assertIn("A-1042", result)

    def test_parse_embedded_images_no_images(self):
        html = "<html><body><p>No images here</p></body></html>"
        result = EmailResponse.parse_embedded_images(html)
        self.assertEqual(result, "")

    def test_get_retry_config_defaults(self):
        import os
        os.environ.pop("MAX_RETRIES", None)
        os.environ.pop("RETRY_DELAY_SECONDS", None)
        config = EmailResponse.get_retry_config()
        self.assertEqual(config["max_retries"], 2)
        self.assertEqual(config["retry_delay"], 60)

    def test_get_retry_config_custom(self):
        import os
        os.environ["MAX_RETRIES"] = "5"
        os.environ["RETRY_DELAY_SECONDS"] = "30"
        config = EmailResponse.get_retry_config()
        self.assertEqual(config["max_retries"], 5)
        self.assertEqual(config["retry_delay"], 30)
        os.environ.pop("MAX_RETRIES", None)
        os.environ.pop("RETRY_DELAY_SECONDS", None)

    def test_rate_limiter_basic(self):
        limiter = EmailResponse.RateLimiter(max_tokens=5, refill_rate=1.0)
        self.assertEqual(limiter.tokens, 5)
        with limiter.acquire():
            self.assertEqual(limiter.tokens, 4)

    @patch('EmailResponse.msal.ConfidentialClientApplication')
    def test_get_ms_graph_token_client_credentials(self, mock_cca_cls):
        mock_cca = MagicMock()
        mock_cca.acquire_token_for_client.return_value = {"access_token": "mock-token"}
        mock_cca_cls.return_value = mock_cca
        
        token = EmailResponse.get_ms_graph_token(
            tenant_id="mock-tenant",
            client_id="mock-client",
            client_secret="mock-secret"
        )
        self.assertEqual(token, "mock-token")
        mock_cca_cls.assert_called_once_with(
            "mock-client",
            authority="https://login.microsoftonline.com/mock-tenant",
            client_credential="mock-secret"
        )
        mock_cca.acquire_token_for_client.assert_called_once_with(
            scopes=["https://graph.microsoft.com/.default"]
        )

    @patch('EmailResponse.rate_limited_request')
    @patch('EmailResponse.answer_email')
    @patch('EmailResponse.get_ms_graph_token')
    def test_process_msgraph_inbox_uses_rate_limiter(self, mock_get_token, mock_answer, mock_rate_limited):
        mock_get_token.return_value = "mock-token"
        
        mock_msg_resp = MagicMock()
        mock_msg_resp.status_code = 200
        mock_msg_resp.json.return_value = {
            "value": [
                {
                    "id": "msg-1",
                    "subject": "Order ETA Request",
                    "from": {"emailAddress": {"address": "client@example.com"}},
                    "receivedDateTime": "2026-06-14T10:00:00Z",
                    "body": {"content": "Hi, when will order A-1042 arrive?", "contentType": "text"},
                    "hasAttachments": True,
                    "categories": []
                }
            ]
        }
        
        mock_att_resp = MagicMock()
        mock_att_resp.status_code = 200
        mock_att_resp.json.return_value = {
            "value": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": "invoice.pdf",
                    "contentBytes": "bW9jay1wZGYtY29udGVudA=="
                }
            ]
        }
        
        mock_reply_resp = MagicMock()
        mock_reply_resp.status_code = 202
        mock_patch_resp = MagicMock()
        mock_patch_resp.status_code = 200
        
        mock_rate_limited.side_effect = [mock_msg_resp, mock_att_resp, mock_reply_resp, mock_patch_resp]
        mock_answer.return_value = "It will arrive in 2 business days."
        
        EmailResponse.process_msgraph_inbox(
            tenant_id="tenant-123",
            client_id="client-123",
            client_secret="secret-123",
            user_email="inbox@example.com",
            auto_reply=True,
            mark_read=False,
            create_draft=True
        )
        
        self.assertEqual(mock_rate_limited.call_count, 4)
        self.assertEqual(mock_rate_limited.call_args_list[0].args[0], "GET")
        self.assertEqual(mock_rate_limited.call_args_list[1].args[0], "GET")
        self.assertEqual(mock_rate_limited.call_args_list[2].args[0], "POST")
        self.assertEqual(mock_rate_limited.call_args_list[3].args[0], "PATCH")

    def test_extract_latest_email_html_with_div_separator(self):
        html_content = """
        <html>
        <body>
        <p>This is the latest message.</p>
        <div id="divRplyFwdMsg">
            <hr>
            <b>From:</b> old@example.com<br>
            <p>This is the old message.</p>
        </div>
        </body>
        </html>
        """
        extracted = EmailResponse.extract_latest_email_html(html_content)
        self.assertIn("This is the latest message.", extracted)
        self.assertNotIn("This is the old message.", extracted)
        self.assertNotIn("old@example.com", extracted)

    def test_extract_latest_email_text_with_separator(self):
        text_content = (
            "This is the latest message.\n\n"
            "-----Original Message-----\n"
            "From: old@example.com\n"
            "Sent: Friday, June 12, 2026 10:00 AM\n"
            "This is the old message."
        )
        extracted = EmailResponse.extract_latest_email_text(text_content)
        self.assertEqual(extracted.strip(), "This is the latest message.")

    @patch('EmailResponse.call_llm_api')
    @patch('EmailResponse.OpenAI')
    def test_answer_email_isolates_latest_message_for_llm(self, mock_openai_cls, mock_call_llm):
        thread_content = (
            "Can you reset my password?\n\n"
            "-----Original Message-----\n"
            "From: client@example.com\n"
            "Subject: when will order A-1042 arrive?\n"
            "Hi, when will order A-1042 arrive?"
        )
        
        mock_msg = MagicMock()
        mock_msg.content = ""
        mock_msg.tool_calls = None
        mock_call_llm.return_value = mock_msg
        
        reply = EmailResponse.answer_email(thread_content)
        
        mock_call_llm.assert_called_once()
        called_args = mock_call_llm.call_args[0]
        llm_input_text = called_args[1]
        
        self.assertIn("Can you reset my password?", llm_input_text)
        self.assertNotIn("order A-1042", llm_input_text)


if __name__ == "__main__":
    unittest.main()

