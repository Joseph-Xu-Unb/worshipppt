import json
import os
import unittest
import urllib.error
from unittest.mock import patch

from scripts import renew_onedrive_token as renew


class FakeResponse:
    def __init__(self, body=b"", status=200, headers=None):
        self._body = body if isinstance(body, bytes) else json.dumps(body).encode()
        self.status = status
        self.headers = headers or {}

    def read(self, amount=None):
        return self._body if amount is None else self._body[:amount]

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


class FakeHTTPError(urllib.error.HTTPError):
    def __init__(self, code, body=b"", headers=None):
        super().__init__(
            "https://api.render.com/v1/test",
            code,
            "error",
            headers or {},
            None,
        )
        self._body = body if isinstance(body, bytes) else json.dumps(body).encode()

    def read(self, amount=None):
        return self._body if amount is None else self._body[:amount]


class FakeOpener:
    def __init__(self, responses):
        self.responses = list(responses)
        self.requests = []

    def open(self, request, timeout=None):
        self.requests.append(request)
        item = self.responses.pop(0)
        if isinstance(item, BaseException):
            raise item
        return item


class RenderEnvTests(unittest.TestCase):
    def test_get_env_reads_value(self):
        opener = FakeOpener(
            [FakeResponse({"key": "ONEDRIVE_REFRESH_TOKEN", "value": "secret"})]
        )
        value = renew.render_get_env(
            "api-key", "srv-123", "ONEDRIVE_REFRESH_TOKEN", opener=opener
        )
        self.assertEqual(value, "secret")
        self.assertEqual(opener.requests[0].method, "GET")
        self.assertIn("/services/srv-123/env-vars/", opener.requests[0].full_url)
        self.assertIn("Bearer api-key", opener.requests[0].get_header("Authorization"))

    def test_get_env_can_be_optional(self):
        self.assertIsNone(
            renew.render_get_env(
                "api-key",
                "srv-123",
                "ONEDRIVE_SCOPE",
                opener=FakeOpener([FakeHTTPError(404)]),
                allow_missing=True,
            )
        )
        with self.assertRaises(renew.RenewError):
            renew.render_get_env(
                "api-key",
                "srv-123",
                "ONEDRIVE_SCOPE",
                opener=FakeOpener([FakeHTTPError(404)]),
            )

    def test_set_env_uses_single_variable_endpoint(self):
        opener = FakeOpener(
            [FakeResponse({"key": "ONEDRIVE_REFRESH_TOKEN", "value": "new"})]
        )
        renew.render_set_env(
            "api-key",
            "srv-123",
            "ONEDRIVE_REFRESH_TOKEN",
            "new-token",
            opener=opener,
        )
        request = opener.requests[0]
        self.assertEqual(request.method, "PUT")
        self.assertIn("/env-vars/ONEDRIVE_REFRESH_TOKEN", request.full_url)
        self.assertEqual(json.loads(request.data), {"value": "new-token"})


class RefreshTests(unittest.TestCase):
    def test_refresh_returns_new_token(self):
        opener = FakeOpener(
            [FakeResponse({"refresh_token": "new-token", "expires_in": 3600})]
        )
        token = renew.refresh_refresh_token(
            "client",
            "consumers",
            "Files.ReadWrite offline_access",
            "old-token",
            opener=opener,
        )
        self.assertEqual(token, "new-token")
        body = opener.requests[0].data.decode()
        self.assertIn("grant_type=refresh_token", body)
        self.assertIn("refresh_token=old-token", body)
        self.assertIn("scope=Files.ReadWrite+offline_access", body)

    def test_invalid_grant_is_actionable(self):
        opener = FakeOpener([FakeResponse({"error": "invalid_grant"})])
        with self.assertRaises(renew.RenewError) as caught:
            renew.refresh_refresh_token(
                "client", "consumers", "Files.ReadWrite offline_access", "old", opener=opener
            )
        self.assertIn("过期", str(caught.exception))

    def test_missing_new_token_is_rejected(self):
        opener = FakeOpener([FakeResponse({})])
        with self.assertRaises(renew.RenewError):
            renew.refresh_refresh_token(
                "client", "consumers", "Files.ReadWrite offline_access", "old", opener=opener
            )


class MainTests(unittest.TestCase):
    def test_missing_render_credentials(self):
        with patch.dict(os.environ, {}, clear=True):
            self.assertEqual(renew.main(), 2)


if __name__ == "__main__":
    unittest.main()
