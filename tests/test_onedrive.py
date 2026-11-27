import json
import os
import threading
import time
import unittest
import urllib.error
from datetime import date, timedelta
from unittest.mock import patch

from src.backend import onedrive
from src.backend.onedrive import (
    ArchiveLookup,
    DriveItemRef,
    GraphClient,
    OneDriveArchive,
    OneDriveFile,
    OneDriveSettings,
    OneDriveUploadError,
    TreeListing,
    auto_last_week_pptx,
    encode_share_url,
    get_status,
    is_allowed_download_url,
    is_safe_upload_name,
    parse_dates_from_filename,
    previous_sunday,
)


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
            "https://graph.microsoft.com/v1.0/test",
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


class StubTokens:
    def __init__(self):
        self.invalidated = 0

    def get_access_token(self, force=False):
        return "fake-token"

    def invalidate(self):
        self.invalidated += 1


def settings(**overrides):
    values = {
        "client_id": "client",
        "refresh_token": "refresh-secret",
        "tenant": "consumers",
        "scope": "Files.ReadWrite offline_access",
        "archive_share_url": "https://onedrive.live.com/personal/abc/_layouts/15/download.aspx?UniqueId=def",
        "archive_path": "",
        "token_issued_at": date.today(),
        "auto_enabled": True,
        "upload_enabled": True,
        "timeout": 5.0,
    }
    values.update(overrides)
    return OneDriveSettings(**values)


class DateTests(unittest.TestCase):
    def test_previous_sunday(self):
        self.assertEqual(previous_sunday(date(2026, 11, 29)), date(2026, 11, 22))
        self.assertEqual(previous_sunday(date(2026, 11, 30)), date(2026, 11, 29))
        self.assertEqual(previous_sunday(date(2026, 11, 28)), date(2026, 11, 22))
        self.assertEqual(previous_sunday(date(2026, 1, 1)), date(2025, 12, 28))

    def test_filename_dates(self):
        expected = date(2026, 11, 22)
        for name in (
            "2026-11-22.pptx",
            "Worship_2026-11-22.pptx",
            "2026.11.22.pptx",
            "2026_11_22.pptx",
            "20261122.pptx",
            "2026年11月22日.pptx",
        ):
            with self.subTest(name=name):
                self.assertEqual(parse_dates_from_filename(name), [expected])

    def test_invalid_and_non_pptx_names(self):
        self.assertEqual(parse_dates_from_filename("2026-02-31.pptx"), [])
        self.assertEqual(parse_dates_from_filename("notes.pptx"), [])
        self.assertFalse(is_safe_upload_name("~$2026-11-22.pptx"))
        self.assertFalse(is_safe_upload_name("2026-11-22.pptx/evil"))
        self.assertFalse(is_safe_upload_name("../2026-11-22.pptx"))
        self.assertTrue(is_safe_upload_name("2026-11-22.pptx"))

    def test_download_hosts(self):
        self.assertTrue(
            is_allowed_download_url(
                "https://my.microsoftpersonalcontent.com/personal/abc/items/def"
            )
        )
        self.assertTrue(
            is_allowed_download_url(
                "https://public.dm.files.1drv.com/personal/abc/file.pptx"
            )
        )
        self.assertFalse(
            is_allowed_download_url("https://example.test/file.pptx")
        )
        self.assertFalse(
            is_allowed_download_url("http://my.microsoftpersonalcontent.com/file")
        )

    def test_share_encoding(self):
        encoded = encode_share_url("https://1drv.ms/f/abc")
        self.assertTrue(encoded.startswith("u!"))
        self.assertNotIn("=", encoded)
        self.assertNotIn("+", encoded)
        self.assertNotIn("/", encoded)


class SettingsTests(unittest.TestCase):
    def test_missing_credentials_disable_auto(self):
        self.assertIsNone(OneDriveSettings.from_env({}))

    def test_env_flags_and_token_status(self):
        values = {
            "ONEDRIVE_CLIENT_ID": "client",
            "ONEDRIVE_REFRESH_TOKEN": "secret",
            "ONEDRIVE_TOKEN_ISSUED_AT": (
                date.today() - timedelta(days=80)
            ).isoformat(),
            "ONEDRIVE_AUTO_ENABLED": "false",
            "ONEDRIVE_UPLOAD_ENABLED": "0",
        }
        parsed = OneDriveSettings.from_env(values)
        self.assertIsNotNone(parsed)
        self.assertFalse(parsed.auto_enabled)
        self.assertFalse(parsed.upload_enabled)
        self.assertEqual(parsed.token_status(), "renew_soon")
        self.assertNotIn("secret", repr(parsed))

    def test_read_only_scope_cannot_write(self):
        parsed = settings(scope="Files.Read offline_access")
        self.assertFalse(parsed.can_write)
        status = get_status()
        self.assertIn("token_status", status)

    def test_status_without_environment(self):
        with patch.dict(os.environ, {}, clear=True):
            status = get_status()
        self.assertFalse(status["configured"])
        self.assertNotIn("secret", json.dumps(status))


class GraphTests(unittest.TestCase):
    def test_resolve_share_uses_remote_item(self):
        payload = {
            "id": "wrapper",
            "name": "Archive",
            "remoteItem": {
                "id": "folder1",
                "name": "Archive",
                "folder": {},
                "parentReference": {"driveId": "drive1"},
            },
        }
        opener = FakeOpener([FakeResponse(payload)])
        graph = GraphClient(settings(), tokens=StubTokens(), opener=opener)
        root = graph.resolve_share(settings().archive_share_url)
        self.assertEqual(root.drive_id, "drive1")
        self.assertEqual(root.item_id, "folder1")
        self.assertTrue(root.is_folder)
        self.assertIn("/shares/", opener.requests[0].full_url)

    def test_pagination_follows_next_link(self):
        first = {
            "value": [
                {
                    "id": "folder2",
                    "name": "2026",
                    "folder": {},
                    "parentReference": {"driveId": "drive1"},
                }
            ],
            "@odata.nextLink": (
                "https://graph.microsoft.com/v1.0/drives/drive1/items/folder1/"
                "children?$skiptoken=abc"
            ),
        }
        second = {
            "value": [
                {
                    "id": "file1",
                    "name": "2026-11-22.pptx",
                    "file": {},
                    "size": 123,
                    "parentReference": {"driveId": "drive1"},
                }
            ]
        }
        opener = FakeOpener([FakeResponse(first), FakeResponse(second)])
        graph = GraphClient(settings(), tokens=StubTokens(), opener=opener)
        children = graph.list_children(
            DriveItemRef("drive1", "folder1", "Archive", True, "Archive")
        )
        self.assertEqual([child.name for child in children], ["2026", "2026-11-22.pptx"])
        self.assertEqual(len(opener.requests), 2)

    def test_401_refreshes_once(self):
        tokens = StubTokens()
        opener = FakeOpener(
            [
                FakeHTTPError(401, {"error": {"code": "InvalidAuthenticationToken"}}),
                FakeResponse({"value": []}),
            ]
        )
        graph = GraphClient(settings(), tokens=tokens, opener=opener)
        children = graph.list_children(
            DriveItemRef("drive1", "folder1", "Archive", True, "Archive")
        )
        self.assertEqual(children, [])
        self.assertEqual(tokens.invalidated, 1)

    def test_error_message_redacts_urls(self):
        opener = FakeOpener(
            [
                FakeHTTPError(
                    403,
                    {
                        "error": {
                            "code": "accessDenied",
                            "message": "see https://example.test/secret-token",
                        }
                    },
                )
            ]
        )
        graph = GraphClient(settings(), tokens=StubTokens(), opener=opener)
        with self.assertRaises(onedrive.GraphHTTPError) as caught:
            graph.list_children(
                DriveItemRef("drive1", "folder1", "Archive", True, "Archive")
            )
        self.assertNotIn("secret-token", str(caught.exception))
        self.assertIn("<url>", str(caught.exception))

    def test_simple_upload_and_share_link(self):
        uploaded = {
            "id": "item1",
            "name": "2026-11-29.pptx",
            "webUrl": "https://onedrive.live.com/item",
            "parentReference": {"driveId": "drive1"},
        }
        link = {"link": {"webUrl": "https://1drv.ms/u/share"}}
        opener = FakeOpener([FakeResponse(uploaded), FakeResponse(link)])
        graph = GraphClient(settings(), tokens=StubTokens(), opener=opener)
        folder = DriveItemRef("drive1", "folder1", "Archive", True, "Archive")
        result = graph.upload_file(folder, "2026-11-29.pptx", b"PK-test")
        self.assertEqual(result["id"], "item1")
        self.assertIn("conflictBehavior=replace", opener.requests[0].full_url)
        self.assertEqual(
            graph.create_share_link("drive1", "item1"),
            "https://1drv.ms/u/share",
        )

    def test_anonymous_download_url_appends_download(self):
        location = (
            "https://onedrive.live.com/:p:/g/personal/ABC/TOKEN"
            "?resid=ABC%21s123&redeem=YWJj"
        )
        opener = FakeOpener(
            [FakeHTTPError(301, headers={"Location": location})]
        )
        graph = GraphClient(settings(), tokens=StubTokens(), opener=FakeOpener([]))
        url = graph.make_anonymous_download_url(
            "https://1drv.ms/p/c/ABC/TOKEN", opener=opener
        )
        self.assertIn("download=1", url)
        self.assertIn("redeem=YWJj", url)
        self.assertTrue(url.startswith("https://onedrive.live.com/"))


class ArchiveTests(unittest.TestCase):
    def _archive_with_files(self, files):
        archive = OneDriveArchive(settings())
        archive._listing = TreeListing(files=files)
        archive._listing_at = time.time()
        return archive

    def test_lookup_single_exact_match(self):
        item = OneDriveFile(
            drive_id="drive1",
            item_id="item1",
            name="2026-11-22.pptx",
            modified="2026-11-23T00:00:00Z",
        )
        lookup = self._archive_with_files([item]).lookup(date(2026, 11, 29))
        self.assertIs(lookup.file, item)
        self.assertEqual(lookup.target, date(2026, 11, 22))

    def test_lookup_multiple_matches_returns_candidates(self):
        files = [
            OneDriveFile("d", "1", "2026-11-22.pptx"),
            OneDriveFile("d", "2", "Worship_2026-11-22.pptx"),
        ]
        lookup = self._archive_with_files(files).lookup(date(2026, 11, 29))
        self.assertIsNone(lookup.file)
        self.assertEqual(len(lookup.candidates), 2)
        self.assertIn("多个", lookup.error)

    def test_lookup_missing_returns_candidates(self):
        files = [OneDriveFile("d", "1", "2026-11-15.pptx")]
        lookup = self._archive_with_files(files).lookup(date(2026, 11, 29))
        self.assertIsNone(lookup.file)
        self.assertEqual(lookup.candidates[0]["name"], "2026-11-15.pptx")
        self.assertIn("没有找到", lookup.error)

    def test_lookup_explicit_filename(self):
        item = OneDriveFile("d", "1", "2026-11-22.pptx")
        lookup = self._archive_with_files([item]).lookup(
            date(2026, 11, 29), filename="2026-11-22.pptx"
        )
        self.assertIs(lookup.file, item)

    def test_upload_generated_deck_uses_date_filename_and_share_link(self):
        archive = OneDriveArchive(settings())
        archive._root = DriveItemRef("drive1", "folder1", "Archive", True, "Archive")
        archive._listing = TreeListing()
        archive._listing_at = time.time()

        class StubGraph:
            def upload_file(self, folder, filename, content):
                return {
                    "id": "item1",
                    "name": filename,
                    "webUrl": "https://onedrive.live.com/item",
                    "parentReference": {"driveId": "drive1"},
                }

            def create_share_link(self, drive_id, item_id):
                return "https://1drv.ms/u/share"

            def make_anonymous_download_url(self, view_url):
                return view_url + "?download=1"

        archive.graph = StubGraph()
        result = archive.upload_generated_deck(date(2026, 11, 29), b"PK-test")
        self.assertTrue(result["uploaded"])
        self.assertEqual(result["filename"], "2026-11-29.pptx")
        self.assertEqual(result["view_url"], "https://1drv.ms/u/share")
        self.assertEqual(
            result["share_url"], "https://1drv.ms/u/share?download=1"
        )

    def test_upload_falls_back_to_view_link(self):
        archive = OneDriveArchive(settings())
        archive._root = DriveItemRef("drive1", "folder1", "Archive", True, "Archive")
        archive._listing = TreeListing()
        archive._listing_at = time.time()

        class StubGraph:
            def upload_file(self, folder, filename, content):
                return {
                    "id": "item1",
                    "name": filename,
                    "webUrl": "https://onedrive.live.com/item",
                    "parentReference": {"driveId": "drive1"},
                }

            def create_share_link(self, drive_id, item_id):
                return "https://1drv.ms/u/share"

            def make_anonymous_download_url(self, view_url):
                raise onedrive.OneDriveError("download link failed")

        archive.graph = StubGraph()
        result = archive.upload_generated_deck(date(2026, 11, 29), b"PK-test")
        self.assertTrue(result["uploaded"])
        self.assertEqual(result["share_url"], "https://1drv.ms/u/share")
        self.assertTrue(any("匿名下载" in w for w in result["warnings"]))

    def test_list_files_does_not_deadlock_when_root_is_reentrant(self):
        archive = OneDriveArchive(settings())
        archive._root = DriveItemRef(
            "drive1", "folder1", "Archive", True, "Archive"
        )
        archive._listing = TreeListing()
        archive._listing_at = 0

        class StubGraph:
            def list_tree(self, root):
                return TreeListing()

        archive.graph = StubGraph()
        outcome = {}

        def run():
            try:
                outcome["listing"] = archive.list_files(force=True)
            except Exception as exc:  # pragma: no cover - defensive
                outcome["error"] = exc

        worker = threading.Thread(target=run, daemon=True)
        worker.start()
        worker.join(2)
        self.assertFalse(worker.is_alive(), "list_files deadlocked")
        self.assertNotIn("error", outcome)
        self.assertEqual(outcome["listing"].files, [])

    def test_upload_failure_is_soft(self):
        archive = OneDriveArchive(settings())
        archive._root = DriveItemRef("drive1", "folder1", "Archive", True, "Archive")
        archive._listing = TreeListing()
        archive._listing_at = time.time()

        class StubGraph:
            def upload_file(self, folder, filename, content):
                raise OneDriveUploadError("upload failed")

        archive.graph = StubGraph()
        result = archive.upload_generated_deck(date(2026, 11, 29), b"PK-test")
        self.assertFalse(result["uploaded"])
        self.assertIn("upload failed", result["error"])


class AutoLastWeekTests(unittest.TestCase):
    def test_auto_fetch_success(self):
        class StubArchive:
            class graph:
                @staticmethod
                def download_file(item):
                    return b"PK-test"

            def lookup(self, service_date, filename=""):
                return ArchiveLookup(
                    file=OneDriveFile("d", "1", "2026-11-22.pptx"),
                    target=date(2026, 11, 22),
                )

        with patch.object(
            onedrive.OneDriveSettings, "from_env", return_value=settings()
        ), patch.object(
            onedrive, "get_onedrive_archive", return_value=StubArchive()
        ):
            result = auto_last_week_pptx(date(2026, 11, 29))
        self.assertTrue(result.attempted)
        self.assertEqual(result.content, b"PK-test")
        self.assertEqual(result.filename, "2026-11-22.pptx")

    def test_auto_fetch_missing_returns_candidates(self):
        class StubArchive:
            def lookup(self, service_date, filename=""):
                return ArchiveLookup(
                    target=date(2026, 11, 22),
                    candidates=[{"name": "2026-11-15.pptx"}],
                    error="没有找到",
                )

        with patch.object(
            onedrive.OneDriveSettings, "from_env", return_value=settings()
        ), patch.object(
            onedrive, "get_onedrive_archive", return_value=StubArchive()
        ):
            result = auto_last_week_pptx(date(2026, 11, 29))
        self.assertTrue(result.attempted)
        self.assertIsNone(result.content)
        self.assertEqual(result.candidates[0]["name"], "2026-11-15.pptx")

    def test_auto_fetch_without_configuration(self):
        with patch.object(onedrive.OneDriveSettings, "from_env", return_value=None):
            result = auto_last_week_pptx(date(2026, 11, 29))
        self.assertFalse(result.attempted)
        self.assertIsNone(result.content)


if __name__ == "__main__":
    unittest.main()
