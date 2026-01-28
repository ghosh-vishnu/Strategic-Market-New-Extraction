import shutil
import tempfile
from pathlib import Path

from django.contrib.auth import get_user_model
from django.test import TestCase, override_settings
from django.core.files.uploadedfile import SimpleUploadedFile
from rest_framework import status
from rest_framework.test import APIClient

from converter.models import UploadedFile
from converter import views as converter_views


def make_docx(name: str = "sample.docx", content: bytes | None = None) -> SimpleUploadedFile:
    payload = content or b"dummy content"
    return SimpleUploadedFile(
        name,
        payload,
        content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


class ConverterSecurityTests(TestCase):
    def setUp(self) -> None:
        super().setUp()
        self._temp_media_root = tempfile.mkdtemp(prefix="converter-tests-")
        self.override = override_settings(MEDIA_ROOT=self._temp_media_root)
        self.override.enable()

        user_model = get_user_model()
        self.user_password = "testpass123"
        self.user = user_model.objects.create_user(
            email="tester@example.com",
            username="tester",
            password=self.user_password,
        )

        self.client = APIClient()
        logged_in = self.client.login(username=self.user.email, password=self.user_password)
        if not logged_in:
            raise AssertionError("Test client failed to log in user")

    def tearDown(self) -> None:
        converter_views.JOBS.clear()
        self.client.logout()
        self.override.disable()
        shutil.rmtree(self._temp_media_root, ignore_errors=True)
        super().tearDown()

    def test_upload_requires_authentication(self) -> None:
        unauthenticated_client = APIClient()
        response = unauthenticated_client.post(
            "/api/upload/",
            {"files": [make_docx()]},
            format="multipart",
        )
        self.assertEqual(response.status_code, status.HTTP_403_FORBIDDEN)

    def test_upload_rejects_unsupported_extensions(self) -> None:
        bad_file = SimpleUploadedFile("exploit.exe", b"payload", content_type="application/octet-stream")
        response = self.client.post(
            "/api/upload/",
            {"files": [bad_file]},
            format="multipart",
        )
        self.assertEqual(response.status_code, status.HTTP_400_BAD_REQUEST)
        self.assertIn("Unsupported file type", str(response.data.get("detail", "")))
        self.assertIn("exploit.exe", response.data.get("files", []))

    def test_upload_sanitizes_filename_and_limits_access(self) -> None:
        malicious_name = "../../Top Secret!.docx"
        response = self.client.post(
            "/api/upload/",
            {"files": [make_docx(malicious_name)]},
            format="multipart",
        )
        self.assertEqual(response.status_code, status.HTTP_200_OK)
        job_id = response.data["jobId"]

        uploaded_record = UploadedFile.objects.get(job_id=job_id)
        self.assertNotIn("..", uploaded_record.file_name)
        self.assertTrue(uploaded_record.file_name.endswith(".docx"))

        expected_path = Path(uploaded_record.file_path)
        self.assertEqual(expected_path.name, uploaded_record.file_name)
        self.assertTrue(expected_path.is_file())

        progress_response = self.client.get("/api/progress/", {"jobId": job_id})
        self.assertEqual(progress_response.status_code, status.HTTP_200_OK)

        other_user = get_user_model().objects.create_user(
            email="second@example.com",
            username="second",
            password="secondpass123",
        )
        other_client = APIClient()
        other_client.login(username=other_user.email, password="secondpass123")
        unauthorized_progress = other_client.get("/api/progress/", {"jobId": job_id})
        self.assertEqual(unauthorized_progress.status_code, status.HTTP_404_NOT_FOUND)

