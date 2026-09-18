import http.client
import json
import os
import threading
import unittest

from export_and_launch import FormHandler
from socketserver import ThreadingTCPServer


class LocalServerSecurityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = ThreadingTCPServer(('127.0.0.1', 0), FormHandler)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.port = cls.server.server_address[1]

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=2)

    def request(self, method, path, body=None, headers=None):
        connection = http.client.HTTPConnection('127.0.0.1', self.port, timeout=3)
        connection.request(method, path, body=body, headers=headers or {})
        response = connection.getresponse()
        payload = response.read()
        connection.close()
        return response.status, payload

    def test_public_form_is_served(self):
        status, _ = self.request('GET', '/index.html')
        self.assertEqual(status, 200)

    def test_workspace_secrets_are_not_served(self):
        for path in ('/.env.local', '/ssh-key-2026-06-30.key', '/google_apps_script_backend.js'):
            status, _ = self.request('GET', path)
            self.assertEqual(status, 404, path)

    def test_cross_origin_write_is_rejected_before_handler(self):
        body = json.dumps({'test': True})
        status, _ = self.request(
            'POST', '/update_master', body,
            {'Content-Type': 'application/json', 'Content-Length': str(len(body)), 'Origin': 'https://example.invalid'},
        )
        self.assertEqual(status, 403)

    def test_local_login_requires_environment_credentials(self):
        old_email = os.environ.pop('LOCAL_ADMIN_EMAIL', None)
        old_password = os.environ.pop('LOCAL_ADMIN_PASSWORD', None)
        try:
            body = json.dumps({'email': 'admin@example.test', 'password': 'test'})
            status, payload = self.request(
                'POST', '/local_login', body,
                {'Content-Type': 'application/json', 'Content-Length': str(len(body)), 'Origin': f'http://127.0.0.1:{self.port}'},
            )
            self.assertEqual(status, 503)
            self.assertEqual(json.loads(payload)['status'], 'error')
        finally:
            if old_email is not None:
                os.environ['LOCAL_ADMIN_EMAIL'] = old_email
            if old_password is not None:
                os.environ['LOCAL_ADMIN_PASSWORD'] = old_password


if __name__ == '__main__':
    unittest.main()
