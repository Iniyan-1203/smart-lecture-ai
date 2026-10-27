import os
import sqlite3
import tempfile
import unittest
from datetime import datetime
from uuid import uuid4

from werkzeug.security import generate_password_hash

import app


class SmartLectureAppTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original_database = app.DATABASE
        cls.database_file = tempfile.NamedTemporaryFile(
            suffix=".db", delete=False
        )
        cls.database_file.close()
        app.DATABASE = cls.database_file.name
        app.initialize_database()
        app.app.config.update(TESTING=True, WTF_CSRF_ENABLED=False)

    @classmethod
    def tearDownClass(cls):
        app.DATABASE = cls.original_database
        try:
            os.remove(cls.database_file.name)
        except OSError:
            pass

    def setUp(self):
        app._auth_attempts.clear()
        self.client = app.app.test_client()

    def csrf_token(self, client=None):
        client = client or self.client
        client.get("/login")
        with client.session_transaction() as session:
            if "_csrf_token" not in session:
                session["_csrf_token"] = "test-csrf-token"
            return session["_csrf_token"]

    def create_user(self, role="student"):
        suffix = uuid4().hex[:10]
        connection = sqlite3.connect(app.DATABASE)
        cursor = connection.execute(
            """
            INSERT INTO users
                (name, username, email, password_hash, created_at, role)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                f"Test {suffix}",
                f"test_{suffix}",
                f"{suffix}@example.com",
                generate_password_hash("password123"),
                datetime.now().isoformat(timespec="seconds"),
                role,
            ),
        )
        user_id = cursor.lastrowid
        connection.execute(
            "INSERT INTO user_settings (user_id) VALUES (?)", (user_id,)
        )
        connection.commit()
        connection.close()
        return user_id

    def login_session(self, user_id, client=None):
        client = client or self.client
        with client.session_transaction() as session:
            session["user_id"] = user_id
            session["_csrf_token"] = "test-csrf-token"
        return client

    def test_form_requires_csrf_and_accepts_valid_token(self):
        response = self.client.post(
            "/login",
            data={"email": "nobody@example.com", "password": "bad"},
        )
        self.assertEqual(response.status_code, 400)

        response = self.client.post(
            "/login",
            data={
                "email": "nobody@example.com",
                "password": "bad",
                "csrf_token": self.csrf_token(),
            },
        )
        self.assertEqual(response.status_code, 200)

    def test_registration_and_username_login(self):
        token = self.csrf_token()
        suffix = uuid4().hex[:8]
        response = self.client.post(
            "/register",
            data={
                "name": "Registered Student",
                "username": f"student_{suffix}",
                "email": f"{suffix}@example.com",
                "password": "password123",
                "role": "student",
                "csrf_token": token,
            },
        )
        self.assertEqual(response.status_code, 302)

        self.client.post("/logout", data={"csrf_token": self.csrf_token()})
        response = self.client.post(
            "/login",
            data={
                "email": f"student_{suffix}",
                "password": "password123",
                "csrf_token": self.csrf_token(),
            },
        )
        self.assertEqual(response.status_code, 302)

    def test_student_sees_only_owned_and_shared_notes(self):
        first_user = self.create_user()
        second_user = self.create_user()
        app.save_lecture("Physics", "Private one", "first", first_user)
        app.save_lecture("Physics", "Private two", "second", second_user)
        connection = sqlite3.connect(app.DATABASE)
        connection.execute(
            """
            INSERT INTO lectures (subject, date, title, content)
            VALUES (?, ?, ?, ?)
            """,
            ("Physics", "25 October 2026", "Starter", "shared"),
        )
        connection.commit()
        connection.close()

        self.login_session(first_user)
        page = self.client.get("/notes").get_data(as_text=True)
        self.assertIn("Private one", page)
        self.assertIn("Starter", page)
        self.assertNotIn("Private two", page)

    def test_classroom_join_and_private_announcement(self):
        teacher_id = self.create_user("teacher")
        student_id = self.create_user("student")
        teacher = self.login_session(teacher_id, app.app.test_client())
        response = teacher.post(
            "/teacher",
            data={
                "action": "create_classroom",
                "classroom_name": "Physics",
                "csrf_token": "test-csrf-token",
            },
        )
        self.assertEqual(response.status_code, 302)

        connection = sqlite3.connect(app.DATABASE)
        classroom = connection.execute(
            "SELECT id, join_code FROM classrooms WHERE teacher_id = ?",
            (teacher_id,),
        ).fetchone()
        connection.close()
        self.assertIsNotNone(classroom)

        student = self.login_session(student_id, app.app.test_client())
        response = student.post(
            "/classrooms/join",
            data={
                "join_code": classroom[1],
                "csrf_token": "test-csrf-token",
            },
        )
        self.assertEqual(response.status_code, 302)

        response = teacher.post(
            "/teacher",
            data={
                "action": "announcement",
                "classroom_id": classroom[0],
                "title": "Private update",
                "body": "Review chapter two.",
                "csrf_token": "test-csrf-token",
            },
        )
        self.assertEqual(response.status_code, 302)
        self.assertIn(
            "Private update",
            student.get("/planner").get_data(as_text=True),
        )

    def test_json_api_requires_login_but_keeps_json_response(self):
        response = self.client.post("/api/personal-coach", json={"message": "Hi"})
        self.assertEqual(response.status_code, 401)
        self.assertEqual(response.json["error"], "Please log in to use this feature.")


if __name__ == "__main__":
    unittest.main()
