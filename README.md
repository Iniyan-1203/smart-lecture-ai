# Smart Lecture AI

Smart Lecture AI is a bilingual learning assistant prototype for school and college students. It is designed around an ESP32 lecture-capture device, while the current software prototype works independently through a Flask web application.

## What the prototype does

- Supports student account registration, login, logout, protected learning pages, and profile pictures from Settings
- Supports unique usernames, Student/Teacher registration choices, and optional contact-based classmate discovery
- Includes a student community where classmates can send friend requests and chat privately after connecting
- Supports student-created study groups with join/leave membership and group chat
- Includes in-app notifications for friend requests, focus-timer completion, and unfinished daily learning tasks
- Tracks active study time with consent-based browser heartbeats and supports background Web Push reminders when VAPID keys and a scheduler are configured
- Includes a Personal AI Coach that turns learning activity and quiz results into one practical next step and a private check-in conversation
- Provides saved display preferences for Dark academic, Light study, Ocean blue, Forest green, Sunset coral, and Lavender calm themes, plus a mirrored left-handed dashboard layout
- Stores and displays lecture notes
- Accepts lecture audio uploads for transcription and AI note generation
- Generates English, Tamil, or bilingual study notes from a transcript or uploaded PDF, DOCX, PPTX, TXT, or Markdown file
- Provides MCQ, fill-in-the-blank, and short-answer practice
- Includes an AI assistant that can answer questions using saved lectures and optional web sources
- Provides a responsive dark academic dashboard with Study Room, learning metrics, focus guidance, and community leaderboard concepts
- Includes PWA manifest and service-worker support for an app-like experience
- Keeps new lecture notes and quiz history attached to the signed-in student, with private progress analytics
- Includes an approved teacher mode with private classrooms, six-character join codes, enrolled-student counts, and class-scoped announcements
- Adds secure session-cookie defaults and documents production HTTPS, OCR, Android, and deployment requirements
- Throttles repeated login and registration attempts per client address to reduce credential abuse
- Protects browser form submissions with session-bound CSRF tokens; JSON API endpoints remain protected by same-origin browser rules and authenticated route checks
- Adds browser security headers for clickjacking, MIME sniffing, referrer leakage, unsafe browser permissions, and HTTPS transport

## Run the software

1. Install Python 3.11+.
2. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

   Browser audio uploads do not require microphone packages on the server.
   Document note generation uses `pypdf`, `python-docx`, and `python-pptx` to extract text before sending it to the AI.

3. Create a `.env` file:

   ```text
   GROQ_API_KEY=your_key_here
   FLASK_SECRET_KEY=replace_with_a_long_random_secret
   ```

4. Start the Flask server:

   ```powershell
   python -m flask --app app run --debug
   ```

5. Open `http://127.0.0.1:5000/`.

Run the regression suite with:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

## Publish on Render

The repository includes `render.yaml` for a no-card demo deployment. Create a
new Render Blueprint from this repository and add `GROQ_API_KEY`,
`TEACHER_EMAILS`, and optional VAPID variables in the Render environment.
The blueprint uses Gunicorn and HTTPS-only cookies. The free demo stores its
SQLite database and uploaded files in `/tmp`, so data can be cleared when the
service restarts or redeploys. For real users, add persistent storage or move
to PostgreSQL and object storage before launch.

After deployment, open the Render HTTPS URL on a phone and choose **Install
app** or **Add to Home Screen**. The existing manifest and service worker make
the website installable as the mobile app experience. An APK/AAB still
requires a later Trusted Web Activity or Capacitor packaging step.

For background browser notifications in a deployed HTTPS environment, add VAPID
settings to `.env` (`VAPID_PUBLIC_KEY`, `VAPID_PRIVATE_KEY`, and
`VAPID_CLAIMS_EMAIL`), then run `python send_reminders.py` from a scheduler.
Students must grant browser notification permission. Without VAPID keys, the
in-app notification center and activity tracking still work, but push messages
cannot be delivered while the app is closed.

The first visit opens the student sign-in page; choose **Create an account** for first-time registration. Use a password of at least 8 characters for the local demo. Set a strong random `FLASK_SECRET_KEY` before any deployment.

## Demo flow

1. Open the dashboard.
2. Add a lecture manually or generate notes from a transcript.
3. Select English, Tamil, or English + Tamil notes.
4. Open **My Lecture Notes**.
5. Try MCQs, fill-ups, or short answers.
6. Ask the AI assistant to explain a saved concept.
7. Demonstrate the Study Room, progress snapshot, focus coach, and bilingual workflow.

## Classroom and production notes

Set `TEACHER_EMAILS=teacher@example.com` in `.env` before registration to provision approved teacher accounts. Teachers can create private classrooms from `/teacher`, share each six-character join code, and publish announcements to a selected class. Students join from the Exam Planner and see only announcements for classrooms they joined, plus legacy all-student announcements. Student lectures and quiz history are scoped to the signed-in account, while the original demo lecture remains available as read-only starter content.

Registration defaults to Student. Selecting Teacher only creates a teacher account when the email is also listed in `TEACHER_EMAILS`; otherwise it safely remains a student account. Contact discovery is opt-in: supported mobile browsers may open the device contact picker, and only email addresses explicitly selected by the student are checked against classmates who enabled discoverability.

The current document parser extracts text-based PDF, DOCX, PPTX, TXT, and Markdown files. Scanned PDFs need an OCR worker (for example Tesseract) before they can be converted into notes. The PWA can be packaged for Android with Trusted Web Activity or Capacitor, but a signed APK/AAB, Play Console setup, and Android UsageStats consent are still deployment steps rather than local Flask features.

## Hardware direction

The planned device uses an ESP32, I2S microphone, OLED display, LEDs, and SD-card backup. The hardware prototype is currently being tested separately from the software. The software can be demonstrated without hardware using the audio upload flow.

## Project scope

This is an innovation and design-thinking prototype. Institution-level accounts, live ESP32 upload, mobile screen-time APIs, notifications, and real student groups are future integration stages that require additional privacy, permissions, and deployment work.
