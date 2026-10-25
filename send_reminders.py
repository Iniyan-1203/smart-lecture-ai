"""Send background study reminders for inactive students.

Run this command from a scheduler every hour or once per day:
    python send_reminders.py
"""

from app import send_due_reminders


if __name__ == "__main__":
    print(f"Created reminders: {send_due_reminders()}")
