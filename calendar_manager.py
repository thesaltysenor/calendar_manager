from datetime import datetime
import os
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Define calendars with their associated colors
calendars = {
    "Gopher Hockey": "#ac503c",
    "My Appointments": "#007ba7",
    "Bills": "#ccae00",
    "Anastasia": "#ccae00",
    "Other Jobs": "#713600",
    "Squish": "#b76e79",
    "Squish Appointment": "#7db480",
    "Twins": "#a9002d",
    "Wild": "#4d7663",
    "Wolves": "#236192",
}

color_map = {
    "#ac503c": "11",
    "#007ba7": "9",
    "#ccae00": "5",
    "#713600": "8",
    "#b76e79": "4",
    "#7db480": "2",
    "#a9002d": "11",
    "#4d7663": "10",
    "#236192": "1",
}

def authenticate_google_calendar():
    """Authenticate and return a Google Calendar API service instance."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file(
            "token.json", ["https://www.googleapis.com/auth/calendar"]
        )
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", ["https://www.googleapis.com/auth/calendar"]
        )
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("calendar", "v3", credentials=creds)

service = authenticate_google_calendar()

def list_calendars():
    """List all calendars in the user's calendar list."""
    calendars_list = service.calendarList().list().execute()
    for calendar in calendars_list["items"]:
        print(f"Calendar Name: {calendar['summary']} (ID: {calendar['id']})")

def create_calendar(name):
    """Create a new calendar with the given name."""
    calendar = {
        "summary": name,
        "timeZone": "America/Chicago",
    }
    created_calendar = service.calendars().insert(body=calendar).execute()
    print(f"Calendar created: {created_calendar['summary']} (ID: {created_calendar['id']})")
    return created_calendar["id"]

def create_event(calendar_name, summary, start_time, end_time):
    """Add an event to a specific calendar."""
    # Get calendar ID
    calendar_id = None
    calendars_list = service.calendarList().list().execute()
    for calendar in calendars_list["items"]:
        if calendar["summary"] == calendar_name:
            calendar_id = calendar["id"]
            break

    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

     # Validate time range
    start = datetime.strptime(start_time, "%Y-%m-%dT%H:%M:%S")
    end = datetime.strptime(end_time, "%Y-%m-%dT%H:%M:%S")
    if start >= end:
        print("Error: Start time must be earlier than end time.")
        return

    # Define the event
    color_hex = calendars.get(calendar_name, "#236192")
    color_id = color_map.get(color_hex, "1")

    event = {
        "summary": summary,
        "start": {"dateTime": start_time, "timeZone": "America/Chicago"},
        "end": {"dateTime": end_time, "timeZone": "America/Chicago"},
        "colorId": color_id,
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "email", "minutes": 30},
                {"method": "popup", "minutes": 10},
            ],
        },
    }

    # Insert the event
    try:
        created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
        print(f"Event created: {created_event['htmlLink']}")
    except Exception as e:
        print(f"Error creating event: {e}")

def prompt_for_datetime(prompt_text):
    """Prompt the user for a date and time in a more structured way."""
    print(f"{prompt_text}:")
    year = input("  Year (e.g., 2024): ")
    month = input("  Month (1-12): ")
    day = input("  Day (1-31): ")
    hour = input("  Hour (0-23): ")
    minute = input("  Minute (0-59): ")

    # Default seconds to "00"
    date_time = f"{year}-{month.zfill(2)}-{day.zfill(2)}T{hour.zfill(2)}:{minute.zfill(2)}:00"
    return date_time


def update_calendar_color(calendar_name, new_color_hex):
    """Update a calendar's color."""
    calendar_id = None
    calendars_list = service.calendarList().list().execute()
    for calendar in calendars_list["items"]:
        if calendar["summary"] == calendar_name:
            calendar_id = calendar["id"]
            break

    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    service.calendars().patch(
        calendarId=calendar_id,
        body={"backgroundColor": new_color_hex},
        colorRgbFormat=True,
    ).execute()
    print(f"Updated color for calendar '{calendar_name}' to {new_color_hex}.")

def main():
    """Main menu for the script."""
    print("Google Calendar Manager")
    print("1. List Calendars")
    print("2. Create Calendar")
    print("3. Add Event")
    print("4. Update Calendar Color")
    choice = input("Enter your choice: ")

    if choice == "1":
        list_calendars()
    elif choice == "2":
        calendar_name = input("Enter calendar name: ")
        create_calendar(calendar_name)
    elif choice == "3":
        calendar_name = input("Enter calendar name: ")
        summary = input("Enter event summary: ")
        start_time = prompt_for_datetime("Enter start date and time")
        end_time = prompt_for_datetime("Enter end date and time")
        create_event(calendar_name, summary, start_time, end_time)
    elif choice == "4":
        calendar_name = input("Enter calendar name: ")
        color = input("Enter new color hex (e.g., #ff5722): ")
        update_calendar_color(calendar_name, color)
    else:
        print("Invalid choice!")

if __name__ == "__main__":
    main()
