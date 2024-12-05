import os
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Topics and Hex Colors
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
    "PWHL": "#2e1a47",
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
    "#2e1a47": "3",
}

def authenticate_google_calendar():
    """Authenticate and return a Google Calendar API service instance."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/calendar'])
    else:
        # If no valid credentials, initiate the OAuth flow
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', ['https://www.googleapis.com/auth/calendar'])
        creds = flow.run_local_server(port=0)
        # Save the credentials for future use
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('calendar', 'v3', credentials=creds)

service = authenticate_google_calendar()

def get_calendar_id(calendar_name):
    """Retrieve the ID of a calendar by its name."""
    calendars_list = service.calendarList().list().execute()
    for calendar in calendars_list['items']:
        if calendar['summary'] == calendar_name:
            return calendar['id']
    return None

def create_calendar(name):
    """Create a new calendar with the given name."""
    calendar = {
        'summary': name,
        'timeZone': 'America/Chicago'
    }
    created_calendar = service.calendars().insert(body=calendar).execute()
    print(f"Calendar created: {name} - ID: {created_calendar['id']}")
    return created_calendar['id']

def create_event(calendar_name, summary, start_time, end_time):
    """Add an event to a specified calendar."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    # Get the color associated with the calendar
    color_hex = calendars.get(calendar_name)
    color_id = color_map.get(color_hex, "1")  # Default to Light Blue if not found

    # Create the event
    event = {
        'summary': summary,
        'start': {'dateTime': start_time, 'timeZone': 'America/Chicago'},
        'end': {'dateTime': end_time, 'timeZone': 'America/Chicago'},
        'colorId': color_id,
    }

    # Insert the event into the calendar
    event = service.events().insert(calendarId=calendar_id, body=event).execute()
    print(f"Event created: {event.get('htmlLink')}")

def main():
    """Main menu for the script."""
    print("1. Create Calendars")
    print("2. Add Event")
    choice = input("Enter your choice: ")

    if choice == "1":
        # Create calendars for all topics
        for name in calendars.keys():
            create_calendar(name)
        print("Calendars created successfully!")
    elif choice == "2":
        # Add an event to a specific calendar
        print(f"Available calendars: {', '.join(calendars.keys())}")
        calendar_name = input("Enter the calendar name: ")
        if calendar_name not in calendars:
            print("Calendar not found!")
            return

        summary = input("Enter event summary: ")
        start_time = input("Enter start time (YYYY-MM-DDTHH:MM:SS): ")
        end_time = input("Enter end time (YYYY-MM-DDTHH:MM:SS): ")

        create_event(calendar_name, summary, start_time, end_time)
    else:
        print("Invalid choice!")

if __name__ == "__main__":
    main()
