from datetime import datetime
import os
import csv
import openpyxl
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
    "Squish Appointments": "#7db480",
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
    """Add an event to a specific calendar without overriding its default color."""
    calendar_id = get_calendar_id(calendar_name)

    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    # Validate time range
    start = datetime.strptime(start_time, "%Y-%m-%dT%H:%M:%S")
    end = datetime.strptime(end_time, "%Y-%m-%dT%H:%M:%S")
    if start >= end:
        print("Error: Start time must be earlier than end time.")
        return

    # Define the event WITHOUT explicit colorId
    event = {
        "summary": summary,
        "start": {"dateTime": start_time, "timeZone": "America/Chicago"},
        "end": {"dateTime": end_time, "timeZone": "America/Chicago"},
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "email", "minutes": 30},
                {"method": "popup", "minutes": 10},
            ],
        },
    }

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
    return f"{year}-{month.zfill(2)}-{day.zfill(2)}T{hour.zfill(2)}:{minute.zfill(2)}:00"

def get_calendar_id(calendar_name):
    """Get the calendar ID for a given calendar name."""
    calendars_list = service.calendarList().list().execute()
    for calendar in calendars_list["items"]:
        if calendar["summary"] == calendar_name:
            return calendar["id"]
    return None

def add_event_with_multiple_dates(calendar_name):
    """Add a single event with multiple dates and times."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    summary = input("Enter event summary (e.g., Twins vs Yankees): ").strip()
    events = []

    while True:
        print("\nEnter details for an event occurrence:")
        start_time = prompt_for_datetime("Enter start date and time")
        end_time = prompt_for_datetime("Enter end date and time")

        if datetime.fromisoformat(start_time) >= datetime.fromisoformat(end_time):
            print("❌ Error: Start time must be earlier than end time.")
            continue

        events.append({"summary": summary, "start_time": start_time, "end_time": end_time})

        more = input("Add another date/time for this event? (y/n): ").strip().lower()
        if more != "y":
            break

    for event in events:
        try:
            create_event(calendar_name, event['summary'], event['start_time'], event['end_time'])
        except Exception as e:
            print(f"❌ Error creating event '{event['summary']}': {e}")

    print(f"✅ {len(events)} occurrence(s) of '{summary}' added successfully to calendar '{calendar_name}'!")


def add_multiple_unique_events(calendar_name):
    """Continuously add unique events to a calendar until the user stops."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    events = []

    while True:
        summary = input("\nEnter event summary (or type 'done' to finish): ").strip()
        if summary.lower() == "done":
            break

        start_time = prompt_for_datetime("Enter start date and time")
        end_time = prompt_for_datetime("Enter end date and time")

        if datetime.fromisoformat(start_time) >= datetime.fromisoformat(end_time):
            print("❌ Error: Start time must be earlier than end time.")
            continue

        events.append({"summary": summary, "start_time": start_time, "end_time": end_time})

    for event in events:
        try:
            create_event(calendar_name, event['summary'], event['start_time'], event['end_time'])
        except Exception as e:
            print(f"❌ Error creating event '{event['summary']}': {e}")

    print(f"✅ {len(events)} unique event(s) added successfully to calendar '{calendar_name}'!")


def import_from_csv(calendar_name, csv_file):
    """Import events from a CSV file."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    color_id = get_calendar_color_id(calendar_name)  # Fetch default calendar color

    try:
        with open(csv_file, mode="r", newline="", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            print("CSV Headers:", reader.fieldnames)  # Debug to check headers
            
            for row in reader:
                summary = row.get("Summary")
                start_date = row.get("Start Date")
                start_time = row.get("Start Time")
                end_date = row.get("End Date")
                end_time = row.get("End Time")
                
                if not summary or not start_date or not start_time or not end_date or not end_time:
                    print("Skipping invalid row (missing fields):", row)
                    continue

                try:
                    start_datetime = f"{start_date}T{start_time}"
                    end_datetime = f"{end_date}T{end_time}"
                    
                    start_dt = datetime.fromisoformat(start_datetime)
                    end_dt = datetime.fromisoformat(end_datetime)
                    
                    if start_dt >= end_dt:
                        print("Skipping row with invalid time range:", row)
                        continue
                    
                    # Pass the colorId dynamically
                    event = {
                        "summary": summary,
                        "start": {"dateTime": start_datetime, "timeZone": "America/Chicago"},
                        "end": {"dateTime": end_datetime, "timeZone": "America/Chicago"},
                        "colorId": color_id,  # Use calendar's default color
                        "reminders": {
                            "useDefault": False,
                            "overrides": [
                                {"method": "email", "minutes": 30},
                                {"method": "popup", "minutes": 10},
                            ],
                        },
                    }

                    created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
                    print(f"Event added: {summary} on {start_date}")
                except ValueError as e:
                    print(f"Skipping row with invalid datetime format: {row}. Error: {e}")
                except KeyError as e:
                    print(f"Missing key in row: {row}. Error: {e}")
    except FileNotFoundError:
        print("CSV file not found. Please provide a valid path.")
    except Exception as e:
        print(f"Error processing CSV: {e}")

def get_calendar_color_id(calendar_name):
    """Retrieve the default color ID of a calendar."""
    try:
        calendars_list = service.calendarList().list().execute()
        for calendar in calendars_list["items"]:
            if calendar["summary"] == calendar_name:
                # Retrieve the calendar's default color ID
                color_id = calendar.get("colorId", "1")  # Default to '1' if not found
                return color_id
    except Exception as e:
        print(f"Failed to retrieve calendar color for '{calendar_name}': {e}")
    return "1"  # Default color ID if retrieval fails

def sync_event_colors(calendar_name):
    """Sync event colors to match the calendar's assigned color."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    # Get the calendar's designated color
    color_hex = calendars.get(calendar_name, "#236192")
    color_id = color_map.get(color_hex, "1")

    try:
        # Fetch all events from the calendar
        events_result = service.events().list(calendarId=calendar_id, singleEvents=True).execute()
        events = events_result.get('items', [])

        updated_count = 0
        for event in events:
            event_id = event.get('id')
            event_summary = event.get('summary', 'No Summary')
            event_color_id = event.get('colorId', None)

            if event_color_id != color_id:
                # Update the event color explicitly
                updated_event = {
                    "colorId": color_id
                }
                service.events().patch(
                    calendarId=calendar_id,
                    eventId=event_id,
                    body=updated_event
                ).execute()
                print(f"Updated color for event: {event_summary}")
                updated_count += 1

        print(f"Finished syncing colors. {updated_count} events updated.")

    except Exception as e:
        print(f"Error syncing event colors: {e}")

def inspect_calendar_color(calendar_name):
    """Inspect the colorId of a calendar."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    try:
        calendar = service.calendarList().get(calendarId=calendar_id).execute()
        color_id = calendar.get('colorId', 'Default')
        print(f"Calendar '{calendar_name}' has colorId: {color_id}")
    except Exception as e:
        print(f"Failed to retrieve calendar color: {e}")

def main():
    """Main menu for the script."""
    print("\nGoogle Calendar Manager")
    print("1. List Calendars")
    print("2. Create Calendar")
    print("3. Add Event with Multiple Dates/Times")
    print("4. Add Multiple Unique Events")
    print("5. Import Bulk Events from CSV")
    print("6. Sync Event Colors with Calendar Color")
    print("7. Inspect Calendar Color")
    choice = input("Enter your choice: ")

    if choice == "1":
        list_calendars()
    elif choice == "2":
        calendar_name = input("Enter calendar name: ")
        create_calendar(calendar_name)
    elif choice == "3":
        calendar_name = input("Enter calendar name: ")
        add_event_with_multiple_dates(calendar_name)
    elif choice == "4":
        calendar_name = input("Enter calendar name: ")
        add_multiple_unique_events(calendar_name)
    elif choice == "5":
        calendar_name = input("Enter calendar name: ")
        csv_file = input("Enter the path to the CSV file: ")
        import_from_csv(calendar_name, csv_file)
    elif choice == "6":
        calendar_name = input("Enter calendar name: ")
        sync_event_colors(calendar_name)
    elif choice == "7":
        calendar_name = input("Enter calendar name: ")
        inspect_calendar_color(calendar_name)
    else:
        print("❌ Invalid choice!")

if __name__ == "__main__":
    main()
