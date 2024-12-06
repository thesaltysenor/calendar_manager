import csv

def prompt_for_event():
    """Prompt the user to input event details."""
    print("Enter details for the event:")
    summary = input("  Event summary (e.g., Twins vs Yankees): ")
    year = input("  Start date - Year (e.g., 2024): ")
    month = input("  Start date - Month (1-12): ")
    day = input("  Start date - Day (1-31): ")
    start_hour = input("  Start time - Hour (0-23): ")
    start_minute = input("  Start time - Minute (0-59): ")

    end_hour = input("  End time - Hour (0-23): ")
    end_minute = input("  End time - Minute (0-59): ")

    # Construct the ISO 8601 datetime strings
    start_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}T{start_hour.zfill(2)}:{start_minute.zfill(2)}:00"
    end_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}T{end_hour.zfill(2)}:{end_minute.zfill(2)}:00"

    return {
        "summary": summary,
        "start_date": start_date,
        "end_date": end_date
    }

def create_csv(csv_file):
    """Create a CSV file for bulk event imports."""
    print("Creating a new CSV file for bulk events...")

    # Define the column headers
    headers = ["summary", "start_date", "end_date"]

    # Open the CSV file for writing
    with open(csv_file, mode='w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=headers)

        # Write the header row
        writer.writeheader()

        # Allow the user to add multiple events
        while True:
            event = prompt_for_event()
            writer.writerow(event)
            print(f"Event '{event['summary']}' added to the CSV.")

            # Ask if the user wants to add another event
            more = input("Add another event? (y/n): ").strip().lower()
            if more != 'y':
                break

    print(f"CSV file '{csv_file}' created successfully!")

def main():
    """Main function to create a CSV for bulk imports."""
    print("Welcome to the Bulk Event CSV Creator!")
    csv_file = input("Enter the name of the CSV file to create (e.g., schedule.csv): ").strip()
    create_csv(csv_file)

if __name__ == "__main__":
    main()
