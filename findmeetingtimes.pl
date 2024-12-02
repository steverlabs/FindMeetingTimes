import msal
import requests
import argparse
from datetime import datetime, timezone, timedelta

# Replace with your credentials
CLIENT_ID = "your-client-id"
AUTHORITY = "https://login.microsoftonline.com/common"
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"

# Authenticate using Device Code Flow
def authenticate_user():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    device_flow = app.initiate_device_flow(scopes=["Calendars.Read"])
    print(device_flow["message"])  # Display authentication instructions
    result = app.acquire_token_by_device_flow(device_flow)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Authentication failed!")

# Calculate start and end times based on the week parameter
def calculate_time_range(week_param):
    now = datetime.now(timezone.utc)
    if week_param == "current":
        # Start of the current week
        start = now - timedelta(days=now.weekday())
        end = start + timedelta(days=6, hours=23, minutes=59, seconds=59)
    elif week_param == "next":
        # Start of the next week
        start = now - timedelta(days=now.weekday()) + timedelta(weeks=1)
        end = start + timedelta(days=6, hours=23, minutes=59, seconds=59)
    else:
        # Default: Next available times, at least 24 hours from now
        start = now + timedelta(days=1)
        end = start + timedelta(days=7)
    return start.isoformat(), end.isoformat()

# Find available meeting times
def find_meeting_times(access_token, duration_minutes, max_slots, week_param):
    headers = {"Authorization": f"Bearer {access_token}"}
    start_time, end_time = calculate_time_range(week_param)
    body = {
        "attendees": [],
        "timeConstraint": {
            "timeslots": [
                {
                    "start": {
                        "dateTime": start_time,
                        "timeZone": "UTC"
                    },
                    "end": {
                        "dateTime": end_time,
                        "timeZone": "UTC"
                    }
                }
            ]
        },
        "meetingDuration": f"PT{duration_minutes}M",
        "maxCandidates": max_slots
    }

    response = requests.post(f"{GRAPH_ENDPOINT}/me/findMeetingTimes", headers=headers, json=body)
    if response.status_code == 200:
        slots = response.json().get("meetingTimeSuggestions", [])
        results = []
        for slot in slots:
            start = slot["meetingTimeSlot"]["start"]["dateTime"]
            end = slot["meetingTimeSlot"]["end"]["dateTime"]
            start_dt = datetime.fromisoformat(start).astimezone(timezone.utc)
            end_dt = datetime.fromisoformat(end).astimezone(timezone.utc)
            results.append(f"- {start_dt.strftime('%A, %B %d, %Y, %I:%M %p')} - {end_dt.strftime('%I:%M %p')}")
        return results
    else:
        print(f"Error: {response.status_code}, {response.text}")
        return []

# Main logic
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Find available meeting times in your Outlook calendar.")
    parser.add_argument(
        "-d", "--duration", type=int, default=60,
        help="Duration of the meeting in minutes (default is 60 minutes)."
    )
    parser.add_argument(
        "-n", "--number", type=int, default=3,
        help="Number of available time slots to retrieve (default is 3)."
    )
    parser.add_argument(
        "-w", "--week", type=str, choices=["current", "next"], default=None,
        help="Week to find available times in: 'current' for this week, 'next' for next week. Defaults to next available times starting 24 hours from now."
    )
    args = parser.parse_args()

    try:
        token = authenticate_user()
        slots = find_meeting_times(token, args.duration, args.number, args.week)
        if slots:
            print("Available Meeting Times:")
            print("\n".join(slots))
        else:
            print("No available times found.")
    except Exception as e:
        print(f"An error occurred: {e}")
