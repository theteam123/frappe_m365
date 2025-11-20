# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Calendar event synchronization module for M365 Email Integration
Core logic for syncing calendar events from M365 to Frappe Events
"""

import json
import frappe
from frappe import _
from datetime import datetime
from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import (
	get_calendar_events_delta,
	get_calendar_event_details,
	make_graph_request
)
from m365email.m365email.utils import parse_m365_datetime


def sync_calendar_events(email_account_name):
	"""
	Main sync function for calendar events
	Fetches events from M365 and creates/updates Frappe Events
	Handles both initial sync and incremental delta sync

	Args:
		email_account_name: Name of M365 Email Account

	Returns:
		dict: Sync results
	"""
	email_account = frappe.get_doc("M365 Email Account", email_account_name)

	if not email_account.sync_events:
		return {"success": False, "message": "Calendar event sync is not enabled"}

	try:
		# Get access token
		access_token = get_access_token(email_account.service_principal)

		# Get delta token for calendar events
		delta_tokens = {}
		if email_account.delta_tokens:
			try:
				delta_tokens = json.loads(email_account.delta_tokens)
			except:
				delta_tokens = {}

		calendar_delta_token = delta_tokens.get("calendar_events")

		# Get sync_from_date from email account settings
		sync_from_date = None
		if email_account.sync_from_date:
			from frappe.utils import get_datetime
			sync_from_date = get_datetime(email_account.sync_from_date)

		# For initial sync (no delta token), use calendarView endpoint to get all events
		# calendarView automatically expands recurring events into individual instances
		# Delta queries only return changes, not all events on first sync
		if not calendar_delta_token:
			print(f"M365 Event Sync: Performing initial full sync for {email_account.account_name}")

			# Use calendarView endpoint which expands recurring events
			from datetime import datetime, timedelta

			# Get events from sync_from_date or 1 year ago
			if sync_from_date:
				start_date = sync_from_date
			else:
				start_date = datetime.now() - timedelta(days=365)

			# Get events up to 1 year in the future
			end_date = datetime.now() + timedelta(days=365)

			start_date_str = start_date.strftime("%Y-%m-%dT00:00:00Z")
			end_date_str = end_date.strftime("%Y-%m-%dT23:59:59Z")

			# Use calendarView which expands recurring events
			endpoint = f"/users/{email_account.email_address}/calendar/calendarView"
			params = {
				"startDateTime": start_date_str,
				"endDateTime": end_date_str,
				"$orderby": "start/dateTime",
				"$top": 1000  # Adjust as needed
			}
			response = make_graph_request(endpoint, access_token, params=params)

			# After initial sync, get delta token for future incremental syncs
			delta_response = get_calendar_events_delta(
				email_account.email_address,
				access_token,
				delta_token=None
			)
			# Save the delta link for next sync
			delta_link = delta_response.get("@odata.deltaLink")
			if delta_link:
				delta_tokens["calendar_events"] = delta_link
				email_account.db_set("delta_tokens", json.dumps(delta_tokens), update_modified=False)
		else:
			# Use delta query for incremental sync
			print(f"M365 Event Sync: Performing incremental sync for {email_account.account_name}")
			response = get_calendar_events_delta(
				email_account.email_address,
				access_token,
				delta_token=calendar_delta_token
			)

		events = response.get("value", [])
		fetched = len(events)
		created = 0
		updated = 0
		deleted = 0
		failed = 0
		skipped = 0

		# Get sync_from_date filter
		sync_from_date = None
		if email_account.sync_from_date:
			from frappe.utils import get_datetime
			sync_from_date = get_datetime(email_account.sync_from_date)

		# Process each event
		for event in events:
			try:
				# Check if event is deleted (removed property exists)
				if event.get("@removed"):
					result = delete_event_from_frappe(event, email_account)
					if result == "deleted":
						deleted += 1
				else:
					# Delta queries often return minimal data, so fetch full event details
					# if critical fields are missing
					event_id = event.get("id")
					if not event.get("subject") and not event.get("iCalUId"):
						# Fetch full event details
						try:
							full_event = get_calendar_event_details(
								email_account.email_address,
								event_id,
								access_token
							)
							# Use full event data instead
							event = full_event
						except Exception as e:
							print(f"Failed to fetch full details for event {event_id[:20]}...: {str(e)}")
							# Continue with partial data

					# Filter by sync_from_date if set
					if sync_from_date:
						event_start = parse_m365_datetime(event.get("start", {}).get("dateTime"))
						if event_start and event_start < sync_from_date:
							skipped += 1
							continue

					result = create_or_update_event(event, email_account, access_token)
					if result == "created":
						created += 1
					elif result == "updated":
						updated += 1
			except Exception as e:
				failed += 1
				frappe.log_error(
					title="M365 Event Sync: Failed to sync event",
					message=f"Event ID: {event.get('id')}\nSubject: {event.get('subject', 'N/A')}\n\nError: {str(e)}"
				)
		
		# Save new delta token
		delta_link = response.get("@odata.deltaLink")
		if delta_link:
			delta_tokens["calendar_events"] = delta_link
			email_account.db_set("delta_tokens", json.dumps(delta_tokens), update_modified=False)
		
		frappe.db.commit()
		
		return {
			"success": True,
			"fetched": fetched,
			"created": created,
			"updated": updated,
			"deleted": deleted,
			"skipped": skipped,
			"failed": failed
		}
		
	except Exception as e:
		frappe.log_error(
			title=f"M365 Event Sync Failed: {email_account.name}",
			message=str(e)
		)
		return {
			"success": False,
			"message": str(e)
		}


def create_or_update_event(event_data, email_account, access_token):
	"""
	Convert M365 event to Frappe Event doctype
	Creates new event or updates existing one

	Args:
		event_data: Event dict from Graph API
		email_account: M365 Email Account doc
		access_token: Access token for additional API calls if needed

	Returns:
		str: "created", "updated", or "skipped"
	"""
	event_id = event_data.get("id")
	ical_uid = event_data.get("iCalUId")
	
	# Check if event already exists
	existing = frappe.db.get_value(
		"Event",
		{"m365_event_id": event_id},
		["name", "modified"]
	)
	
	# Parse event data
	subject = event_data.get("subject")

	# Debug logging (optional)
	# if not subject:
	# 	print(f"M365 Event {event_id[:20]}... has no subject field in API response")
	# elif subject.strip() == "":
	# 	print(f"M365 Event {event_id[:20]}... has empty subject")

	# Set default if subject is missing or empty
	if not subject or subject.strip() == "":
		subject = "(No Subject)"

	body_content = event_data.get("body", {}).get("content", "")

	# Get timezone information from M365
	# M365 provides times in UTC, but also includes the original timezone
	original_timezone = event_data.get("originalStartTimeZone")

	# Parse datetimes - M365 sends them in UTC
	# We'll convert to the USER's timezone (not the event's original timezone)
	start_datetime_utc = parse_m365_datetime(event_data.get("start", {}).get("dateTime"))
	end_datetime_utc = parse_m365_datetime(event_data.get("end", {}).get("dateTime"))

	# Get user's timezone from email account settings
	user_timezone = email_account.user_timezone or "Australia/Perth"

	# Convert from UTC to user's timezone
	if start_datetime_utc:
		from m365email.m365email.utils import get_timezone
		import pytz

		try:
			# Get timezone objects
			utc_tz = pytz.UTC
			user_tz = get_timezone(user_timezone)

			# Make UTC datetime timezone-aware
			start_aware = utc_tz.localize(start_datetime_utc)
			end_aware = utc_tz.localize(end_datetime_utc) if end_datetime_utc else None

			# Convert to user's timezone
			start_local = start_aware.astimezone(user_tz)
			end_local = end_aware.astimezone(user_tz) if end_aware else None

			# Convert back to naive datetime for Frappe
			start_datetime = start_local.replace(tzinfo=None)
			end_datetime = end_local.replace(tzinfo=None) if end_local else None
		except Exception as e:
			print(f"Timezone conversion failed for event {event_id[:20]}...: {str(e)}")
			# Fallback to UTC times
			start_datetime = start_datetime_utc
			end_datetime = end_datetime_utc
	else:
		# No datetime, skip
		start_datetime = start_datetime_utc
		end_datetime = end_datetime_utc

	location = event_data.get("location", {}).get("displayName", "")
	is_all_day = event_data.get("isAllDay", False)

	# Get attendees
	attendees = []
	for attendee in event_data.get("attendees", []):
		email = attendee.get("emailAddress", {}).get("address")
		if email:
			attendees.append(email)

	if existing:
		# Update existing event
		event_doc = frappe.get_doc("Event", existing[0])
		event_doc.subject = subject
		event_doc.description = body_content
		event_doc.starts_on = start_datetime
		event_doc.ends_on = end_datetime
		event_doc.event_location = location
		event_doc.all_day = 1 if is_all_day else 0
		event_doc.send_reminder = 0  # Always disable reminders for synced events
		event_doc.m365_timezone = original_timezone  # Store original timezone
		event_doc.save(ignore_permissions=True)
		frappe.db.commit()
		return "updated"
	else:
		# Create new event
		event_doc = frappe.get_doc({
			"doctype": "Event",
			"subject": subject,
			"description": body_content,
			"starts_on": start_datetime,
			"ends_on": end_datetime,
			"event_location": location,
			"all_day": 1 if is_all_day else 0,
			"send_reminder": 0,  # Always disable reminders for synced events
			"m365_event_id": event_id,
			"m365_email_account": email_account.name,
			"m365_icaluid": ical_uid,
			"m365_timezone": original_timezone,  # Store original timezone
			"owner": email_account.user,
			"event_participants": [{"reference_doctype": "User", "reference_docname": email_account.user}]
		})
		event_doc.insert(ignore_permissions=True)
		frappe.db.commit()
		return "created"


def delete_event_from_frappe(event_data, email_account):
	"""
	Delete event from Frappe when it's deleted in M365

	Args:
		event_data: Event dict from Graph API (with @removed property)
		email_account: M365 Email Account doc

	Returns:
		str: "deleted" or "skipped"
	"""
	event_id = event_data.get("id")

	# Check if event exists
	existing = frappe.db.get_value("Event", {"m365_event_id": event_id})

	if existing:
		frappe.delete_doc("Event", existing, ignore_permissions=True)
		frappe.db.commit()
		return "deleted"

	return "skipped"

