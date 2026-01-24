# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Utility script to fix existing M365 synced events
- Disable send_reminder for all M365 synced events
- Optionally delete events older than a certain date
"""

import frappe
from frappe import _


def fix_send_reminder():
	"""
	Disable send_reminder for all M365 synced events
	"""
	# Get all events with m365_event_id
	events = frappe.get_all(
		"Event",
		filters={"m365_event_id": ["!=", ""]},
		fields=["name", "subject", "send_reminder"]
	)
	
	if not events:
		print("No M365 synced events found")
		return
	
	print(f"Found {len(events)} M365 synced events")
	
	updated = 0
	for event in events:
		if event.send_reminder == 1:
			frappe.db.set_value("Event", event.name, "send_reminder", 0, update_modified=False)
			updated += 1
	
	frappe.db.commit()
	print(f"Updated {updated} events to disable send_reminder")


def delete_old_events(before_date):
	"""
	Delete M365 synced events older than a specific date
	
	Args:
		before_date: Date string (YYYY-MM-DD) - events before this date will be deleted
	"""
	from frappe.utils import get_datetime
	
	cutoff_date = get_datetime(before_date)
	
	# Get all M365 events older than cutoff date
	old_events = frappe.get_all(
		"Event",
		filters={
			"m365_event_id": ["!=", ""],
			"starts_on": ["<", cutoff_date]
		},
		fields=["name", "subject", "starts_on"]
	)
	
	if not old_events:
		print(f"No M365 synced events found before {before_date}")
		return
	
	print(f"Found {len(old_events)} M365 synced events before {before_date}")
	
	for event in old_events:
		try:
			frappe.delete_doc("Event", event.name, ignore_permissions=True)
		except Exception as e:
			print(f"Failed to delete event {event.name}: {str(e)}")
	
	frappe.db.commit()
	print(f"Deleted {len(old_events)} old events")


def reset_calendar_delta_token(email_account_name):
	"""
	Reset the calendar delta token to force a full resync
	This will re-fetch all events from M365

	Args:
		email_account_name: Name of Email Account with service='M365'
	"""
	import json

	email_account = frappe.get_doc("Email Account", email_account_name)

	if email_account.service != "M365":
		print(f"Error: {email_account_name} is not an M365 email account")
		return

	delta_tokens = {}
	if email_account.m365_delta_tokens:
		try:
			delta_tokens = json.loads(email_account.m365_delta_tokens)
		except:
			delta_tokens = {}

	# Remove calendar delta token
	if "calendar_events" in delta_tokens:
		del delta_tokens["calendar_events"]
		email_account.db_set("m365_delta_tokens", json.dumps(delta_tokens), update_modified=False)
		frappe.db.commit()
		print(f"Reset calendar delta token for {email_account_name}")
	else:
		print(f"No calendar delta token found for {email_account_name}")


def delete_all_m365_events(email_account_name=None):
	"""
	Delete all M365 synced events (optionally for a specific account)

	Args:
		email_account_name: Name of Email Account with service='M365' (optional - if None, deletes all M365 events)
	"""
	filters = {"m365_event_id": ["!=", ""]}
	if email_account_name:
		filters["m365_email_account"] = email_account_name

	events = frappe.get_all(
		"Event",
		filters=filters,
		fields=["name"]
	)

	if not events:
		print("No M365 synced events found")
		return

	print(f"Found {len(events)} M365 synced events to delete")

	for event in events:
		try:
			frappe.delete_doc("Event", event.name, ignore_permissions=True)
		except Exception as e:
			print(f"Failed to delete event {event.name}: {str(e)}")

	frappe.db.commit()
	print(f"Deleted {len(events)} M365 events")


@frappe.whitelist()
def run_fix(fix_reminders=True, delete_before_date=None, reset_delta_token=None, delete_all_and_resync=None):
	"""
	Run fixes for M365 synced events

	Args:
		fix_reminders: Disable send_reminder for all M365 events (default: True)
		delete_before_date: Delete events before this date (YYYY-MM-DD) (optional)
		reset_delta_token: Email account name to reset delta token (optional)
		delete_all_and_resync: Email account name to delete all events and force fresh sync (optional)
	"""
	if delete_all_and_resync:
		print(f"Deleting all M365 events for {delete_all_and_resync}...")
		delete_all_m365_events(delete_all_and_resync)
		print(f"Resetting delta token for {delete_all_and_resync}...")
		reset_calendar_delta_token(delete_all_and_resync)
		print("Done! Run the sync again to re-import events with correct data.")
		return {"success": True, "message": "All events deleted and delta token reset. Run sync to re-import."}

	if fix_reminders:
		print("Fixing send_reminder for M365 events...")
		fix_send_reminder()

	if delete_before_date:
		print(f"Deleting events before {delete_before_date}...")
		delete_old_events(delete_before_date)

	if reset_delta_token:
		print(f"Resetting delta token for {reset_delta_token}...")
		reset_calendar_delta_token(reset_delta_token)

	return {"success": True, "message": "Fixes applied successfully"}

