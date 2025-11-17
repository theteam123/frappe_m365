# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Check what's stored in existing M365 events
"""

import frappe


@frappe.whitelist()
def check_stored_events(limit=10):
	"""
	Check what's actually stored in the Event doctype for M365 events
	"""
	events = frappe.get_all(
		"Event",
		filters={"m365_event_id": ["!=", ""]},
		fields=["name", "subject", "description", "starts_on", "m365_event_id", "m365_icaluid"],
		limit=limit,
		order_by="creation desc"
	)
	
	print(f"\n{'='*80}")
	print(f"Checking {len(events)} stored M365 events")
	print(f"{'='*80}\n")
	
	for i, event in enumerate(events, 1):
		print(f"\n--- Event {i} ---")
		print(f"Name: {event.name}")
		print(f"Subject: {event.subject}")
		print(f"Subject type: {type(event.subject)}")
		print(f"Subject repr: {repr(event.subject)}")
		print(f"Description length: {len(event.description or '')}")
		print(f"Starts on: {event.starts_on}")
		print(f"M365 Event ID: {event.m365_event_id[:50]}..." if event.m365_event_id else "None")
		print(f"M365 iCalUId: {event.m365_icaluid[:50]}..." if event.m365_icaluid else "None")
	
	return {
		"success": True,
		"events_checked": len(events)
	}

