# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Debug script to inspect M365 event data
"""

import frappe
import json
from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import get_calendar_events_delta


@frappe.whitelist()
def inspect_event_data(email_account_name, limit=5, include_future=False):
	"""
	Fetch and inspect raw M365 event data to debug field mapping

	Args:
		email_account_name: Name of M365 Email Account
		limit: Number of events to inspect (default: 5)
		include_future: If True, fetch future events instead of using delta (default: False)
	"""
	email_account = frappe.get_doc("M365 Email Account", email_account_name)

	# Get access token
	access_token = get_access_token(email_account.service_principal)

	if include_future:
		# Fetch future events using regular list endpoint
		from m365email.m365email.graph_api import make_graph_request
		from datetime import datetime

		# Get events from today onwards
		today = datetime.utcnow().strftime("%Y-%m-%dT00:00:00Z")
		endpoint = f"/users/{email_account.email_address}/calendar/events"
		params = {
			"$filter": f"start/dateTime ge '{today}'",
			"$orderby": "start/dateTime",
			"$top": limit
		}
		response = make_graph_request(endpoint, access_token, params=params)
		events = response.get("value", [])
	else:
		# Fetch events using delta query
		response = get_calendar_events_delta(
			email_account.email_address,
			access_token,
			delta_token=None
		)
		events = response.get("value", [])[:limit]
	
	print(f"\n{'='*80}")
	print(f"Inspecting {len(events)} events from {email_account.email_address}")
	print(f"{'='*80}\n")
	
	for i, event in enumerate(events, 1):
		print(f"\n--- Event {i} ---")
		print(f"ID: {event.get('id', 'N/A')}")
		print(f"Subject: {event.get('subject', 'N/A')}")
		print(f"Subject type: {type(event.get('subject'))}")
		print(f"Subject repr: {repr(event.get('subject'))}")
		print(f"iCalUId: {event.get('iCalUId', 'N/A')}")
		print(f"Start: {event.get('start', {})}")
		print(f"End: {event.get('end', {})}")
		print(f"Location: {event.get('location', {})}")
		print(f"IsAllDay: {event.get('isAllDay', 'N/A')}")
		print(f"\nAll fields in event:")
		for key in sorted(event.keys()):
			if key not in ['id', 'subject', 'iCalUId', 'start', 'end', 'location', 'isAllDay']:
				print(f"  {key}: {type(event[key]).__name__}")
		print(f"\nFull event JSON:")
		print(json.dumps(event, indent=2, default=str))
		print(f"\n{'-'*80}")
	
	return {
		"success": True,
		"events_inspected": len(events),
		"message": "Check console output for details"
	}

