# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Debug helpers for M365 Email Integration
"""

import frappe


def check_email_queue_status():
	"""
	Check the status of emails in the queue
	Useful for debugging M365 email sending
	"""
	# Check if custom fields exist
	has_m365_fields = frappe.db.has_column("Email Queue", "m365_send")

	# Build fields list based on what exists
	fields = ["name", "status", "creation", "error"]

	if has_m365_fields:
		fields.extend(["m365_send", "m365_account"])

	# Get recent Email Queue entries
	emails = frappe.get_all(
		"Email Queue",
		fields=fields,
		order_by="creation desc",
		limit=10
	)

	print("\n" + "="*80)
	print("EMAIL QUEUE STATUS (Last 10)")
	print("="*80)

	if not has_m365_fields:
		print("\n⚠️  WARNING: M365 custom fields not found in Email Queue!")
		print("Run: bench --site your-site.local migrate")
		print()

	for email in emails:
		print(f"\nName: {email.name}")
		print(f"Status: {email.status}")
		if has_m365_fields:
			print(f"M365 Send: {email.get('m365_send', 0)}")
			print(f"M365 Account: {email.get('m365_account', 'N/A')}")
		print(f"Created: {email.creation}")
		if email.get("error"):
			print(f"Error: {email.error}")
		print("-" * 80)

	return emails


def check_m365_sending_config():
	"""
	Check M365 sending configuration
	"""
	from m365email.m365email.send import can_send_via_m365

	print("\n" + "="*80)
	print("M365 SENDING CONFIGURATION")
	print("="*80)

	can_send = can_send_via_m365()
	print(f"\nM365 Sending Available: {can_send}")

	if can_send:
		# Get default outgoing M365 Email Account
		account = frappe.get_value(
			"Email Account",
			{"service": "M365", "default_outgoing": 1, "enable_outgoing": 1},
			["name", "email_account_name", "email_id", "m365_service_principal", "enable_incoming", "enable_outgoing"],
			as_dict=True
		)
		if account:
			print(f"Sending Account: {account.email_account_name}")
			print(f"Email Address: {account.email_id}")
			print(f"Service Principal: {account.m365_service_principal}")
			print(f"Enable Incoming: {account.enable_incoming}")
			print(f"Enable Outgoing: {account.enable_outgoing}")
		else:
			print("No default outgoing M365 Email Account found")
	else:
		print("No M365 Email Account marked for sending")

	print("="*80 + "\n")


def manually_process_queue():
	"""
	Manually process the M365 email queue
	Useful for testing without waiting for scheduled task
	"""
	from m365email.m365email.send import process_email_queue_m365
	
	print("\n" + "="*80)
	print("MANUALLY PROCESSING M365 EMAIL QUEUE")
	print("="*80 + "\n")
	
	result = process_email_queue_m365()
	
	print(f"\nResults:")
	print(f"  Sent: {result['sent']}")
	print(f"  Failed: {result['failed']}")
	print("="*80 + "\n")
	
	return result


def check_recent_errors():
	"""
	Check recent M365 email errors
	"""
	errors = frappe.get_all(
		"Error Log",
		filters={"error": ["like", "%M365 Email%"]},
		fields=["name", "creation", "error"],
		order_by="creation desc",
		limit=5
	)
	
	print("\n" + "="*80)
	print("RECENT M365 EMAIL ERRORS")
	print("="*80)
	
	if not errors:
		print("\nNo errors found! ✅")
	else:
		for error in errors:
			print(f"\nError: {error.name}")
			print(f"Created: {error.creation}")
			print(f"Message: {error.error[:200]}...")
			print("-" * 80)
	
	print("="*80 + "\n")
	
	return errors

