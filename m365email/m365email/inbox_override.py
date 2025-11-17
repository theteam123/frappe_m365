# Copyright (c) 2025, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

import frappe
from frappe.email import inbox


# Store original function
_original_get_email_accounts = inbox.get_email_accounts


def get_email_accounts_with_m365(user=None):
	"""
	Override of frappe.email.inbox.get_email_accounts
	Extends the email accounts list to include M365 Email Accounts
	
	Returns email accounts from:
	1. Standard Frappe Email Accounts (via User Emails child table)
	2. M365 Email Accounts where user matches
	3. M365 Email Accounts where user has the assigned role
	
	Args:
		user: User to get email accounts for (defaults to session user)
	
	Returns:
		dict: Email accounts list with M365 accounts included
	"""
	if not user:
		user = frappe.session.user
	
	# Get standard email accounts from original function
	result = _original_get_email_accounts(user)
	email_accounts = result.get("email_accounts", [])
	
	# Get M365 Email Accounts for this user
	m365_accounts = get_m365_email_accounts_for_user(user)
	
	# Find the index where we should insert M365 accounts
	# We want to insert before "Sent Mail", "Spam", "Trash"
	insert_index = len(email_accounts)
	for i, account in enumerate(email_accounts):
		if account.get("email_id") in ["Sent Mail", "Spam", "Trash"]:
			insert_index = i
			break
	
	# Insert M365 accounts before the special folders
	for m365_account in reversed(m365_accounts):
		email_accounts.insert(insert_index, m365_account)
	
	result["email_accounts"] = email_accounts
	
	return result


def get_m365_email_accounts_for_user(user):
	"""
	Get M365 Email Accounts available to a user
	
	Includes accounts where:
	- User field matches the user
	- Role field matches any of the user's roles
	- Account has enable_outgoing enabled
	
	Args:
		user: User to get accounts for
	
	Returns:
		list: List of M365 email accounts in the format expected by email dialog
	"""
	if not user:
		user = frappe.session.user
	
	# Get user's roles
	user_roles = frappe.get_roles(user)
	
	# Build filters for M365 Email Accounts
	# We want accounts where:
	# 1. User matches, OR
	# 2. Role matches any of user's roles
	# AND enable_outgoing is enabled
	
	filters = {
		"enable_outgoing": 1
	}
	
	# Get accounts where user matches
	user_accounts = frappe.get_all(
		"M365 Email Account",
		filters={
			**filters,
			"user": user
		},
		fields=["name", "email_address", "account_name"],
		order_by="account_name"
	)
	
	# Get accounts where role matches
	role_accounts = []
	if user_roles:
		role_accounts = frappe.get_all(
			"M365 Email Account",
			filters={
				**filters,
				"role": ["in", user_roles]
			},
			fields=["name", "email_address", "account_name"],
			order_by="account_name"
		)
	
	# Combine and deduplicate
	all_accounts = user_accounts + role_accounts
	seen = set()
	unique_accounts = []
	for account in all_accounts:
		if account.name not in seen:
			seen.add(account.name)
			unique_accounts.append(account)
	
	# Format for email dialog
	# The email dialog expects:
	# - email_account: The account name (used as identifier)
	# - email_id: The email address (displayed to user)
	# - enable_outgoing: Whether outgoing is enabled
	
	formatted_accounts = []
	for account in unique_accounts:
		formatted_accounts.append({
			"email_account": account.name,
			"email_id": account.email_address,
			"enable_outgoing": 1
		})
	
	return formatted_accounts


def apply_monkey_patch():
	"""Apply monkey patch to frappe.email.inbox.get_email_accounts"""
	inbox.get_email_accounts = get_email_accounts_with_m365
	print("M365 Email: Monkey patched frappe.email.inbox.get_email_accounts")

