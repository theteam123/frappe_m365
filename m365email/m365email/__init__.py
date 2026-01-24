# Copyright (c) 2025, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

import frappe


def boot_session(bootinfo):
	"""
	Called when user session is created.

	Note: With the Email Account integration (service='M365'),
	monkey patches are no longer needed. M365 accounts are now
	regular Email Accounts and are automatically included in
	Frappe's email account lists.
	"""
	pass


# Monkey patches are no longer needed - M365 accounts are now
# integrated into Email Account with service='M365'
