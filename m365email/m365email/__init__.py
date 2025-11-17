# Copyright (c) 2025, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

import frappe


def boot_session(bootinfo):
	"""
	Called when user session is created
	Apply monkey patches and add M365 email accounts to boot info
	"""
	# Apply monkey patch for email accounts
	from m365email.m365email.inbox_override import apply_monkey_patch
	apply_monkey_patch()


# Apply monkey patches on module import
from m365email.m365email.inbox_override import apply_monkey_patch
apply_monkey_patch()
