# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Custom fields for M365 Email Integration
"""

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields


def create_m365_custom_fields():
	"""
	Create custom fields for M365 Email Integration
	"""
	custom_fields = {
		"Communication": [
			{
				"fieldname": "m365_message_id",
				"label": "M365 Message ID",
				"fieldtype": "Data",
				"length": 500,
				"insert_after": "uid",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1,
				"unique": 1
			},
			{
				"fieldname": "m365_email_account",
				"label": "M365 Email Account",
				"fieldtype": "Link",
				"options": "M365 Email Account",
				"insert_after": "m365_message_id",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1
			}
		],
		"Email Queue": [
			{
				"fieldname": "m365_send",
				"label": "Send via M365",
				"fieldtype": "Check",
				"insert_after": "status",
				"read_only": 1,
				"no_copy": 1,
				"default": "0"
			},
			{
				"fieldname": "m365_account",
				"label": "M365 Account",
				"fieldtype": "Link",
				"options": "M365 Email Account",
				"insert_after": "m365_send",
				"read_only": 1,
				"no_copy": 1
			}
		],
		"Event": [
			{
				"fieldname": "m365_event_id",
				"label": "M365 Event ID",
				"fieldtype": "Data",
				"length": 500,
				"insert_after": "event_type",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1,
				"unique": 1
			},
			{
				"fieldname": "m365_email_account",
				"label": "M365 Email Account",
				"fieldtype": "Link",
				"options": "M365 Email Account",
				"insert_after": "m365_event_id",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1
			},
			{
				"fieldname": "m365_icaluid",
				"label": "M365 iCalUId",
				"fieldtype": "Data",
				"length": 500,
				"insert_after": "m365_email_account",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1
			},
			{
				"fieldname": "m365_timezone",
				"label": "M365 Timezone",
				"fieldtype": "Data",
				"length": 100,
				"insert_after": "m365_icaluid",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 0,
				"description": "Original timezone from Microsoft 365"
			}
		]
	}

	create_custom_fields(custom_fields, update=True)
	frappe.db.commit()


def execute():
	"""
	Execute function for patches
	"""
	create_m365_custom_fields()

