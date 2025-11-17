# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Migration patch to add m365_timezone field to Event doctype
"""

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields


def execute():
	"""
	Add m365_timezone custom field to Event doctype
	"""
	custom_fields = {
		"Event": [
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
	
	print("Added m365_timezone field to Event doctype")

