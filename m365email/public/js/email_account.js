// Copyright (c) 2025, Zeke Tierney and contributors
// For license information, please see license.txt

/**
 * Client-side script for Email Account M365 integration
 *
 * This script handles:
 * - Hiding SMTP/IMAP fields when M365 service is selected
 * - Showing M365-specific fields when M365 is selected
 * - Setting appropriate defaults for M365 accounts
 * - Managing field visibility based on service type
 */

// Add M365 to email defaults (similar to Gmail, Outlook, etc.)
frappe.email_defaults = frappe.email_defaults || {};
frappe.email_defaults['M365'] = {
	// M365 doesn't use SMTP/IMAP, so we disable these settings
	use_imap: 0,
	use_ssl: 0,
	use_tls: 0,
	use_starttls: 0,
	use_ssl_for_outgoing: 0,
	no_smtp_authentication: 1,
	// Clear server settings
	email_server: '',
	incoming_port: '',
	smtp_server: '',
	smtp_port: ''
};

// Fields to hide when M365 is selected (SMTP/IMAP specific)
// NOTE: We keep enable_incoming, default_incoming, enable_outgoing, default_outgoing visible
// as these are used by M365 as well. We also keep the section breaks visible but hide
// the SMTP/IMAP-specific fields within them.
const SMTP_IMAP_FIELDS = [
	// Authentication section fields
	'authentication_column',
	'auth_method',
	'backend_app_flow',
	'authorize_api_access',
	'password',
	'awaiting_password',
	'ascii_encode_password',
	'column_break_10',
	'connected_app',
	'connected_user',
	'login_id_is_different',
	'login_id',
	// Mailbox settings (IMAP/POP) - hide IMAP/POP specific fields but NOT the section or enable_incoming/default_incoming
	// 'mailbox_settings', // KEEP - section break needed for enable_incoming/default_incoming
	// 'enable_incoming', // KEEP - used by M365
	// 'default_incoming', // KEEP - used by M365
	'use_imap',
	'use_ssl',
	'use_starttls',
	'email_server',
	'incoming_port',
	'column_break_18',
	'attachment_limit',
	'email_sync_option',
	'initial_sync_count',
	// IMAP folder section
	'section_break_25',
	'imap_folder',
	// Email processing section
	'section_break_12',
	'append_emails_to_sent_folder',
	'sent_folder_name',
	'append_to',
	'create_contact',
	'enable_automatic_linking',
	// Notification section (keep for M365? or hide?)
	'section_break_13',
	'notify_if_unreplied',
	'unreplied_for_mins',
	'send_notification_to',
	// Outgoing settings - hide SMTP specific fields but NOT the section or enable_outgoing/default_outgoing
	// 'outgoing_mail_settings', // KEEP - section break needed for enable_outgoing/default_outgoing
	// 'enable_outgoing', // KEEP - used by M365
	// 'default_outgoing', // KEEP - used by M365
	'use_tls',
	'use_ssl_for_outgoing',
	'smtp_server',
	'smtp_port',
	'column_break_38',
	'always_use_account_email_id_as_sender',
	'always_use_account_name_as_sender_name',
	'send_unsubscribe_message',
	'track_email_status',
	'no_smtp_authentication'
];

// M365-specific fields to show when M365 is selected
const M365_FIELDS = [
	// M365 Settings Section
	'm365_settings_section',
	'm365_service_principal',
	'm365_account_type',
	'm365_column_break_1',
	'm365_sync_events',
	// M365 Sync Settings Section
	'm365_sync_settings_section',
	'm365_sync_from_date',
	'm365_user_timezone',
	'm365_auto_create_contact',
	'm365_sync_column_break',
	'm365_sync_attachments',
	'm365_max_attachment_size',
	// M365 Folder Sync Section
	'm365_folder_section',
	'm365_folder_filter',
	// M365 Sync Status Section
	'm365_status_section',
	'm365_last_sync_time',
	'm365_last_sync_status',
	'm365_status_column_break',
	'm365_sync_error_message',
	'm365_delta_tokens'
];

frappe.ui.form.on('Email Account', {
	refresh: function(frm) {
		// Apply field visibility on form load
		frm.trigger('toggle_m365_fields');
	},

	service: function(frm) {
		// Toggle field visibility when service changes
		frm.trigger('toggle_m365_fields');

		// Apply defaults for the selected service
		if (frm.doc.service === 'M365') {
			// Set M365 defaults
			$.each(frappe.email_defaults['M365'], function(key, value) {
				frm.set_value(key, value);
			});
		}
	},

	enable_incoming: function(frm) {
		// Re-toggle M365 fields when enable_incoming changes
		// (some M365 fields depend on enable_incoming)
		frm.trigger('toggle_m365_fields');
	},

	m365_account_type: function(frm) {
		// Re-toggle when account type changes
		frm.trigger('toggle_m365_fields');
	},

	toggle_m365_fields: function(frm) {
		const is_m365 = frm.doc.service === 'M365';

		// Hide/show SMTP/IMAP fields based on service type
		SMTP_IMAP_FIELDS.forEach(function(fieldname) {
			if (frm.fields_dict[fieldname]) {
				frm.toggle_display(fieldname, !is_m365);
			}
		});

		// Show/hide M365 fields based on service type
		M365_FIELDS.forEach(function(fieldname) {
			if (frm.fields_dict[fieldname]) {
				frm.toggle_display(fieldname, is_m365);
			}
		});

		// When M365 is selected, update section labels and hide domain
		if (is_m365) {
			frm.toggle_display('domain', false);

			// Update section labels to be more appropriate for M365
			// For section breaks, we need to update both the df.label and the DOM element
			if (frm.fields_dict['mailbox_settings']) {
				frm.fields_dict['mailbox_settings'].df.label = 'Incoming Settings';
				// Update the DOM element directly
				let section = frm.fields_dict['mailbox_settings'].$wrapper;
				if (section) {
					section.find('.section-head, .col-sm-12.section-head').text('Incoming Settings');
				}
			}
			if (frm.fields_dict['outgoing_mail_settings']) {
				frm.fields_dict['outgoing_mail_settings'].df.label = 'Outgoing Settings';
				// Update the DOM element directly
				let section = frm.fields_dict['outgoing_mail_settings'].$wrapper;
				if (section) {
					section.find('.section-head, .col-sm-12.section-head').text('Outgoing Settings');
				}
			}

			// Also update the description on enable_outgoing to remove SMTP reference
			frm.set_df_property('enable_outgoing', 'description', 'Enable sending emails via M365 Graph API');

			// Show/hide sync settings based on enable_incoming
			const show_sync_settings = frm.doc.enable_incoming;
			frm.toggle_display('m365_sync_settings_section', show_sync_settings);
			frm.toggle_display('m365_sync_from_date', show_sync_settings);
			frm.toggle_display('m365_user_timezone', show_sync_settings);
			frm.toggle_display('m365_auto_create_contact', show_sync_settings);
			frm.toggle_display('m365_sync_column_break', show_sync_settings);
			frm.toggle_display('m365_sync_attachments', show_sync_settings);
			frm.toggle_display('m365_max_attachment_size', show_sync_settings);
			frm.toggle_display('m365_folder_section', show_sync_settings);
			frm.toggle_display('m365_folder_filter', show_sync_settings);
		} else {
			// Restore original section labels when not M365
			if (frm.fields_dict['mailbox_settings']) {
				frm.fields_dict['mailbox_settings'].df.label = 'Incoming (POP/IMAP) Settings';
				let section = frm.fields_dict['mailbox_settings'].$wrapper;
				if (section) {
					section.find('.section-head, .col-sm-12.section-head').text('Incoming (POP/IMAP) Settings');
				}
			}
			if (frm.fields_dict['outgoing_mail_settings']) {
				frm.fields_dict['outgoing_mail_settings'].df.label = 'Outgoing (SMTP) Settings';
				let section = frm.fields_dict['outgoing_mail_settings'].$wrapper;
				if (section) {
					section.find('.section-head, .col-sm-12.section-head').text('Outgoing (SMTP) Settings');
				}
			}
			frm.set_df_property('enable_outgoing', 'description', 'SMTP Settings for outgoing emails');
		}
	}
});

