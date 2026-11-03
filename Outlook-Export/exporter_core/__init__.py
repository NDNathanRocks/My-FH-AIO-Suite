"""Platform-independent core of the Inbox Email Exporter.

Everything here is testable without Outlook: naming, database migration and
history, the export engine (driven through a ``MailSource`` interface),
reconciliation and reporting. ``outlook_com`` is the only Windows-specific
module and is imported lazily by the GUI.
"""
SCHEMA_VERSION = 2
