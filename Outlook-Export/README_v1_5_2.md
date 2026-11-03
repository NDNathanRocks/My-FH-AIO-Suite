# Inbox Email Exporter v1.5.2 database fix

1. Put `upgrade_v1_5_2_database_fix.py` beside `inbox_email_exporter_v1_5_1.py`.
2. Run `py upgrade_v1_5_2_database_fix.py`.
3. Run `py -m py_compile inbox_email_exporter_v1_5_2.py`.
4. Run `py build_v1_5_2.py`.

The app now checks and migrates the database at startup, before Outlook scanning begins. It creates a timestamped backup before adding missing columns. Existing columns and records are retained.
