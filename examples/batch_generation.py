"""
Batch email generation example
"""

from pymsgkit import create_email
import csv
from io import StringIO

def main():
    # Sample recipient data (in real scenario, load from CSV file)
    recipients_csv = """name,email,account_id
John Doe,john@example.com,12345
Jane Smith,jane@example.com,12346
Bob Johnson,bob@example.com,12347"""

    recipients = csv.DictReader(StringIO(recipients_csv))

    for recipient in recipients:
        msg = create_email(
            subject=f"Account Statement for {recipient['name']}",
            body=f"""Dear {recipient['name']},

Your monthly account statement is ready.

Account ID: {recipient['account_id']}

Thank you for your business!

Best regards,
Customer Service Team
""",
            sender_email="noreply@company.com",
            sender_name="Customer Service",
            to_recipients=[(recipient['email'], recipient['name'])]
        )

        filename = f"statement_{recipient['account_id']}.msg"
        msg.save(filename)
        print(f"✓ Created {filename}")

    print(f"\n✓ Generated emails for all recipients")

if __name__ == "__main__":
    main()
