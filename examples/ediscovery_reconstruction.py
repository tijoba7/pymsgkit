"""
eDiscovery use case - reconstruct email with custom sender
"""

from pymsgkit import MSGWriter
from datetime import datetime, timezone

def main():
    """
    Reconstruct an email from archived data with original sender preserved.
    This bypasses Outlook's security restrictions that prevent setting
    arbitrary sender addresses.
    """

    # Simulate archived email data
    archived_data = {
        'sender_email': 'original.sender@company.com',
        'sender_name': 'John Original',
        'recipient_email': 'legal@company.com',
        'recipient_name': 'Legal Department',
        'subject': 'Reconstructed Evidence Email',
        'body': 'This email was reconstructed from archive data for legal review.',
        'date': datetime(2024, 1, 15, 10, 30, 0, tzinfo=timezone.utc)
    }

    # Create reconstructed MSG
    msg = MSGWriter()

    # Set original sender (this is why we can't use standard Outlook API!)
    msg.set_sender(
        archived_data['sender_email'],
        archived_data['sender_name']
    )

    msg.set_subject(archived_data['subject'])
    msg.set_body(archived_data['body'])
    msg.add_recipient(
        archived_data['recipient_email'],
        archived_data['recipient_name']
    )

    # Set original timestamp
    from pymsgkit.properties import PropertyTag, PropertyType
    msg.set_property(
        PropertyTag.PR_CLIENT_SUBMIT_TIME,
        PropertyType.PT_SYSTIME,
        archived_data['date']
    )

    msg.save("reconstructed_evidence.msg")
    print("✓ Created reconstructed_evidence.msg")
    print(f"  Sender: {archived_data['sender_name']} <{archived_data['sender_email']}>")
    print("  This demonstrates custom sender control for eDiscovery workflows")

if __name__ == "__main__":
    main()
