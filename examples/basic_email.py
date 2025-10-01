"""
Basic email creation example
"""

from pymsgkit import create_email

def main():
    # Create a simple email
    msg = create_email(
        subject="Hello from PyMsgKit",
        body="This is a test email created with PyMsgKit!",
        sender_email="sender@example.com",
        sender_name="John Doe",
        to_recipients=[
            ("recipient1@example.com", "Jane Smith"),
            ("recipient2@example.com", "Bob Johnson")
        ],
        cc_recipients=[
            ("cc@example.com", "CC Recipient")
        ]
    )

    msg.save("basic_email.msg")
    print("✓ Created basic_email.msg")

if __name__ == "__main__":
    main()
