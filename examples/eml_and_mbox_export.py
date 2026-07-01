"""
Export messages to standard EML and MBOX formats.

MBOX is the practical way to collect many messages into a single, portable
archive that Outlook, Thunderbird, Apple Mail, and eDiscovery tools can import
-- the sensible alternative to a from-scratch .pst writer.
"""

from pymsgkit import MSGWriter, MboxWriter, RecipientType


def build_message(subject, body, sender, to):
    msg = MSGWriter()
    msg.set_subject(subject)
    msg.set_sender(sender, sender.split("@")[0].title())
    msg.add_recipient(to, to.split("@")[0].title(), RecipientType.TO)
    msg.set_body(body)
    return msg


def main():
    # 1) Single message -> .eml
    msg = build_message(
        "Welcome aboard",
        "Thanks for joining us!",
        "welcome@company.com",
        "newuser@example.com",
    )
    msg.save("welcome.msg")   # still available as MSG
    msg.save_eml("welcome.eml")
    print("✓ Wrote welcome.msg and welcome.eml")

    # 2) Many messages -> one .mbox archive
    box = MboxWriter()
    for i in range(1, 4):
        box.add(build_message(
            subject=f"Newsletter #{i}",
            body=f"This is edition {i} of our newsletter.",
            sender="news@company.com",
            to=f"subscriber{i}@example.com",
        ))
    box.save("newsletters.mbox")
    print(f"✓ Wrote newsletters.mbox with {len(box)} messages")


if __name__ == "__main__":
    main()
