"""
Email threading example - create conversation chains
"""

from pymsgkit import MSGWriter, PropertyTag

def main():
    # First message in thread
    msg1 = MSGWriter()
    msg1.set_subject("Project Discussion")
    msg1.set_sender("alice@company.com", "Alice")
    msg1.add_recipient("bob@company.com", "Bob")
    msg1.set_body("Hi Bob, what do you think about the new proposal?")
    msg1.set_conversation_index()  # Start new conversation
    msg1.save("thread_01.msg")
    print("✓ Created thread_01.msg")

    # Get conversation index for threading
    conv_index = msg1.properties[PropertyTag.PR_CONVERSATION_INDEX].value

    # Reply 1
    msg2 = MSGWriter()
    msg2.set_subject("RE: Project Discussion")
    msg2.set_sender("bob@company.com", "Bob")
    msg2.add_recipient("alice@company.com", "Alice")
    msg2.set_body("I think it looks great! Let's move forward.")
    msg2.set_conversation_index(conv_index)  # Link to thread
    msg2.save("thread_02.msg")
    print("✓ Created thread_02.msg (reply)")

    # Reply 2
    msg3 = MSGWriter()
    msg3.set_subject("RE: Project Discussion")
    msg3.set_sender("alice@company.com", "Alice")
    msg3.add_recipient("bob@company.com", "Bob")
    msg3.set_body("Perfect! I'll schedule a meeting.")
    msg3.set_conversation_index(conv_index)  # Link to same thread
    msg3.save("thread_03.msg")
    print("✓ Created thread_03.msg (reply)")

    print("\n✓ Thread created! Open these in Outlook to see conversation view.")

if __name__ == "__main__":
    main()
