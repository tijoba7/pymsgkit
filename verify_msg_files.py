#!/usr/bin/env python3
"""
Verify generated MSG files using extract-msg library
"""

import extract_msg
import os
import sys

def verify_msg(filename):
    """Verify a single MSG file"""
    print(f"\n{'='*60}")
    print(f"Verifying: {filename}")
    print(f"{'='*60}")

    if not os.path.exists(filename):
        print(f"❌ File not found: {filename}")
        return False

    try:
        msg = extract_msg.Message(filename)

        # Check basic properties
        print(f"✅ File opened successfully")
        print(f"   Subject: {msg.subject}")
        print(f"   Sender: {msg.sender}")
        print(f"   To: {msg.to}")
        print(f"   CC: {msg.cc if msg.cc else '(none)'}")
        print(f"   Date: {msg.date}")

        # Check body
        if msg.body:
            print(f"   Body length: {len(msg.body)} chars")
            print(f"   Body preview: {msg.body[:100]}...")

        # Check HTML body
        if msg.htmlBody:
            print(f"   HTML body length: {len(msg.htmlBody)} bytes")

        # Check attachments
        attachments = msg.attachments
        if attachments:
            print(f"   Attachments: {len(attachments)}")
            for i, att in enumerate(attachments):
                print(f"      [{i}] {att.longFilename or att.shortFilename} ({len(att.data)} bytes)")
                if hasattr(att, 'cid') and att.cid:
                    print(f"          Content-ID: {att.cid}")
        else:
            print(f"   Attachments: 0")

        # Check message class
        if hasattr(msg, 'messageClass'):
            print(f"   Message class: {msg.messageClass}")

        # Check conversation properties
        if hasattr(msg, 'conversationTopic'):
            print(f"   Conversation topic: {msg.conversationTopic}")

        msg.close()
        return True

    except Exception as e:
        print(f"❌ Error reading MSG file: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Verify all test MSG files"""
    test_files = [
        'basic_email.msg',
        'html_email.msg',
        'thread_01.msg',
        'thread_02.msg',
        'thread_03.msg',
        'reconstructed_evidence.msg'
    ]

    print("MSG File Verification using extract-msg")
    print("=" * 60)

    results = {}
    for filename in test_files:
        results[filename] = verify_msg(filename)

    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")

    passed = sum(1 for v in results.values() if v)
    total = len(results)

    for filename, success in results.items():
        status = "✅ PASS" if success else "❌ FAIL"
        print(f"{status}: {filename}")

    print(f"\n{passed}/{total} files verified successfully")

    if passed == total:
        print("\n🎉 All MSG files are valid and readable!")
        return 0
    else:
        print(f"\n⚠️  {total - passed} file(s) failed verification")
        return 1

if __name__ == '__main__':
    sys.exit(main())
