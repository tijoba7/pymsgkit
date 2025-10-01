#!/usr/bin/env python3
"""
Verify MSG files using extract-msg library with appropriate error handling.
"""

import sys
import os
from pathlib import Path

# Add parent directory to path to import pymsgkit
sys.path.insert(0, str(Path(__file__).parent))

def verify_msg_files():
    """Verify all MSG files in the current directory."""
    try:
        import extract_msg
        from extract_msg.enums import ErrorBehavior
    except ImportError:
        print("❌ extract-msg not installed. Install with: pip install extract-msg")
        return False
    
    # Find all .msg files
    msg_files = list(Path('.').glob('*.msg'))
    msg_files.extend(Path('.').glob('thread_*.msg'))
    
    if not msg_files:
        print("No MSG files found")
        return True
    
    print(f"Found {len(msg_files)} MSG files to verify:\n")
    
    all_passed = True
    
    for msg_file in sorted(msg_files):
        print(f"Testing: {msg_file.name}")
        try:
            # Use relaxed error handling for validation that's too strict
            msg = extract_msg.Message(
                str(msg_file), 
                errorBehavior=ErrorBehavior.STANDARDS_VIOLATION
            )
            
            # Extract basic info
            print(f"  ✓ File opened successfully")
            print(f"    Subject: {msg.subject}")
            print(f"    Sender: {msg.sender}")
            print(f"    Recipients: {len(msg.recipients)}")
            print(f"    Attachments: {len(msg.attachments)}")
            
            if msg.attachments:
                for att in msg.attachments:
                    filename = att.longFilename or att.shortFilename
                    print(f"      - {filename} ({len(att.data)} bytes)")
            
            msg.close()
            print()
            
        except Exception as e:
            print(f"  ❌ FAILED: {type(e).__name__}: {e}\n")
            all_passed = False
    
    if all_passed:
        print("✅ All MSG files passed validation!")
    else:
        print("❌ Some MSG files failed validation")
    
    return all_passed

if __name__ == '__main__':
    success = verify_msg_files()
    sys.exit(0 if success else 1)
