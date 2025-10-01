# PyMsgKit

[![Python 3.7+](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Pure Python library for creating Microsoft Outlook MSG files without requiring Outlook or external dependencies. Designed for eDiscovery, forensic email reconstruction, and programmatic email generation with full control over sender properties and MAPI attributes.

## Features

- ✅ **Create MSG files from scratch** - No Outlook installation required
- ✅ **Custom sender control** - Bypass Outlook restrictions for forensic reconstruction
- ✅ **Email threading** - Support for conversation chains and replies
- ✅ **Attachments** - Regular and inline attachments with content IDs
- ✅ **HTML & Plain Text** - Full support for both body formats
- ✅ **Pure Python** - No external dependencies, works on Linux/Mac/Windows
- ✅ **Full MAPI property support** - Complete control over message properties

## Installation

```bash
pip install pymsgkit
```

Or install from source:

```bash
git clone https://github.com/yourusername/pymsgkit.git
cd pymsgkit
pip install -e .
```

## Quick Start

### Simple Email

```python
from pymsgkit import create_email

msg = create_email(
    subject="Hello World",
    body="This is a test email",
    sender_email="sender@example.com",
    sender_name="John Doe",
    to_recipients=[("recipient@example.com", "Jane Smith")]
)
msg.save("email.msg")
```

### Advanced Usage

```python
from pymsgkit import MSGWriter, RecipientType

# Create message
msg = MSGWriter()
msg.set_subject("Project Update")
msg.set_sender("manager@company.com", "Project Manager")
msg.set_body("<h1>Status Report</h1><p>All systems operational.</p>", is_html=True)

# Add recipients
msg.add_recipient("team@company.com", "Development Team", RecipientType.TO)
msg.add_recipient("stakeholder@company.com", "Stakeholder", RecipientType.CC)

# Add attachment
with open("report.pdf", "rb") as f:
    msg.add_attachment("report.pdf", f.read(), mime_type="application/pdf")

# Save
msg.save("project_update.msg")
```

### Email Threading

```python
from pymsgkit import MSGWriter, PropertyTag

# First message
msg1 = MSGWriter()
msg1.set_subject("Initial Message")
msg1.set_sender("alice@example.com", "Alice")
msg1.add_recipient("bob@example.com", "Bob")
msg1.set_body("Starting a conversation thread")
msg1.set_conversation_index()  # Create new thread
msg1.save("thread_01.msg")

# Reply
conversation_index = msg1.properties[PropertyTag.PR_CONVERSATION_INDEX].value

msg2 = MSGWriter()
msg2.set_subject("RE: Initial Message")
msg2.set_sender("bob@example.com", "Bob")
msg2.add_recipient("alice@example.com", "Alice")
msg2.set_body("Reply to the conversation")
msg2.set_conversation_index(conversation_index)  # Link to thread
msg2.save("thread_02.msg")
```

### Inline Images (HTML Email)

```python
from pymsgkit import MSGWriter

msg = MSGWriter()
msg.set_subject("Newsletter")
msg.set_sender("marketing@company.com", "Marketing Team")
msg.set_body(
    '<html><body><p>Check out our logo:</p><img src="cid:logo" /></body></html>',
    is_html=True
)
msg.add_recipient("customer@example.com", "Customer")

# Add inline image
with open("logo.png", "rb") as f:
    msg.add_attachment(
        "logo.png",
        f.read(),
        content_id="logo",
        mime_type="image/png",
        is_inline=True
    )

msg.save("newsletter.msg")
```

## Use Cases

### eDiscovery & Forensics

PyMsgKit is designed for eDiscovery workflows where you need to reconstruct emails with specific sender addresses that standard APIs won't allow:

```python
# Reconstruct email from archived data with original sender preserved
msg = MSGWriter()
msg.set_sender("original.sender@company.com", "Original Sender")  # ✅ Works!
msg.set_subject("Reconstructed Email")
msg.set_body("This email was reconstructed from archive data.")
msg.add_recipient("legal@company.com", "Legal Team")
msg.save("evidence_email.msg")
```

Standard Outlook APIs won't let you set arbitrary senders, but MSG files created with PyMsgKit bypass this restriction by writing directly to the file structure.

### Automated Email Generation

Generate templated emails programmatically:

```python
import csv
from pymsgkit import create_email

with open("recipients.csv") as f:
    for row in csv.DictReader(f):
        msg = create_email(
            subject=f"Welcome, {row['name']}!",
            body=f"Dear {row['name']},\n\nWelcome to our service...",
            sender_email="welcome@company.com",
            sender_name="Welcome Team",
            to_recipients=[(row['email'], row['name'])]
        )
        msg.save(f"welcome_{row['id']}.msg")
```

## API Reference

### MSGWriter

Main class for creating MSG files.

**Methods:**

- `set_subject(subject: str)` - Set email subject
- `set_body(body: str, is_html: bool = False)` - Set email body
- `set_sender(email: str, name: str = "", addr_type: str = "SMTP")` - Set sender
- `add_recipient(email: str, name: str = "", recipient_type: RecipientType = RecipientType.TO)` - Add recipient
- `add_attachment(filename: str, data: bytes, content_id: str = None, mime_type: str = None, is_inline: bool = False)` - Add attachment
- `set_conversation_index(parent_index: bytes = None)` - Set threading
- `set_property(prop_tag: int, prop_type: int, value: Any)` - Set custom MAPI property
- `save(filepath: str)` - Save to MSG file

### Helper Functions

- `create_email(...)` - Quick email creation with sensible defaults

### Enums

- `RecipientType.TO`, `RecipientType.CC`, `RecipientType.BCC`
- `PropertyType` - MAPI property types
- `AttachMethod` - Attachment methods

## Technical Details

PyMsgKit implements the Microsoft specifications:

- **MS-CFB**: Compound File Binary Format
- **MS-OXMSG**: Outlook MSG File Format
- **MS-OXPROPS**: Exchange Server Protocols Property Tags

The library creates valid MSG files by:

1. Building a Compound File Binary (CFB) container structure
2. Encoding MAPI properties according to type specifications
3. Creating proper directory hierarchies for recipients and attachments
4. Writing property streams with correct naming conventions

## Requirements

- Python 3.7+
- No external dependencies (pure Python standard library)

## Testing

```bash
# Run tests
python -m pytest tests/

# Run specific test
python -m pytest tests/test_basic.py -v
```

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Submit a pull request

## License

MIT License - see LICENSE file for details

## Acknowledgments

- Microsoft for documenting the MSG and CFB specifications
- The eDiscovery community for use case feedback
- MsgKit (C#) for architectural inspiration

## Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/pymsgkit/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/pymsgkit/discussions)

## Changelog

### v1.0.0 (2025-01-XX)
- Initial release
- Full MSG file creation support
- Custom sender control
- Threading support
- Attachment support (regular and inline)
- HTML and plain text bodies
