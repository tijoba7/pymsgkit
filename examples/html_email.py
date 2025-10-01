"""
HTML email with inline image example
"""

from pymsgkit import MSGWriter

def main():
    msg = MSGWriter()
    msg.set_subject("HTML Newsletter")
    msg.set_sender("marketing@company.com", "Marketing Team")

    # HTML body with inline image reference
    html_body = """
    <html>
        <body>
            <h1>Welcome to Our Newsletter!</h1>
            <p>Check out our new logo:</p>
            <img src="cid:company_logo" alt="Company Logo" />
            <p>Best regards,<br>The Team</p>
        </body>
    </html>
    """
    msg.set_body(html_body, is_html=True)
    msg.add_recipient("customer@example.com", "Valued Customer")

    # Create a simple 1x1 red pixel PNG for demonstration
    red_pixel_png = bytes([
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
        0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
        0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
        0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D,
        0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
        0x44, 0xAE, 0x42, 0x60, 0x82
    ])

    # Add as inline attachment
    msg.add_attachment(
        filename="logo.png",
        data=red_pixel_png,
        content_id="company_logo",
        mime_type="image/png",
        is_inline=True
    )

    msg.save("html_email.msg")
    print("✓ Created html_email.msg with inline image")

if __name__ == "__main__":
    main()
