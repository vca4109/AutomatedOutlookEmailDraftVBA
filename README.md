Automated Outlook Email Draft with HTML Body

This Excel VBA script helps generate a pre-filled Outlook email draft with specified recipients, subject, and a styled HTML body. It can be used for consistent document submissions via email, especially when using platforms like Aconex.

How it works:
1) AutomatedOutlookEmailDraft is a simple entry point that calls SendEmailWithAttachment, preloading a default To and CC list of recipients.
2) SendEmailWithAttachment(mailTo, mailCC):
3) Launches Microsoft Outlook from Excel using CreateObject("Outlook.Application").
4) Fills in: To and CC addresses, Subject line
5) A formatted HTML email body with placeholders like "Our ref.#" and "Please find attached..."
6) Opens the email in draft mode using .Display, allowing the user to manually attach files before sending.

Notes:
1) The macro does not automatically send the email. It opens the draft so you can review or attach files.
2) Email content is styled using basic HTML (Calibri font, colored reference text).
3) Ideal for repeatable submissions with consistent formatting.

Example usage:
1) Call SendEmailWithAttachment("abc@xyz.com", "dxm@xyz.com; tvl@xyz.com; plv@xyz.com; qaq@xyz.com")
2) This will open a new Outlook draft addressed to abc@xyz.com, with multiple CC recipients and a predefined subject and body ready to go.

