# O4U-Match-Emails

## Overview
This is a Python-based desktop application built using Tkinter to send personalized emails to mentors and mentees participating in the O4U Mentorship Program. The program allows you to upload a CSV file with details of the mentorship pairs, edit the email body, preview emails, and send them in bulk while logging the status of each email sent.

![image](https://github.com/user-attachments/assets/faf9b8d6-8613-4df1-9a83-84f767d32fe3)

## Features
- Upload CSV files with mentorship details.
- Edit the default email subject and body.
- Preview the personalized emails before sending.
- Send emails to both mentors and mentees.
- Real-time logging of email sending status.
- Validation of email addresses to avoid sending to invalid emails.
- Save log and failed emails to a file for further review.

## Installation
1. Ensure Python 3.x is installed on your system.
2. Install the required dependencies by running the following:

   ```
   pip install pandas smtplib email_validator
   ```

4. Download the script and place it in your working directory.

## How to Use
1. **Run the Application:**
   Run the script in a terminal using:
   ```
   python email_sender_app.py
   ```

2. **Upload CSV File:**
   - Click the “Upload CSV” button.
   - Choose a CSV file with the following columns: `MentorEmail`, `MenteeEmail`, `MentorFirstName`, `MentorLastName`, `MenteeFistName`, `MenteeLastName`, `JobTitle`, `PlaceOfWork`, `Major`, and `University`.

3. **Edit Email Body (Optional):**
   - Use the “Edit Email Body” button to customize the email content.
   - Make sure to use placeholders like `{MentorFirstName}`, `{MenteeFistName}`, etc., which will be replaced by actual values from the CSV.

4. **Preview Email:**
   - Click “Preview Email” to see how the emails will look.
   - Navigate through different mentorship pairs using “Prev” and “Next” buttons.

5. **Send Emails:**
   - Click “Send Emails,” and you will be prompted for your email address and password.
   - The application will display real-time logs in a popup window showing which emails were successfully sent and any failures.

6. **Review Failed Emails:**
   - Any emails that fail due to invalid addresses or other issues will be saved in a separate CSV file (`failed_emails.csv`).

## CSV Requirements
The CSV file should include the following columns:
- `MentorEmail`
- `MenteeEmail`
- `MentorFirstName`
- `MentorLastName`
- `MenteeFistName`
- `MenteeLastName`
- `JobTitle`
- `PlaceOfWork`
- `Major`
- `University`

## Logging and Error Handling
- The application logs the sending status for each email in real time.
- If an email cannot be sent (e.g., due to invalid email addresses), the affected rows are logged, and the application skips them.
- A log file (`email_send_log.txt`) and a CSV of failed rows (`failed_emails.csv`) are generated for your records.

## Known Limitations
- Only supports Gmail SMTP server (`smtp.gmail.com`).
- Requires enabling less secure apps in your Gmail account settings.
- The email content is basic and does not support HTML formatting.

## Author
Jared Bach
