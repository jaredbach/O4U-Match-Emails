import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, scrolledtext
from tkinter import Toplevel, Text, Scrollbar, END
from tkinter import ttk
import pandas as pd
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email_validator import validate_email, EmailNotValidError
import base64
import smtplib
import os
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://mail.google.com/'
]

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("O4U Match Email Sender")
        self.root.geometry("1200x800")

        self.csv_file_path = ''
        self.df = None
        self.current_page = 0
        self.rows_per_page = 25
        self.current_preview_index = 0

        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 587

        self.body_template = """Hey {MentorFirstName} and {MenteeFirstName}!\n\nIn preparation for program kickoff, we are so excited to formally introduce you to your Spring 2026 O4U mentorship program pairing!\n\nMentor: {MentorFirstName} {MentorLastName}, {JobTitle} at {PlaceOfWork}\nMentee: {MenteeFirstName} {MenteeLastName}, studying {Major} at {University}\n\nPlease feel free to use this email thread to introduce yourselves to each other before kickoff on Thursday, [Put Date Here] at 7:30pm EST -- and get your first meeting in the books! Remember that we expect mentors and mentees to meet twice a month. \n\nYou should have already received the training materials, including the discussion guide, student 1-pager, and the recording of last week's training sessions. Please review if you were not able to attend the trainings. Please also keep these dates blocked on your calendar:\n\nMidway Event: Thursday,[Put Date Here] @ 7:30pm ET\nClosing Event: Thursday, [Put Date Here] @ 7:30pm ET\n\nIf you don't hear back from your mentor / mentee within two days, feel free to reach out to mentoring@outforundergrad.org. \n\nBest,\nJared"""
        self.email_subject = "Spring 2026 O4U Mentorship Program Pairing"

        self.creds = None
        self.sender_email = None

        self.build_ui()

    def build_ui(self):
        self.title_label = tk.Label(self.root, text="O4U Match Email Sender", font=("Helvetica", 16, "bold"))
        self.title_label.pack(pady=10)

        self.upload_btn = tk.Button(self.root, text="Upload CSV", command=self.upload_csv)
        self.upload_btn.pack(pady=10)

        self.subject_label = tk.Label(self.root, text="Email Subject:")
        self.subject_label.pack()
        self.subject_entry = tk.Entry(self.root, width=80)
        self.subject_entry.insert(0, self.email_subject)
        self.subject_entry.pack(pady=10)

        self.cc_label = tk.Label(self.root, text="CC Email:")
        self.cc_label.pack()
        self.cc_entry = tk.Entry(self.root, width=80)
        self.cc_entry.insert(0, 'mentoring@outforundergrad.org')
        self.cc_entry.pack(pady=10)

        self.tree_frame = tk.Frame(self.root)
        self.tree_frame.pack(pady=10, expand=True, fill=tk.BOTH)

        self.tree_scroll_y = tk.Scrollbar(self.tree_frame)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_scroll_x = tk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(self.tree_frame, columns=[], show='headings',
                                 yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(expand=True, fill=tk.BOTH)

        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.pagination_frame = tk.Frame(self.root)
        self.pagination_frame.pack(pady=10)

        self.prev_btn = tk.Button(self.pagination_frame, text="<< Prev", command=self.show_prev_page)
        self.prev_btn.grid(row=0, column=0, padx=5)

        self.page_label = tk.Label(self.pagination_frame, text="Page 0")
        self.page_label.grid(row=0, column=1, padx=5)

        self.next_btn = tk.Button(self.pagination_frame, text="Next >>", command=self.show_next_page)
        self.next_btn.grid(row=0, column=2, padx=5)

        self.edit_body_btn = tk.Button(self.root, text="Edit Email Body", command=self.edit_email_body, state=tk.DISABLED)
        self.edit_body_btn.pack(pady=10)

        self.preview_btn = tk.Button(self.root, text="Preview Email", command=self.preview_email, state=tk.DISABLED)
        self.preview_btn.pack(pady=10)

        self.send_btn = tk.Button(self.root, text="Send Emails (OAuth2 Login)", command=self.authenticate_and_send, state=tk.DISABLED)
        self.send_btn.pack(pady=10)

    def upload_csv(self):
        self.csv_file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not self.csv_file_path:
            return
        try:
            self.df = pd.read_csv(self.csv_file_path)
            required_columns = ['MentorEmail', 'MenteeEmail', 'MentorFirstName', 'MentorLastName',
                                'MenteeFirstName', 'MenteeLastName', 'JobTitle', 'PlaceOfWork', 'Major', 'University']
            if not all(col in self.df.columns for col in required_columns):
                raise ValueError("CSV is not correctly formatted. Please check the required columns.")

            self.current_page = 0
            self.update_treeview()
            self.page_label.config(text=f"Page {self.current_page + 1}")

            messagebox.showinfo("Success", "CSV uploaded successfully!")
            self.edit_body_btn.config(state=tk.NORMAL)
            self.preview_btn.config(state=tk.NORMAL)
            self.send_btn.config(state=tk.NORMAL)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to upload CSV: {e}")

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        if self.df is not None:
            columns = self.df.columns.tolist()
            self.tree.config(columns=columns)
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120)

            start_row = self.current_page * self.rows_per_page
            end_row = min(start_row + self.rows_per_page, len(self.df))
            for index, row in self.df.iloc[start_row:end_row].iterrows():
                self.tree.insert("", tk.END, values=row.tolist())

            self.prev_btn.config(state=tk.NORMAL if self.current_page > 0 else tk.DISABLED)
            self.next_btn.config(state=tk.NORMAL if end_row < len(self.df) else tk.DISABLED)

    def show_prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.update_treeview()
            self.page_label.config(text=f"Page {self.current_page + 1}")

    def show_next_page(self):
        if (self.current_page + 1) * self.rows_per_page < len(self.df):
            self.current_page += 1
            self.update_treeview()
            self.page_label.config(text=f"Page {self.current_page + 1}")

    def edit_email_body(self):
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Email Body")

        text_widget = scrolledtext.ScrolledText(edit_window, width=100, height=30)
        text_widget.insert(tk.END, self.body_template)
        text_widget.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)

        def save_body():
            self.body_template = text_widget.get("1.0", tk.END).strip()
            edit_window.destroy()

        save_button = tk.Button(edit_window, text="Save", command=save_body)
        save_button.pack(pady=10)

    def preview_email(self):
        if self.df is None:
            messagebox.showerror("Error", "No CSV file loaded.")
            return

        preview_window = tk.Toplevel(self.root)
        preview_window.title("Email Preview")
        preview_window.geometry("700x600")

        self.current_preview_index = 0

        email_text = tk.Text(preview_window, wrap=tk.WORD)
        email_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        def show_email(index):
            if 0 <= index < len(self.df):
                row = self.df.iloc[index]
                email_body = self.body_template.format(**row)
                email_text.delete("1.0", tk.END)
                email_text.insert(tk.END, f"Subject: {self.subject_entry.get()}\n\n{email_body}")
                self.current_preview_index = index
                prev_btn.config(state=tk.NORMAL if index > 0 else tk.DISABLED)
                next_btn.config(state=tk.NORMAL if index < len(self.df) - 1 else tk.DISABLED)

        def prev_email():
            show_email(self.current_preview_index - 1)

        def next_email():
            show_email(self.current_preview_index + 1)

        control_frame = tk.Frame(preview_window)
        control_frame.pack(pady=5)

        prev_btn = tk.Button(control_frame, text="<< Prev", command=prev_email)
        prev_btn.grid(row=0, column=0, padx=5)

        next_btn = tk.Button(control_frame, text="Next >>", command=next_email)
        next_btn.grid(row=0, column=1, padx=5)

        show_email(0)

    def authenticate_and_send(self):
        # Run OAuth flow and get creds
        threading.Thread(target=self.oauth_and_send_thread, daemon=True).start()

    def oauth_and_send_thread(self):
        try:
            # Load token if exists
            if os.path.exists('token.json'):
                self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)

            # If no valid creds, do OAuth flow
            if not self.creds or not self.creds.valid:
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    self.creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                    self.creds = flow.run_local_server(port=0)

                with open('token.json', 'w') as token_file:
                    token_file.write(self.creds.to_json())

            # Extract email from creds info (should be the Gmail account)
            from googleapiclient.discovery import build
            service = build('gmail', 'v1', credentials=self.creds)
            profile = service.users().getProfile(userId='me').execute()
            self.sender_email = profile['emailAddress']

            self.show_logging_window()
            self.send_all_emails()

        except Exception as e:
            messagebox.showerror("OAuth2 Error", f"Failed OAuth2 authentication or sending emails:\n{e}")

    def show_logging_window(self):
        self.log_window = Toplevel(self.root)
        self.log_window.title("Email Sending Logs")
        self.log_window.geometry("700x500")

        self.log_text = scrolledtext.ScrolledText(self.log_window, state='disabled')
        self.log_text.pack(expand=True, fill=tk.BOTH)

        # Bring log window to front
        self.log_window.lift()
        self.log_window.attributes('-topmost', True)
        self.log_window.after_idle(self.log_window.attributes, '-topmost', False)

    def log(self, message):
        def append():
            self.log_text.config(state='normal')
            self.log_text.insert(END, message + '\n')
            self.log_text.see(END)
            self.log_text.config(state='disabled')
        self.log_text.after(0, append)

    def create_message(self, to_email, cc_email, subject, body):
        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = to_email
        msg['Cc'] = cc_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        return msg.as_string()

    def send_all_emails(self):
        if self.df is None:
            self.log("No CSV data loaded. Aborting sending.")
            return

        cc_email = self.cc_entry.get().strip()
        subject = self.subject_entry.get().strip()
        success_count = 0
        fail_rows = []

        try:
            smtp_conn = smtplib.SMTP(self.smtp_server, self.smtp_port)
            smtp_conn.ehlo()
            smtp_conn.starttls()
            smtp_conn.ehlo()

            # Authenticate with OAuth2 access token for SMTP
            access_token = self.creds.token
            auth_string = self.generate_oauth2_string(self.sender_email, access_token, base64_encode=True)

            smtp_conn.docmd('AUTH', 'XOAUTH2 ' + auth_string)
        except Exception as e:
            self.log(f"SMTP connection or authentication failed: {e}")
            messagebox.showerror("SMTP Auth Error", f"SMTP connection or authentication failed:\n{e}")
            return

        for idx, row in self.df.iterrows():
            try:
                mentor_email = row['MentorEmail']
                mentee_email = row['MenteeEmail']
                # Validate emails
                validate_email(mentor_email)
                validate_email(mentee_email)
                if cc_email:
                    validate_email(cc_email)

                body = self.body_template.format(**row)
                to_emails = [mentor_email, mentee_email]
                all_recipients = to_emails + ([cc_email] if cc_email else [])
                msg_str = self.create_message(','.join(to_emails), cc_email, subject, body)

                smtp_conn.sendmail(self.sender_email, all_recipients, msg_str)
                self.log(f"Email sent successfully to {mentor_email} and {mentee_email}")
                success_count += 1
            except EmailNotValidError as ev:
                self.log(f"Invalid email found for row {idx + 2} (CSV line): {ev}")
                fail_rows.append(row)
            except Exception as e:
                self.log(f"Failed to send email for row {idx + 2} (CSV line): {e}")
                fail_rows.append(row)

        smtp_conn.quit()

        self.log(f"\nFinished sending emails. Successfully sent: {success_count}. Failed: {len(fail_rows)}")

        if fail_rows:
            fail_df = pd.DataFrame(fail_rows)
            fail_csv_path = 'failed_emails.csv'
            fail_df.to_csv(fail_csv_path, index=False)
            self.log(f"Failed emails saved to {fail_csv_path}")

    def generate_oauth2_string(self, username, access_token, base64_encode=True):
        auth_string = f"user={username}\1auth=Bearer {access_token}\1\1"
        if base64_encode:
            auth_string = base64.b64encode(auth_string.encode()).decode()
        return auth_string


if __name__ == '__main__':
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
