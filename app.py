import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from tkinter import scrolledtext
from tkinter import ttk
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import threading
from email_validator import validate_email, EmailNotValidError

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("O4U Match Email Sender")
        self.root.geometry("1200x800")

        self.csv_file_path = ''
        self.df = None
        self.current_page = 0
        self.rows_per_page = 25
        self.current_preview_index = 0  # Index for preview pagination

        self.smtp_server = "smtp.gmail.com"
        self.smtp_port = 587

        self.body_template = """Hey {MentorFirstName} and {MenteeFistName}!\n\nIn preparation for program kickoff, we are so excited to formally introduce you to your Summer 2025 O4U mentorship program pairing!\n\nMentor: {MentorFirstName} {MentorLastName}, {JobTitle} at {PlaceOfWork}\nMentee: {MenteeFistName} {MenteeLastName}, studying {Major} at {University}\n\nPlease feel free to use this email thread to introduce yourselves to each other before kickoff on Wednesday, [Put Date Here] at 7:30pm EST -- and get your first meeting in the books! Remember that we expect mentors and mentees to meet twice a month. \n\nYou will also be receiving the training materials, including the discussion guide and the recording of last week's sessions shortly. Please review if you were not able to attend the trainings. Please also keep these dates blocked on your calendar:\n\nMidway Event: Wednesday,[Put Date Here] @ 7:30pm ET\nClosing Event: Wednesday, [Put Date Here] @ 7:30pm ET\n\nIf you don't hear back from your mentor / mentee within a week, feel free to reach out to mentoring@outforundergrad.org. \n\nBest,\nJared"""
        self.email_subject = "Summer 2025 O4U Mentorship Program Pairing"

        # Title Label
        self.title_label = tk.Label(root, text="O4U Match Email Sender", font=("Helvetica", 16, "bold"))
        self.title_label.pack(pady=10)

        # Upload Button
        self.upload_btn = tk.Button(root, text="Upload CSV", command=self.upload_csv)
        self.upload_btn.pack(pady=10)

        # Email Subject
        self.subject_label = tk.Label(root, text="Email Subject:")
        self.subject_label.pack()
        self.subject_entry = tk.Entry(root, width=80)
        self.subject_entry.insert(0, self.email_subject)
        self.subject_entry.pack(pady=10)

        # CC Entry
        self.cc_label = tk.Label(root, text="CC Email:")
        self.cc_label.pack()
        self.cc_entry = tk.Entry(root, width=80)
        self.cc_entry.insert(0, 'mentoring@outforundergrad.org')
        self.cc_entry.pack(pady=10)

        # Treeview and Scrollbars
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, expand=True, fill=tk.BOTH)

        self.tree_scroll_y = tk.Scrollbar(self.tree_frame)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_scroll_x = tk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = tk.ttk.Treeview(self.tree_frame, columns=[], show='headings', yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(expand=True, fill=tk.BOTH)

        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        # Pagination Controls
        self.pagination_frame = tk.Frame(root)
        self.pagination_frame.pack(pady=10)

        self.prev_btn = tk.Button(self.pagination_frame, text="<< Prev", command=self.show_prev_page)
        self.prev_btn.grid(row=0, column=0, padx=5)

        self.page_label = tk.Label(self.pagination_frame, text="Page 0")
        self.page_label.grid(row=0, column=1, padx=5)

        self.next_btn = tk.Button(self.pagination_frame, text="Next >>", command=self.show_next_page)
        self.next_btn.grid(row=0, column=2, padx=5)

        # Buttons to Edit Body, Preview, and Send
        self.edit_body_btn = tk.Button(root, text="Edit Email Body", command=self.edit_email_body, state=tk.DISABLED)
        self.edit_body_btn.pack(pady=10)

        self.preview_btn = tk.Button(root, text="Preview Email", command=self.preview_email, state=tk.DISABLED)
        self.preview_btn.pack(pady=10)

        self.send_btn = tk.Button(root, text="Send Emails", command=self.prompt_for_credentials, state=tk.DISABLED)
        self.send_btn.pack(pady=10)

    def prompt_for_credentials(self):
        self.sender_email = simpledialog.askstring("Input", "Enter your email address:", parent=self.root)
        self.sender_password = simpledialog.askstring("Input", "Enter your email password:", parent=self.root, show='*')
        if self.sender_email and self.sender_password:
            # Create a popup for real-time log
            self.log_popup = tk.Toplevel(self.root)
            self.log_popup.title("Email Sending Log")

            self.log_text = scrolledtext.ScrolledText(self.log_popup, width=100, height=30)
            self.log_text.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)

            self.log_text.insert(tk.END, "Starting email sending...\n")
            self.log_text.yview(tk.END)

            threading.Thread(target=self.send_emails).start()
        else:
            messagebox.showerror("Error", "Email address and/or password not provided.")

    def upload_csv(self):
        self.csv_file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not self.csv_file_path:
            return
        try:
            self.df = pd.read_csv(self.csv_file_path)
            required_columns = ['MentorEmail', 'MenteeEmail', 'MentorFirstName', 'MentorLastName', 'MenteeFistName', 'MenteeLastName', 'JobTitle', 'PlaceOfWork', 'Major', 'University']
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

            # Disable/enable pagination buttons
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
            messagebox.showerror("Error", "No CSV file uploaded.")
            return

        # Create a preview window
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Email Preview")

        # Display the email body
        preview_text = scrolledtext.ScrolledText(preview_window, width=100, height=30)
        preview_text.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)

        # Show a preview of the first row
        self.current_preview_index = 0
        self.show_email_preview(preview_text, self.current_preview_index)

        # Pagination Controls
        preview_controls = tk.Frame(preview_window)
        preview_controls.pack(pady=10)

        prev_btn = tk.Button(preview_controls, text="<< Prev", command=lambda: self.show_email_preview(preview_text, max(self.current_preview_index - 1, 0)))
        prev_btn.pack(side=tk.LEFT, padx=5)

        next_btn = tk.Button(preview_controls, text="Next >>", command=lambda: self.show_email_preview(preview_text, min(self.current_preview_index + 1, len(self.df) - 1)))
        next_btn.pack(side=tk.LEFT, padx=5)

    def show_email_preview(self, text_widget, index):
        if self.df is not None and 0 <= index < len(self.df):
            row = self.df.iloc[index]
            email_body = self.body_template.format(
                MentorFirstName=row['MentorFirstName'],
                MentorLastName=row['MentorLastName'],
                MenteeFistName=row['MenteeFistName'],
                MenteeLastName=row['MenteeLastName'],
                JobTitle=row['JobTitle'],
                PlaceOfWork=row['PlaceOfWork'],
                Major=row['Major'],
                University=row['University']
            )
            text_widget.delete("1.0", tk.END)
            text_widget.insert(tk.END, email_body)
            self.current_preview_index = index

    def send_emails(self):
        if self.df is None:
            messagebox.showerror("Error", "No CSV file uploaded.")
            return

        log_file = "email_send_log.txt"
        failed_csv_file = "failed_emails.csv"

        failed_rows = []

        for index, row in self.df.iterrows():
            mentor_email = row['MentorEmail']
            mentee_email = row['MenteeEmail']

            # Validate emails
            try:
                validate_email(mentor_email)
                validate_email(mentee_email)
            except EmailNotValidError:
                message = f"Invalid email addresses for row {index + 1}. Skipping email send.\n"
                self.log_text.insert(tk.END, message)
                self.log_text.yview(tk.END)
                failed_rows.append(row)
                continue

            # Prepare email
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = f"{mentor_email}, {mentee_email}"
            msg['Cc'] = self.cc_entry.get()  # CC field from the GUI
            msg['Subject'] = self.subject_entry.get()

            body = self.body_template.format(
                MentorFirstName=row['MentorFirstName'],
                MentorLastName=row['MentorLastName'],
                MenteeFistName=row['MenteeFistName'],
                MenteeLastName=row['MenteeLastName'],
                JobTitle=row['JobTitle'],
                PlaceOfWork=row['PlaceOfWork'],
                Major=row['Major'],
                University=row['University']
            )

            msg.attach(MIMEText(body, 'plain'))

            # Send email
            try:
                with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                    server.starttls()
                    server.login(self.sender_email, self.sender_password)
                    server.sendmail(self.sender_email, [mentor_email, mentee_email, self.cc_entry.get()], msg.as_string())
                    log_message = f"Email sent successfully to: {mentor_email}, {mentee_email}\n"
                    self.log_text.insert(tk.END, log_message)
                    self.log_text.yview(tk.END)
                    print(log_message)
            except Exception as e:
                log_message = f"Failed to send email to: {mentor_email}, {mentee_email}. Error: {e}\n"
                self.log_text.insert(tk.END, log_message)
                self.log_text.yview(tk.END)
                print(log_message)
                failed_rows.append(row)

        # Save failed rows to a CSV file
        if failed_rows:
            pd.DataFrame(failed_rows).to_csv(failed_csv_file, index=False)
            messagebox.showinfo("Completed", f"Emails processed. Failed emails saved to: {failed_csv_file}")

        messagebox.showinfo("Completed", f"Email sending completed. Log displayed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()

