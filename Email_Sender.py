import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
from PIL import Image, ImageTk
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import csv
import time
import threading
import json

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Login")
        self.root.geometry("600x250")
        self.root.configure(bg="#FAFAF5")
        self.root.resizable(False, False)

        # Adding background image
        bg_image = Image.open("bi.png")
        bg_photo = ImageTk.PhotoImage(bg_image)
        bg_label = tk.Label(self.root, image=bg_photo)
        bg_label.image = bg_photo
        bg_label.place(relwidth=1, relheight=1)

        # Adding title image
        title_image = tk.PhotoImage(file="title2.png")
        title_label = tk.Label(self.root, image=title_image, bg="#FAFAF5")
        title_label.image = title_image
        title_label.grid(row=0, column=0, columnspan=3, padx=0, pady=(0,20), sticky="n")

        self.email_label = tk.Label(self.root, text="Email:", bg="#FAFAF5", font=("Arial bold", 12))
        self.email_label.grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.email_entry = tk.Entry(self.root, width=30, font=("Arial", 12))
        self.email_entry.grid(row=2, column=1, pady=5)

        self.password_label = tk.Label(self.root, text="Password:", bg="#FAFAF5", font=("Arial bold", 12))
        self.password_label.grid(row=3, column=0, padx=10, pady=(5,10), sticky="e")
        self.password_entry = tk.Entry(self.root, width=30, show="*", font=("Arial", 12))
        self.password_entry.grid(row=3, column=1, pady=(5,10))

        self.login_button = tk.Button(self.root, text=" LOGIN ", command=self.login, bg="#2196f3", fg="white", font=("Arial bold", 12))
        self.login_button.grid(row=4, column=1, padx=10, pady=10, sticky="e")

        self.login_button.bind("<Enter>", lambda event: self.change_button_color(self.login_button, "#044980"))
        self.login_button.bind("<Leave>", lambda event: self.change_button_color(self.login_button, "#2196f3"))

    def change_button_color(self, button, color):
        button.config(bg=color)

    def login(self):
        email = self.email_entry.get()
        password = self.password_entry.get()

        # Hardcoded credentials
        valid_email = "xyz@gmail.com" #Enter the Mail Id
        valid_password = "abcd efgh ijkl mnop" #App password of above mail Id

        if email == valid_email and password == valid_password:
            self.root.destroy()
            main(email)
        else:
            messagebox.showerror("Login Failed", "Invalid email or password")

class EmailSenderApp:
    def __init__(self, root, sender_email):
        self.root = root
        self.root.title("Email Sender")
        self.root.geometry("850x500")
        self.root.configure(bg="#FAFAF5")
        self.root.resizable(False, False)        

        # Adding background image
        bg_image = Image.open("bi.png")
        bg_photo = ImageTk.PhotoImage(bg_image)
        bg_label = tk.Label(self.root, image=bg_photo)
        bg_label.image = bg_photo
        bg_label.place(relwidth=1, relheight=1)

        self.file_paths = []
        self.sender_email = sender_email

        self.create_widgets()

    def create_widgets(self):
        title_image = tk.PhotoImage(file="title1.png")
        title_label = tk.Label(self.root, image=title_image, bg="#FAFAF5")
        title_label.image = title_image
        title_label.grid(row=0, column=0, columnspan=6, pady=0)

        self.email_icon = tk.PhotoImage(file="email_icon.png")
        self.root.iconphoto(False, self.email_icon)

        self.receiver_label = tk.Label(self.root, text="To:", bg="#FAFAF5", font=("Arial bold", 12))
        self.receiver_label.grid(row=1, column=0, padx=10, pady=(10, 5), sticky="e")
        self.receiver_entry = tk.Entry(self.root, width=60, font=("Arial", 12))
        self.receiver_entry.grid(row=1, column=1, pady=(10, 5), sticky="w")

        self.browse_csv_button = tk.Button(self.root, text="Browse CSV", command=self.browse_csv, bg="#2196f3", fg="white", font=("Arial bold", 12), relief=tk.FLAT)
        self.browse_csv_button.grid(row=1, column=2, padx=10, pady=(10, 5), sticky="w")

        self.subject_label = tk.Label(self.root, text="Subject:", bg="#FAFAF5", font=("Arial bold", 12))
        self.subject_label.grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.subject_entry = tk.Entry(self.root, width=60, font=("Arial", 12))
        self.subject_entry.grid(row=2, column=1, pady=5, sticky="w")

        self.message_label = tk.Label(self.root, text="Message:", bg="#FAFAF5", font=("Arial bold", 12))
        self.message_label.grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.message_text = tk.Text(self.root, width=60, height=6, font=("Arial", 12))
        self.message_text.grid(row=3, column=1, pady=5, sticky="w")
        self.message_scrollbar = tk.Scrollbar(self.root, command=self.message_text.yview)
        self.message_scrollbar.grid(row=3, column=2, sticky="ns")
        self.message_text.config(yscrollcommand=self.message_scrollbar.set)

        button_frame = tk.Frame(self.root, bg="#FAFAF5")
        button_frame.grid(row=4, column=1, columnspan=3, pady=(10, 5), sticky="w")

        self.bold_icon = tk.PhotoImage(file="bold.png").subsample(20)
        self.bold_button = tk.Button(button_frame, command=self.apply_bold, image=self.bold_icon, bg="#FAFAF5", bd=0, relief=tk.FLAT)
        self.bold_button.grid(row=0, column=2, padx=(0, 5), pady=5)

        self.italic_icon = tk.PhotoImage(file="italic-font.png").subsample(20)
        self.italic_button = tk.Button(button_frame, command=self.apply_italic, image=self.italic_icon, bg="#FAFAF5", bd=0, relief=tk.FLAT)
        self.italic_button.grid(row=0, column=3, padx=5, pady=5)

        self.underline_icon = tk.PhotoImage(file="underline.png").subsample(20)
        self.underline_button = tk.Button(button_frame, command=self.apply_underline, image=self.underline_icon, bg="#FAFAF5", bd=0, relief=tk.FLAT)
        self.underline_button.grid(row=0, column=4, padx=(5, 10), pady=5)

        self.emoji_icon = tk.PhotoImage(file="emoji.png").subsample(20)
        self.emoji_button = tk.Button(button_frame, command=self.insert_emoji, image=self.emoji_icon, bg="#FAFAF5", bd=0, relief=tk.FLAT)
        self.emoji_button.grid(row=0, column=5, padx=5, pady=5)

        self.attach_icon = tk.PhotoImage(file="attachment.png").subsample(20)
        self.attach_button = tk.Button(button_frame, command=self.attach_file, image=self.attach_icon, compound="center", bg="#FAFAF5", fg="white", font=("Arial", 12), relief=tk.FLAT)
        self.attach_button.grid(row=0, column=8, padx=10, pady=5)

        self.attachment_action_label = tk.Label(button_frame, text="Attach files", bg="#FAFAF5", font=("Arial", 10))
        self.attachment_action_label.grid(row=0, column=9, pady=5)

        self.menu_icon = tk.PhotoImage(file="menu.png").subsample(15)
        self.menu_button = tk.Button(button_frame, image=self.menu_icon, bg="#FAFAF5", bd=0, relief=tk.FLAT, command=self.menu_action)
        self.menu_button.grid(row=0, column=10, padx=15, pady=5)

        self.password_label = tk.Label(self.root, text="Password:", bg="#FAFAF5", font=("Arial bold", 12))
        self.password_label.grid(row=5, column=0, padx=10, pady=5, sticky="e")
        self.password_entry = tk.Entry(self.root, width=40, show="*", font=("Arial", 12))
        self.password_entry.grid(row=5, column=1, pady=5, sticky="w")

        self.send_icon = tk.PhotoImage(file="send_icon.png").subsample(15)
        self.send_button = tk.Button(self.root, command=self.send_email, text="Send", image=self.send_icon, compound="left", bg="#7fc400", fg="white", font=("Arial bold", 12), padx=10, relief=tk.FLAT)
        self.send_button.grid(row=6, column=1, padx=(10, 0), pady=10, sticky="e")

        self.schedule_icon = tk.PhotoImage(file="schedule.png").subsample(15)
        self.schedule_mail_button = tk.Button(self.root, command=self.schedule_mail, text="Schedule mail", image=self.schedule_icon, compound="left", bg="#ff9800", fg="white", font=("Arial bold", 12), relief=tk.FLAT)
        self.schedule_mail_button.grid(row=6, column=2, padx=(10, 0), pady=10, sticky="e")

        self.bold_button.bind("<Enter>", lambda event: self.change_button_color(self.bold_button, "#CCCCCC"))
        self.bold_button.bind("<Leave>", lambda event: self.change_button_color(self.bold_button, "#FAFAF5"))

        self.italic_button.bind("<Enter>", lambda event: self.change_button_color(self.italic_button, "#CCCCCC"))
        self.italic_button.bind("<Leave>", lambda event: self.change_button_color(self.italic_button, "#FAFAF5"))

        self.underline_button.bind("<Enter>", lambda event: self.change_button_color(self.underline_button, "#CCCCCC"))
        self.underline_button.bind("<Leave>", lambda event: self.change_button_color(self.underline_button, "#FAFAF5"))

        self.emoji_button.bind("<Enter>", lambda event: self.change_button_color(self.emoji_button, "#CCCCCC"))
        self.emoji_button.bind("<Leave>", lambda event: self.change_button_color(self.emoji_button, "#FAFAF5"))

        self.attach_button.bind("<Enter>", lambda event: self.change_button_color(self.attach_button, "#CCCCCC"))
        self.attach_button.bind("<Leave>", lambda event: self.change_button_color(self.attach_button, "#FAFAF5"))

        self.send_button.bind("<Enter>", lambda event: self.change_button_color(self.send_button, "#56a447"))
        self.send_button.bind("<Leave>", lambda event: self.change_button_color(self.send_button, "#7fc400"))

        self.schedule_mail_button.bind("<Enter>", lambda event: self.change_button_color(self.schedule_mail_button, "#b56d04"))
        self.schedule_mail_button.bind("<Leave>", lambda event: self.change_button_color(self.schedule_mail_button, "#ff9800"))


    def change_button_color(self, button, color):
        button.config(bg=color)

    def menu_action(self):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Save Draft", command=self.save_draft)
        menu.add_command(label="Load Draft", command=self.load_draft)

        # Display the submenu at the position of the menu button
        menu.post(self.menu_button.winfo_rootx(), self.menu_button.winfo_rooty() + self.menu_button.winfo_height())

    def save_draft(self):
        draft_data = {
            'receiver_emails': self.receiver_entry.get(),
            'subject': self.subject_entry.get(),
            'message_body': self.message_text.get("1.0", "end-1c"),
            'file_paths': self.file_paths
        }

        with open("draft.json", "w") as f:
            json.dump(draft_data, f)

    def load_draft(self):
        try:
            with open("draft.json", "r") as f:
                draft_data = json.load(f)

            self.receiver_entry.delete(0, tk.END)
            self.receiver_entry.insert(0, draft_data['receiver_emails'])

            self.subject_entry.delete(0, tk.END)
            self.subject_entry.insert(0, draft_data['subject'])

            self.message_text.delete("1.0", tk.END)
            self.message_text.insert("1.0", draft_data['message_body'])

            self.file_paths = draft_data['file_paths']

            if self.file_paths:
                file_names = "\n".join([path.split("/")[-1] for path in self.file_paths])
                self.attachment_action_label.config(text="Files Attached:\n" + file_names)
            else:
                self.attachment_action_label.config(text="No files attached")

        except FileNotFoundError:
            messagebox.showerror("Error", "No draft found.")

    def apply_bold(self):
        current_tags = self.message_text.tag_names("sel.first")
        if "bold" in current_tags:
            self.message_text.tag_remove("bold", "sel.first", "sel.last")
        else:
            self.message_text.tag_add("bold", "sel.first", "sel.last")
            self.message_text.tag_configure("bold", font=("Arial", 12, "bold"))

    def apply_italic(self):
        current_tags = self.message_text.tag_names("sel.first")
        if "italic" in current_tags:
            self.message_text.tag_remove("italic", "sel.first", "sel.last")
        else:
            self.message_text.tag_add("italic", "sel.first", "sel.last")
            self.message_text.tag_configure("italic", font=("Arial", 12, "italic"))

    def apply_underline(self):
        current_tags = self.message_text.tag_names("sel.first")
        if "underline" in current_tags:
            self.message_text.tag_remove("underline", "sel.first", "sel.last")
        else:
            self.message_text.tag_add("underline", "sel.first", "sel.last")
            self.message_text.tag_configure("underline", underline=True)

    def insert_emoji(self):
        pass

    def attach_file(self):
        self.file_paths = filedialog.askopenfilenames()
        if self.file_paths:
            file_names = "\n".join([path.split("/")[-1] for path in self.file_paths])
            self.attachment_action_label.config(text="Files Attached:\n" + file_names)
        else:
            self.attachment_action_label.config(text="No files attached")

    def browse_csv(self):
        csv_file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if csv_file_path:
            selected_emails = []
            with open(csv_file_path, newline='') as file:
                reader = csv.reader(file)
                for row in reader:
                    if row:
                        selected_emails.append(row[0])
            emails = ', '.join(selected_emails)
            self.receiver_entry.delete(0, tk.END)
            self.receiver_entry.insert(0, emails)

    def schedule_mail(self):
        scheduled_time = simpledialog.askstring("Schedule Mail", "Enter the time to schedule (format: HH:MM):")
        if scheduled_time:
            scheduled_time = scheduled_time.strip()
            if ":" not in scheduled_time or len(scheduled_time.split(":")) != 2:
                messagebox.showerror("Invalid Time", "Please enter the time in the correct format (HH:MM).")
                return

            def send_email_at_time():
                current_time = time.strftime("%H:%M")
                while current_time != scheduled_time:
                    time.sleep(1)
                    current_time = time.strftime("%H:%M")
                self.send_email()

            threading.Thread(target=send_email_at_time, daemon=True).start()

    def send_email(self):
        receiver_emails = self.receiver_entry.get().split(",")
        subject = self.subject_entry.get()
        message_body = self.message_text.get("1.0", "end-1c")
        password = self.password_entry.get()

        smtp_server = 'smtp.gmail.com'
        smtp_port = 587

        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = ", ".join(receiver_emails)
        msg['Subject'] = subject

        body = MIMEText(message_body)
        msg.attach(body)

        for file_path in self.file_paths:
            attachment = open(file_path, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {file_path}")
            msg.attach(part)

        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, password)
                server.sendmail(self.sender_email, receiver_emails, msg.as_string())
            messagebox.showinfo("Success", "Email sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

class MainWindow:
    def __init__(self, root, sender_email):
        self.root = root
        self.root.title("Main Window")
        self.root.geometry("600x350")
        self.root.configure(bg="#FAFAF5")
        self.root.resizable(False, False)

        # Adding background image
        bg_image = Image.open("bi.png")
        bg_photo = ImageTk.PhotoImage(bg_image)
        bg_label = tk.Label(self.root, image=bg_photo)
        bg_label.image = bg_photo
        bg_label.place(relwidth=1, relheight=1)

        # Adding title image
        title_image = tk.PhotoImage(file="title2.png")
        title_label = tk.Label(self.root, image=title_image, bg="#FAFAF5")
        title_label.image = title_image
        title_label.grid(row=0, column=0, columnspan=3, padx=0, pady=0, sticky="n")

        self.sender_email = sender_email

        button_frame = tk.Frame(self.root, bg="#FAFAF5")
        button_frame.grid(row=1, column=1, pady=10)

        self.send_single_button = tk.Button(button_frame, text="Send Mail to Single Person", command=self.send_single_mail, bg="#2196f3", fg="white", font=("Arial bold", 12), relief=tk.FLAT, compound=tk.TOP)
        self.send_single_button.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.send_single_button.icon = tk.PhotoImage(file="single.png").subsample(15)  # Load the icon image
        self.send_single_button.config(image=self.send_single_button.icon)  # Set the icon image to the button

        self.send_multiple_button = tk.Button(button_frame, text="Send Mail to Multiple People", command=self.send_multiple_mail, bg="#2196f3", fg="white", font=("Arial bold", 12), relief=tk.FLAT, compound=tk.TOP)
        self.send_multiple_button.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.send_multiple_button.icon = tk.PhotoImage(file="multiple.png").subsample(15)  # Load the icon image
        self.send_multiple_button.config(image=self.send_multiple_button.icon)  # Set the icon image to the button

        self.send_csv_button = tk.Button(button_frame, text="Send Mail to All Emails in CSV", command=self.send_csv_mail, bg="#2196f3", fg="white", font=("Arial bold", 12), relief=tk.FLAT, compound=tk.TOP)
        self.send_csv_button.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.send_csv_button.icon = tk.PhotoImage(file="csv.png").subsample(15)  # Load the icon image
        self.send_csv_button.config(image=self.send_csv_button.icon)  # Set the icon image to the button


        self.send_single_button.bind("<Enter>", lambda event: self.change_button_color(self.send_single_button, "#56aef5"))
        self.send_single_button.bind("<Leave>", lambda event: self.change_button_color(self.send_single_button, "#2196f3"))

        self.send_multiple_button.bind("<Enter>", lambda event: self.change_button_color(self.send_multiple_button, "#56aef5"))
        self.send_multiple_button.bind("<Leave>", lambda event: self.change_button_color(self.send_multiple_button, "#2196f3"))

        self.send_csv_button.bind("<Enter>", lambda event: self.change_button_color(self.send_csv_button, "#56aef5"))
        self.send_csv_button.bind("<Leave>", lambda event: self.change_button_color(self.send_csv_button, "#2196f3"))

    def change_button_color(self, button, color):
        button.config(bg=color)

    def send_single_mail(self):
        email_sender_window = tk.Toplevel(self.root)
        app = EmailSenderAppSingle(email_sender_window, self.sender_email)

    def send_multiple_mail(self):
        email_sender_window = tk.Toplevel(self.root)
        app = EmailSenderAppMultiple(email_sender_window, self.sender_email)

    def send_csv_mail(self):
        email_sender_window = tk.Toplevel(self.root)
        app = EmailSenderApp(email_sender_window, self.sender_email)

class EmailSenderAppSingle(EmailSenderApp):
    def create_widgets(self):
        super().create_widgets()
        # Remove the creation of the browse_csv_button
        self.browse_csv_button.destroy()
        pass

    def browse_csv(self):
        pass

    def send_email(self):
        receiver_emails = self.receiver_entry.get()
        # Check if comma is present in receiver_emails
        if ',' in receiver_emails:
            messagebox.showerror("Error", "You can only send to single recipient mail at once.")
            return
        subject = self.subject_entry.get()
        message_body = self.message_text.get("1.0", "end-1c")
        password = self.password_entry.get()

        smtp_server = 'smtp.gmail.com'
        smtp_port = 587

        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = receiver_emails
        msg['Subject'] = subject

        body = MIMEText(message_body)
        msg.attach(body)

        for file_path in self.file_paths:
            attachment = open(file_path, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {file_path}")
            msg.attach(part)

        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, password)
                server.sendmail(self.sender_email, receiver_emails.split(','), msg.as_string())
            messagebox.showinfo("Success", "Email sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


class EmailSenderAppMultiple(EmailSenderApp):
    def create_widgets(self):
        super().create_widgets()
        self.browse_csv_button.destroy()

    def browse_csv(self):
        pass

def main(sender_email=None):
    if sender_email:
        root = tk.Tk()
        app = MainWindow(root, sender_email)
        root.mainloop()
    else:
        root = tk.Tk()
        login_window = LoginWindow(root)
        root.mainloop()

if __name__ == "__main__":
    main()
