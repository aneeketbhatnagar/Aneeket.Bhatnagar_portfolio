# Email_Sender_GUI_Outlook365_with_Month_Year_Fixed_Folder.py
# Updated: From address = CPA-IPSS InternalBilling <IPSSInternalBilling@cpaglobal.com>

import os
import win32com.client as win32
from tkinter import messagebox, ttk, filedialog
import tkinter as tk
from datetime import datetime

# ====================== CONFIG ======================
FROM_ADDRESS = "IPSSInternalBilling@cpaglobal.com"  # Yeh change kar sakta hai

# ====================== GUI ======================
root = tk.Tk()
root.title("Outlook 365 - Monthly Report Sender")
root.geometry("700x600")
root.configure(bg="#f0f4f8")

# Fonts
title_font = ("Arial", 16, "bold")
label_font = ("Arial", 12)
btn_font = ("Arial", 11, "bold")

# Global variables
selected_folder = ""
files_list = []
selected_month = tk.StringVar(value="January")
selected_year = tk.StringVar(value=str(datetime.now().year))

# ====================== LABELS ======================
folder_label = tk.Label(root, text="No folder selected", font=label_font, bg="#f0f4f8", fg="#7f8c8d")
status_label = tk.Label(root, text="Total Files: 0", font=label_font, bg="#f0f4f8", fg="#2c3e50")


def select_folder():
    global selected_folder, files_list
    try:
        folder = filedialog.askdirectory(
            title="Attachments Wala Folder Select Kar",
            initialdir=os.path.expanduser("~")
        )
        if folder:
            selected_folder = folder
            folder_label.config(text=f"Selected Folder: {folder}")
            refresh_files_list()
        else:
            folder_label.config(text="No folder selected")
    except Exception as e:
        messagebox.showerror("Folder Select Error",
                             f"Error: {str(e)}\nTry running as Administrator or check permissions.")


def refresh_files_list():
    global files_list
    file_listbox.delete(0, tk.END)
    files_list = []

    if selected_folder and os.path.exists(selected_folder):
        all_files = [f for f in os.listdir(selected_folder) if os.path.isfile(os.path.join(selected_folder, f))]
        for f in all_files:
            file_listbox.insert(tk.END, f)
            files_list.append(f)

    status_label.config(text=f"Total Files: {len(files_list)}")


def extract_email_from_filename(filename):
    name_part = filename.split("_")[0] if "_" in filename else filename.split(".")[0]
    if "@" in name_part:
        return name_part.strip()
    return None


def send_emails():
    if not selected_folder:
        messagebox.showwarning("Warning", "Please  select the folder!")
        return

    selected_indices = file_listbox.curselection()
    if not selected_indices:
        messagebox.showwarning("Warning", "No file select!")
        return

    selected_files = [files_list[i] for i in selected_indices]

    month = selected_month.get()
    year = selected_year.get()
    subject = f"Request for Billing Input Data - {month} {year}"

    progress_bar['value'] = 0
    progress_bar['maximum'] = len(selected_files)
    root.update()

    success_count = 0
    failed = []

    try:
        outlook = win32.Dispatch("outlook.application")
    except Exception as e:
        messagebox.showerror("Outlook Error",
                             f"Outlook connect nahi ho raha:\n{str(e)}\nOutlook app open rakh aur login kar.")
        return

    for idx, file in enumerate(selected_files):
        email = extract_email_from_filename(file)
        if email:
            try:
                mail = outlook.CreateItem(0)  # olMailItem
                mail.To = email
                mail.Subject = subject
                mail.Body = f"Dear User,\n\nThis is a system-generated notification to request the required billing input data for the upcoming billing cycle.\n\nKindly update and share the billing input file in the prescribed format with all mandatory fields completed, ensuring the data is accurate and submitted within the defined timeline to avoid any delays in billing processing.\n\nThank you for your cooperation.\n\nRegards,\nBilling Automation System\nClarivate"

                # IMPORTANT: From address set
                mail.Sender = outlook.Session.Accounts(FROM_ADDRESS)
                # Ya agar Sender kaam na kare to:
                # mail.SendUsingAccount = outlook.Session.Accounts(FROM_ADDRESS)

                file_path = os.path.join(selected_folder, file)
                mail.Attachments.Add(file_path)

                mail.Send()
                success_count += 1
            except Exception as e:
                failed.append(f"{file} → {str(e)}")
        else:
            failed.append(f"{file} → Email not found in filename")

        progress_bar['value'] = idx + 1
        progress_label.config(text=f"Progress: {idx + 1}/{len(selected_files)}")
        root.update()

    msg = f"Process complete!\nSuccess: {success_count}\nFailed: {len(failed)}\nSubject used: {subject}\nFrom: {FROM_ADDRESS}"
    if failed:
        msg += "\n\nFailed Files:\n" + "\n".join(failed)

    messagebox.showinfo("Result", msg)
    progress_bar['value'] = 0
    progress_label.config(text="")


# ====================== GUI LAYOUT ======================
title_label = tk.Label(root, text="Outlook 365 - Monthly Report Sender", font=title_font, bg="#f0f4f8", fg="#2c3e50")
title_label.pack(pady=20)

folder_btn = tk.Button(root, text="Select Attachments Folder", font=btn_font, bg="#3498db", fg="white",
                       command=select_folder, width=30, height=2)
folder_btn.pack(pady=10)

folder_label.pack(pady=5)

status_label.pack(pady=5)

month_year_frame = tk.Frame(root, bg="#f0f4f8")
month_year_frame.pack(pady=10)

tk.Label(month_year_frame, text="Select Month & Year for Subject:", font=label_font, bg="#f0f4f8").pack(side=tk.LEFT,
                                                                                                        padx=10)

month_combo = ttk.Combobox(month_year_frame, textvariable=selected_month,
                           values=["January", "February", "March", "April", "May", "June", "July", "August",
                                   "September", "October", "November", "December"], width=15)
month_combo.pack(side=tk.LEFT, padx=5)

year_combo = ttk.Combobox(month_year_frame, textvariable=selected_year, values=[str(y) for y in range(2020, 2030)],
                          width=10)
year_combo.pack(side=tk.LEFT, padx=5)

file_frame = tk.Frame(root, bg="#f0f4f8")
file_frame.pack(pady=10, fill="both", expand=True)

scrollbar = tk.Scrollbar(file_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

file_listbox = tk.Listbox(file_frame, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set, font=("Arial", 11),
                          height=10)
file_listbox.pack(side=tk.LEFT, fill="both", expand=True)

scrollbar.config(command=file_listbox.yview)

send_btn = tk.Button(root, text="Send Selected Files", font=btn_font, bg="#27ae60", fg="white",
                     command=send_emails, width=30, height=2)
send_btn.pack(pady=20)

progress_label = tk.Label(root, text="", font=label_font, bg="#f0f4f8", fg="#2c3e50")
progress_label.pack()

progress_bar = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
progress_bar.pack(pady=10)

root.mainloop()