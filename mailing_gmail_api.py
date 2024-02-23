import base64
import email
from tkinter import *
import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from datetime import datetime
import re
from tkinter import filedialog, messagebox
import customtkinter

customtkinter.set_appearance_mode("system")
# customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

csv_file = None
columns = [" "]
df = pd.DataFrame({})
draft_id = []
sheetname_ = [" "]


def build_scope():
    global service
    scopes = ['https://mail.google.com/']
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build("gmail", "v1", credentials=creds)


def check_email(choice, update_=False):
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    n = 0
    for i, row in df.iterrows():
        email_ = str(row[choice])
        if re.fullmatch(regex, email_):
            n += 1
    if n == 0:
        font_color = 'red'
    else:
        font_color = 'green'
    output_csv.configure(text=f"Valid email found : {n}", text_color=font_color)
    if update_:
        window.update()


def check_email_before_send(email_):
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    if re.fullmatch(regex, email_):
        return True


def refresh():
    global df, sheetname_
    if label_csv.get():
        filename, file_extension = os.path.splitext(label_csv.get())

        if file_extension == '.csv':
            df = pd.read_csv(label_csv.get())
            combo_sn.configure(state='disabled')
            combo_sn.set(" ")

        elif file_extension == '.xlsx':
            combo_sn.configure(state='normal')
            sheetname_ = pd.ExcelFile(label_csv.get()).sheet_names
            combo_sn.configure(values=sheetname_)
            combo_sn.set(sheetname_[0])

            df = pd.read_excel(label_csv.get())
        else:
            messagebox.showwarning(title="Unvalid file", message="Select CSV file or Excel file only")
        if file_extension == '.csv' or file_extension == '.xlsx':
            columns_ = df.columns
            combo_col.configure(values=columns_)
            combo_col.set(columns_[0])
    if combo_col.get() in df.columns:
        select_col_with_mail()


def select_col_with_mail():
    for col in df.columns:
        check_email(col)
        if output_csv.cget("text") != "Valid email found : 0":
            combo_col.set(col)
            check_email(col, update_=True)
            break


def update_excel(sheet):
    global df
    df = pd.read_excel(label_csv.get(), sheet_name=sheet)
    columns_ = df.columns
    combo_col.configure(values=columns_)
    # combo_col.set(columns_[0])
    # check_email(combo_col.get())
    select_col_with_mail()


def select_csv_():
    window.filename = filedialog.askopenfilename(initialdir="/",
                                                 title="Select excel file or csv file",
                                                 filetypes=(("CSV files", "*.csv"),
                                                            ("Excel files", "*.xlsx"),
                                                            ("All files", "*.*")))
    return window.filename


def select_csv():
    global csv_file
    csv_file = select_csv_()
    label_csv.delete(0, 'end')
    label_csv.insert(0, csv_file)
    refresh()


def get_draft_id(x):
    global draft_id
    draft_id = x[0]
    draft_window.destroy()
    print(x[1].cget("text"))
    label_draft.configure(text=(x[1].cget("text")))


def select_draft_frame():
    global draft_window
    draft_window = customtkinter.CTkToplevel(window)
    draft_window.title("Draft Selection")
    draft_window.geometry('600x300')
    draft_window.resizable(width=False, height=False)
    draft_window.configure(padx=1, pady=1)
    draft_window.transient()  # place this window on top of the root window
    draft_window.grab_set()
    date_=datetime.now()
    frame_d = customtkinter.CTkScrollableFrame(draft_window, width=598, height=300,
                                               label_text="Select a draft from gmail")
    frame_d.pack()
    drafts = service.users().drafts().list(userId='me').execute()

    for draf in drafts['drafts']:
        draft = service.users().drafts().get(userId='me', id=draf['id']).execute()
        subject = 'Not Defined'
        for dic in draft['message']['payload']['headers']:
            if dic.get('name') == 'Subject':
                subject = dic.get('value')
            if dic.get('name') == 'Date':
                date_ = dic.get('value')

        if subject == '':
            subject = 'Not Defined'
        n = f"{subject} / {date_[:-5]}"
        btn = customtkinter.CTkButton(frame_d, text=n,  width=590, )
        btn.pack()
        btn.configure(command=lambda x=(draf['id'], btn): (get_draft_id(x)))
    draft_window.lift()


def send_email():
    recipients = []
    if not csv_file:
        message = "Select valid excel or csv file"
        print(message)
        messagebox.showwarning(title="Unable to send", message=message)
        return
    if output_csv.cget("text") == "Valid email found : 0":
        message="Select valid column with Email Address"
        print(message)
        messagebox.showwarning(title="Unable to send", message=message)
        return
    if not draft_id:
        message = "Select valid draft"
        print(message)
        messagebox.showwarning(title="Unable to send", message=message)
        return
    status.grid()
    window.update()
    subject_ = label_draft.cget("text")
    # subject_ = label_draft.cget("text").split("/")[0]
    for i, row in df.iterrows():
        email_add = row[combo_col.get()]
        if check_email_before_send(email_add):
            if email_add in recipients:
                if add_var.get() == 1:
                    df.loc[i, subject_] = f"Duplicate : sent"
            else:
                draft = service.users().drafts().get(userId='me', id=draft_id, format='raw').execute()
                raw = draft['message']['raw']

                msg_email = email.message_from_bytes(base64.urlsafe_b64decode(raw))
                del msg_email['To']
                msg_email.add_header('To', email_add)
                encoded_message = base64.urlsafe_b64encode(msg_email.as_bytes()).decode()
                create_message = {"message": {"raw": encoded_message}}
                draft_ = (
                    service.users()
                    .drafts()
                    .create(userId="me", body=create_message)
                    .execute())

                service.users().drafts().send(userId="me", body={'id': draft_['id']}).execute()
                recipients.append(email_add)
                if add_var.get() == 1:
                    df.loc[i,subject_] = f"sent - {datetime.now()}".split('.')[0]
                print(f"{i+1}) {datetime.now()} {email_add}")
                status.set((i+1)/len(df))
                window.update()
    status.set(0)
    status.grid_remove()

    if add_var.get() == 1:
        filename, file_extension = os.path.splitext(label_csv.get())

        if file_extension == '.csv':
            df.to_csv(csv_file, index=False)

        elif file_extension == '.xlsx':
            with pd.ExcelWriter(csv_file, mode='a',
                                if_sheet_exists='replace') as x0:
                df.to_excel(x0, sheet_name=combo_sn.get(), index=False)



def quit_window():
    window.destroy()


window = customtkinter.CTk()

window.title("Email Manager")
# window.iconbitmap("youtube_icon.ico")
window.geometry('800x230')
window.resizable(width=False, height=False)
window.configure(padx=1, pady=1)

build_scope()

frame1 = customtkinter.CTkFrame(window, width=799, height=150)
frame1.pack_propagate(False)
frame1.pack()

label_csv0 = customtkinter.CTkLabel(frame1, text="EMAIL SELECTION FROM SPREADSHEET", text_color=('#062557','#ffffff'),
                                    font=("Arial", 15, 'bold', 'italic'), height=10, width=795, justify='center')
label_csv0.grid(row=0, column=0, sticky=EW, columnspan=3)

button_csv = customtkinter.CTkButton(frame1, text="Select file", command=select_csv, width=150,)
button_csv.grid(row=1, column=0, sticky=EW)


label_csv = customtkinter.CTkEntry(frame1, placeholder_text="Select Excel or csv file", font=("Arial", 15),
                                   height=10, width=630, )
label_csv.bind('<FocusOut>', lambda x: refresh())
label_csv.grid(row=1, column=1, sticky=EW, columnspan=2)


label_sn = customtkinter.CTkLabel(frame1, text="Select a sheet", text_color='grey', justify='right', padx=5,
                                  anchor=E, font=("Arial", 12, "italic"), height=10, width=150)
label_sn.grid(row=2, column=0, sticky=EW,)

combo_sn = customtkinter.CTkOptionMenu(frame1, values=sheetname_, command=update_excel, state='disabled')
combo_sn.grid(row=2, column=1, sticky=EW)

combo_col = customtkinter.CTkOptionMenu(frame1, values=columns, command=check_email)
combo_col.grid(row=3, column=1, sticky=EW)

label_csv0 = customtkinter.CTkLabel(frame1, text="Select a column", text_color='grey',  justify='right', padx=5,
                                    anchor=E, font=("Arial", 12, "italic"), height=10, width=150)
label_csv0.grid(row=3, column=0, sticky=EW,)

output_csv = customtkinter.CTkLabel(frame1, text="",   anchor=E, font=("Arial", 11), height=10,)
output_csv.grid(row=4, column=0, sticky=EW, )

frame2 = customtkinter.CTkFrame(window, width=799, height=150)
frame2.pack_propagate(False)
frame2.pack()

label_csv0 = customtkinter.CTkLabel(frame2, text="DRAFT SELECTION FROM GMAIL", text_color=('#062557','#ffffff'),
                                    font=("Arial", 15, 'bold', 'italic'), height=10, width=795)
label_csv0.grid(row=0, column=0, sticky=EW, columnspan=2)

button_draft= customtkinter.CTkButton(frame2, text="Select draft", command=select_draft_frame, width=150,)
button_draft.grid(row=1, column=0, sticky=EW)

label_draft = customtkinter.CTkLabel(frame2, text="",   anchor=W, font=("Arial", 15),
                                     height=10, width=630)
label_draft .grid(row=1, column=1, sticky=EW)


add_var = customtkinter.IntVar(value=0)
add_check = customtkinter.CTkCheckBox(frame2, text="Add a status in the spreadsheet", variable=add_var,
                                      onvalue=1, offvalue=0)
add_check.grid(row=3, column=0, sticky=EW, columnspan=2)

status = customtkinter.CTkProgressBar(frame2, height=10, width=780, corner_radius=2, progress_color="#9CB0EB")
status.grid(row=4, column=0, sticky=EW, columnspan=2)
status.set(0)
status.grid_remove()

button_send = customtkinter.CTkButton(window, text="Send Email", command=send_email, width=150,)
button_send.pack(side=LEFT)
#
button_quit = customtkinter.CTkButton(window, text="Quit", command=quit_window,  width=150, )
button_quit.pack(side=RIGHT)
#
window.mainloop()