from datetime import datetime, timedelta
import schedule
from win32com.client import Dispatch
import tkinter
from tkinter import ttk
from operator import itemgetter


def send_alert_email(site_id, alarm_type):
    outlook = Dispatch('outlook.application')
    msg = outlook.CreateItem(0)
    msg.to = email_recipient.get()
    msg.Subject = site_id + ' ' + alarm_type
    msg.HTMLBody = '<html><body><p>Time is up!</p></body></html>'
    try:
        msg.Send()
    except:
        pass


def check_alarms():
    for item in watching:
        if item[1] <= datetime.now():
            send_alert_email(item[0], item[2])
            item[1] = item[1] + timedelta(minutes=15)
            display_watched()


def run_schedule():
    schedule.run_pending()
    root.after(1000, run_schedule)


def add_site(a):
    entry = []
    clinic = clinic_entry.get()
    time_down = time_entry.get()
    alarm_type = alarm_type_choice.get()
    if clinic == '':
        return
    if time_down == '' and (alarm_type != "Helmer Temps" and alarm_type != "Aruba Down"):
        print("Bad statement")
        return
    if time_down != '':
        time_down = int(time_down)
    clinic_choice.set('')
    time_down_choice.set('')
    clinic_entry.focus()
    if alarm_type == "One Hour":
        alarm = datetime.now() + timedelta(minutes=60 - time_down)
    elif alarm_type == "Three Hour":
        alarm = datetime.now() + timedelta(minutes=180 - time_down)
    elif alarm_type == "Helmer Temps":
        alarm = datetime.now() + timedelta(minutes=30)
    elif alarm_type == "Aruba Down":
        alarm = datetime.now() + timedelta(minutes=15)
    entry.append(clinic)
    entry.append(alarm)
    entry.append(alarm_type)
    watching.append(entry)
    display_watched()


def display_watched():
    global watching
    clinic_list.delete(0, 'end')
    watching = sorted(watching, key=itemgetter(1))
    for item in watching:
        clinic_list.insert('end', item[0] + '    ' + item[1].strftime('%H:%M') + '    ' + item[2])


def remove_alarm():
    try:
        selection = int(clinic_list.curselection()[0])
        del watching[selection]
        clinic_list.delete(selection)
    except IndexError:
        return


root = tkinter.Tk()
root.title("Clinic Down Acker 1.0")
root.geometry('390x260+200+200')

clinic_choice = tkinter.Variable(root)
alarm_type_choice = tkinter.Variable(root)
time_down_choice = tkinter.Variable(root)
email_recipient = tkinter.Variable(root)

watching = []
email_recipients = ['some-email@email.com', 'another-email@email.com', 'yet-another-email@email.com']
email_recipients = sorted(email_recipients)

entry_frame = tkinter.Frame(root)
entry_frame.grid(row=0, column=0)
tkinter.Label(entry_frame, text="Clinic:").grid(row=0, column=0)
clinic_entry = ttk.Entry(entry_frame, textvariable=clinic_choice, width=8)
clinic_entry.grid(row=0, column=1)
tkinter.Label(entry_frame, text="Minutes Down:").grid(row=0, column=2)
time_entry = ttk.Entry(entry_frame, textvariable=time_down_choice, width=8)
time_entry.grid(row=0, column=3)
alarm_type_choice.set("One Hour")
alarm_type = tkinter.OptionMenu(entry_frame, alarm_type_choice, "One Hour", "Three Hour", "Helmer Temps", "Aruba Down")
alarm_type.grid(row=0, column=4)
addButton = tkinter.Button(entry_frame, text="Add", command=lambda: add_site(watching), default='active')
addButton.grid(row=0, column=5)

list_frame = tkinter.Frame(root)
list_frame.grid(row=1, column=0)
clinic_list = tkinter.Listbox(list_frame, width=50)
clinic_list.grid(row=0, column=0)
delete_button = tkinter.Button(list_frame, text="Remove", command=remove_alarm)
delete_button.grid(row=2, column=0)

mail_frame = tkinter.Frame(root)
mail_frame.grid(row=2, column=0)
mail_recipient = tkinter.OptionMenu(mail_frame, email_recipient, *email_recipients)
mail_recipient.grid(row=0, column=0)

schedule.every(30).seconds.do(check_alarms)

root.after(1000, run_schedule)

clinic_entry.bind("<Return>", add_site)
time_entry.bind("<Return>", add_site)
clinic_entry.focus()
root.mainloop()
