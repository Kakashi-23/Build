from tkinter import *
from Meet import get_meet_id

master = Tk()
master.geometry('400x250')
subject = StringVar()
Label(master, text='Enter Subject').place(x=150, y=70)
subject_entry = Entry(master, textvariable=subject)
subject_entry.place(x=125, y=100)


def process(*args):
    id = subject_entry.get()
    subject_entry.delete(0, END)
    get_meet_id(id)


subject_entry.bind('<Return>', process)
mainloop()



