import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl
from PIL import ImageTk, Image
import pathlib


def enter_data():
    accepted = accept_var.get()
    
    if accepted=="Accepted":
        # User info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        
        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            state = state_combobox.get()
            
            #ClientInfo
            Aadharcard = Aadharcard_entry.get()
            Lightbill = Lightbill_entry.get()
            # Plan Info
            Cap= Cap_combobox.get()
            Price= price_withsub_combobox.get()
            
            
            print("First name: ", firstname, "Last name: ", lastname)
            print("Title: ", title, "Age: ", age, "State: ", state)
            print("AadharCard No.:",Aadharcard , "Lightbill Link: ", Lightbill)
            print("Plant Capacity:", Cap)
            print("System Price with Subsidy:",Price)
            print("------------------------------------------")
            filepath = r"C:\Users\AAYUSH\Desktop\Data.xlsx"
            
            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet['A1']="First Name"
                sheet['B1']="Last Name"
                sheet['C1']="Title"
                sheet['D1']="Age"
                sheet['E1']="State"
                sheet['F1']="AadharCard No."
                sheet['G1']="Lightbill Link"
                sheet['H1']="Plant Capacity"
                sheet['I1']="System Price with Subsidy"
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title, age, state, Aadharcard,
                          Lightbill, Cap, Price])
            workbook.save(filepath)   
            
        else:
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:
        tkinter.messagebox.showwarning(title= "Error", message="You have not accepted the terms")

window = tkinter.Tk()
window.title("NETCCO")



lbl_title=tkinter.Label(text="NETCCO",font=("times new roman",37,"bold"),fg="blue",bg="white")
lbl_title.place(x=0,y=0,width=300,height=80)


frame = tkinter.Frame(window)
frame.pack()
frame.place(x=150,y=155,width=700,height=500)

# Saving User Info
user_info_frame =tkinter.LabelFrame(frame, text="Client information")
user_info_frame.grid(row= 1, column=0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=1, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=1, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=2, column=0)
last_name_entry.grid(row=2, column=1)

title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms.", "Dr."])
title_label.grid(row=1, column=2)
title_combobox.grid(row=2, column=2)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=3, column=0)
age_spinbox.grid(row=4, column=0)

state_label = tkinter.Label(user_info_frame, text="State")
state_combobox = ttk.Combobox(user_info_frame, values=['Andhra Pradesh', 'Arunachal Pradesh', 'Assam', 'Bihar', 'Chhattisgarh', 'Goa', 'Gujarat', 'Haryana', 'Himachal Pradesh', 'Jharkhand', 'Karnataka', 'Kerala', 'Madhya Pradesh', 'Maharashtra', 'Manipur', 'Meghalaya', 'Mizoram', 'Nagaland', 'Odisha', 'Punjab', 'Rajasthan', 'Sikkim', 'Tamil Nadu', 'Telangana', 'Tripura', 'Uttar Pradesh', 'Uttarakhand', 'West Bengal'])
state_label.grid(row=3, column=1)
state_combobox.grid(row=4, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# ClientInfo
registered_frame = tkinter.LabelFrame(frame, text="Documents Required")
registered_frame.grid(row=4, column=0, sticky='news', padx=20, pady=10)

Aadharcard_label = tkinter.Label(registered_frame, text= "Aadhar Card No.")
Aadharcard_entry= tkinter.Entry(registered_frame)
Aadharcard_label.grid(row=4, column=0)
Aadharcard_entry.grid(row=5, column=0)

Lightbill_label = tkinter.Label(registered_frame, text="Light Bill(Enter Google Drive Link)")
Lightbill_entry = tkinter.Entry(registered_frame)
Lightbill_label.grid(row=4, column=2)
Lightbill_entry.grid(row=5, column=2)    

for widget in registered_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)
    
#Plan Info
plan_frame = tkinter.LabelFrame(frame, text='Plan Info')
plan_frame.grid(row=5, column=0, sticky='news', padx=20, pady=10)

Cap_label= tkinter.Label(plan_frame, text='Choose Plan No.')
Cap_combobox = ttk.Combobox(plan_frame, values=['1kW','2kW','3kW','4kW','5kW','10kW'])
Cap_label.grid(row=6,column=0)
Cap_combobox.grid(row=5,column=0)

price_withsub_label= tkinter.Label(plan_frame, text='Choose Respective Price with Subsidy')
price_withsub_combobox = ttk.Combobox(plan_frame, values=['₹93,200','₹1,46,400','₹1,89,600','₹2,63,600','₹3,35,300','₹6,94,700'])
price_withsub_label.grid(row=6,column=1)
price_withsub_combobox.grid(row=5,column=1)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=7, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text= "I accept the terms and conditions.",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=7, column=0)

# Button
button = tkinter.Button(frame, text="Enter data", command= enter_data)
button.grid(row=8, column=0, sticky="news", padx=20, pady=10)

#close window
def quit_tk():
    window.quit()

close=tkinter.Button(frame, text="Close Window",command=quit_tk)
close.grid(column=0,row=10, sticky="news",padx=20,pady=10)

window.mainloop()


