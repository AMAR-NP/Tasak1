import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os

def enterdata():
    accepted=accept.get()
    if accepted=="Accepted":
            Firstname=first_name_entry.get()
            Lastname=Last_name_entry.get()
            if Firstname and Lastname:
                Title=title_combobox.get()
                Age=age_spinbox.get()
                Gender=gender_combobox.get()
                Occupation=occpu_entry.get()
                Disabled=reg.get()
                Email=info_check.get()
                Phone_no=number_entry.get()
                Alternative_no=anumber_entry.get()
                number_entry.get()
                Address=address.get("1.0",tkinter.END).strip()
                Landmark=land.get("1.0",tkinter.END).strip()
                Pincode=pinc.get("1.0",tkinter.END).strip()
                State=Statech.get()
                Nation=nationch.get()
            
                print("FirstName: ",Firstname,"LastName: ",Lastname)
                print("Title: ",Title,"Age: ",Age,"Disable: ",Disabled)
                print("Gender: ",Gender,"E-mail: ",Email,"Occupation: ",Occupation)
                print("Phone Number: ",Phone_no,"Alternative No: ",Alternative_no)
                print("Address: ",Address,"LandMark: ",Landmark)
                print("State: ",State,"Nation",Nation,"Pincode: ",Pincode)
                print("----------------------------------------------")
            
                filepath = "C:/Users/JOHNS/OneDrive/Desktop/assignment/Next24Tech/data.xlsx"

                if not os.path.exists(filepath):
                     workbook = openpyxl.Workbook()
                     sheet = workbook.active
                     heading = ["Title","First Name","Last Name", "Age","Gender","Occupation","Disability","E-mail","Phone","Alternative No","Address","Landmark","Pincode","State","Nation"]
                     sheet.append(heading)
                     workbook.save(filepath)
                workbook = openpyxl.load_workbook(filepath)
                sheet = workbook.active
                sheet.append([Title,Firstname,Lastname,Age,Gender,Occupation,Disabled,Email,Phone_no,Alternative_no,Address,Landmark,Pincode,State,Nation])
                workbook.save(filepath)
            
            else:
                 tkinter.messagebox.showwarning(title="ERROR",message="Name u idiot")    

    else:
         tkinter.messagebox.showwarning(title="ERROR",message="Fill u Idiot")

def validate_numeric_input(new_value):
    return new_value.isdigit() or new_value == ""

window = tkinter.Tk()
window.title("Registration Form")

frame = tkinter.Frame(window)
frame.pack()

user_info_frame= tkinter.LabelFrame(frame, text="USER DETAILS")
user_info_frame.grid(row=0, column=0,padx=20,pady=20)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0,column=1)
Last_name_label = tkinter.Label(user_info_frame, text="Last Name")
Last_name_label.grid(row=0,column=2)

first_name_entry = tkinter.Entry(user_info_frame,width=30)
Last_name_entry = tkinter.Entry(user_info_frame, width=30)
first_name_entry.grid(row=1,column=1)
Last_name_entry.grid(row=1,column=2)


title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["Mr.","Ms.","Dr.","Prof."," "], width=5)
title_label.grid(row=0,column=0)
title_combobox.grid(row=1,column=0)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=11, to=100,  width=5)
age_label.grid(row=2,column=0)
age_spinbox.grid(row=3,column=0)

gender_label = tkinter.Label(user_info_frame, text="Gender")
gender_combobox = ttk.Combobox(user_info_frame, values=["Male","Female","Trans","private"])
gender_label.grid(row=2,column=1)
gender_combobox.grid(row=3,column=1)

occpu = tkinter.Label(user_info_frame,text="Occupation")
occpu.grid(row=2,column=2)
occpu_entry= tkinter.Entry(user_info_frame,width=20)
occpu_entry.grid(row=3,column=2)

info = tkinter.Label(user_info_frame, text="Person With Disabilty")
reg=tkinter.StringVar(value="No")
info_checks = tkinter.Checkbutton(user_info_frame, text="Yes",variable=reg,onvalue="Yes",offvalue="No")

info.grid(row=4,column=0)
info_checks.grid(row=4,column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=2)

wid2 = tkinter.LabelFrame(frame, text="Contact Details")
wid2.grid(row=1,column=0,sticky="news",padx=20,pady=20)

info = tkinter.Label(wid2, text="Email-ID")
info_check = tkinter.Entry(wid2,width=30)
info.grid(row=0,column=0)
info_check.grid(row=0,column=1)

number = tkinter.Label(wid2,text="Phone No",anchor="e",justify="left")
number.grid(row=1,column=0)
vcmd = (window.register(validate_numeric_input), '%P') 
number_entry = tkinter.Entry(wid2,width=30,validate="key", validatecommand=vcmd)
number_entry.grid(row=1,column=1)

anumber = tkinter.Label(wid2,text="Aternative Phone No",anchor="w")
anumber.grid(row=2,column=0)
vcmd = (window.register(validate_numeric_input), '%P') 
anumber_entry = tkinter.Entry(wid2,width=30,validate="key", validatecommand=vcmd)
anumber_entry.grid(row=2,column=1)

for widget in wid2.winfo_children():
    widget.grid_configure(padx=10, pady=2)

wid3 = tkinter.LabelFrame(frame,text="Location")
wid3.grid(row=2,column=0,sticky="news",padx=20,pady=20)
add=tkinter.Label(wid3,text="Address",anchor="w",justify="left")
address = tkinter.Text(wid3,height=1,width=50)
add.grid(row=0,column=0)
address.grid(row=0,column=1)

landmark = tkinter.Label(wid3,text="Landmark")
landmark.grid(row=1,column=0)
land = tkinter.Text(wid3,height=1,width=50)
land.grid(row=1,column=1)

pin = tkinter.Label(wid3,text="Pincode")
pin.grid(row=2,column=0)
vcmd = (window.register(validate_numeric_input), '%P')
pinc =tkinter.Entry(wid3,width=20,validate="key", validatecommand=vcmd)
pinc.grid(row=2,column=1)

state = tkinter.Label(wid3,text="State")
Statech = ttk.Combobox(wid3,values=["Out Station","Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chhattisgarh",
    "Goa", "Gujarat", "Haryana", "Himachal Pradesh", "Jharkhand", "Karnataka",
    "Kerala", "Madhya Pradesh", "Maharashtra", "Manipur", "Meghalaya",
    "Mizoram", "Nagaland", "Odisha", "Punjab", "Rajasthan", "Sikkim", "Tamil Nadu",
    "Telangana", "Tripura", "Uttar Pradesh", "Uttarakhand", "West Bengal"],width=50)
state.grid(row=3,column=0)
Statech.grid(row=3,column=1)

nation = tkinter.Label(wid3,text="Nationality")
nationch = ttk.Combobox(wid3,values=[
    "Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda",
    "Argentina", "Armenia", "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain",
    "Bangladesh", "Barbados", "Belarus", "Belgium", "Belize", "Benin", "Bhutan", "Bolivia",
    "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei Darussalam", "Bulgaria",
    "Burkina Faso", "Burundi", "Cambodia", "Cameroon", "Canada", "Cape Verde",
    "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros",
    "Congo", "Costa Rica", "Côte d'Ivoire", "Croatia", "Cuba", "Cyprus", "Czech Republic",
    "Democratic People's Republic of Korea", "Democratic Republic of the Congo", "Denmark",
    "Djibouti", "Dominica", "Dominican Republic", "Ecuador", "Egypt", "El Salvador",
    "Equatorial Guinea", "Eritrea", "Estonia", "Ethiopia", "Fiji", "Finland", "France",
    "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Greece", "Grenada", "Guatemala",
    "Hungary", "Iceland", "India", "Indonesia", "Iran (Islamic Republic of)", "Iraq", "Ireland",
    "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kiribati", "Kuwait",
    "Kyrgyzstan", "Lao People's Democratic Republic", "Latvia", "Lebanon", "Lesotho", "Liberia", 
    "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Madagascar", "Malawi", "Malaysia", "Maldives", 
    "Mali", "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia (Federated States of)",
    "Monaco", "Mongolia", "Montenegro", "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal", "Netherlands",
    "New Zealand", "Nicaragua", "Niger", "Nigeria", "Niue", "North Korea", "Norway", "Oman", "Pakistan", "Palau"],width=50)
nation.grid(row=4,column=0)
nationch.grid(row=4,column=1)

for widget in wid3.winfo_children():
    widget.grid_configure(padx=10, pady=2)

wid4 = tkinter.LabelFrame(frame,text="Terms & Condition")
wid4.grid(row=3,column=0,sticky="news",padx=20,pady=20)

accept=tkinter.StringVar(value="Not Accepted")
ck=ttk.Checkbutton(wid4,text="I accept the T&C",variable=accept,onvalue="Accepted",offvalue="Not Accepted")
ck.grid(row=0,column=0)

btn= tkinter.Button(frame,text="Enter Data",command= enterdata)
btn.grid(row=4,column=0,sticky="news",padx=10,pady=10)

window.mainloop()