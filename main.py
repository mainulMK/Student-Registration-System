from tkinter import *
from datetime import date
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import os
from tkinter.ttk import Combobox
import openpyxl,xlrd
from openpyxl import Workbook
import pathlib




root=Tk()
root.geometry("1250x700+210+100")
root.title('Student Registration System')
root.configure(bg="plum2")

file=pathlib.Path("Student Data.xlsx")
if file.exists():
    pass
else:
    file= Workbook()
    sheet = file.active
    sheet["A1"] = "Registration No:"
    sheet["B1"] = "Name"
    sheet["C1"] = "Class"
    sheet["D1"] = "Gender"
    sheet["E1"] = "DOB"
    sheet["F1"] = "Date of Registration"
    sheet["G1"] = "Religion"
    sheet["H1"] = "Skill"
    sheet["I1"] = "Father Name"
    sheet["J1"] = "Mother Name"
    sheet["K1"] = "Father's Occupation"
    sheet["L1"] = "Mother's Occupation"
    file.save("Student Data.xlsx")

# Upload image
def showimage():
    global filename
    global img
    filename=filedialog.askopenfilename(initialdir=os.getcwd(),
                                        title="Select image file", filetype=(("JPG File","*.jpg"),
                                                                             ("PNG File","*.png"),
                                                                              ("All Files","*.txt")))
    img =(Image.open(filename))
    resized_image = img.resize((190,190))
    photo2= ImageTk.PhotoImage(resized_image)
    lbl.config(image=photo2)
    lbl.image=photo2

# Save button
def save():
    R1=Registrarion.get()
    N1=Name.get()
    C1=Class.get()
    try:
        G1=gender
    except:
        messagebox.showerror("Rrror","select Gender")
    D2=DOB.get()
    D1=Date.get()
    Re1=Religion.get()
    S1=Skill.get()
    Fathername=F_Name.get()
    Mothername=M_Name.get()
    F1=Father_Occupation.get()
    M1=Mother_Occupation.get()

    if N1=="" or C1=="Select Class" or D2=="" or Re1=="" or S1=="" or Fathername=="" or Mothername=="" or F1=="" or M1=="":
        messagebox.showerror("Error","Few Data is Missing")
    else:
        file=openpyxl.load_workbook("Student Data.xlsx")
        sheet=file.active
        sheet.cell(column=1,row=sheet.max_row+1,value=R1)
        sheet.cell(column=2, row=sheet.max_row, value=N1)
        sheet.cell(column=3, row=sheet.max_row, value=C1)
        sheet.cell(column=4, row=sheet.max_row, value=G1)
        sheet.cell(column=5, row=sheet.max_row, value=D2)
        sheet.cell(column=6, row=sheet.max_row, value=D1)
        sheet.cell(column=7, row=sheet.max_row, value=Re1)
        sheet.cell(column=8, row=sheet.max_row, value=S1)
        sheet.cell(column=9, row=sheet.max_row, value=Fathername)
        sheet.cell(column=10, row=sheet.max_row, value=Mothername)
        sheet.cell(column=11, row=sheet.max_row, value=F1)
        sheet.cell(column=12, row=sheet.max_row, value=M1)

        file.save(r"Student Data.xlsx")

        try:
            img.save("Student Image/"+str(R1)+".jpg")
        except:
            messagebox.showinfo("info", "Profile picture is not available")
        messagebox.showinfo("info", "Successfully data entered")
        Clear()
        registration_no()


# Search Box

def search():
    text = Search.get()
    Clear()
    savebutton.config(state='disable')
    file=openpyxl.load_workbook("Student Data.xlsx")
    Sheet=file.active

    for row in Sheet.rows:
        if row[0].value == int(text):
            name = row[0]
            print(name)

            reg_no_position = str(name)[14:-1]
            reg_number= str(name)[15:-1]
            try:
                print(str(name))
            except:
                messagebox.showerror(("Ivalid","Invalid Registration Number"))
            x1 = Sheet.cell(row=int(reg_number), column=1).value
            x2 = Sheet.cell(row=int(reg_number), column=2).value
            x3 = Sheet.cell(row=int(reg_number), column=3).value
            x4 = Sheet.cell(row=int(reg_number), column=4).value
            x5 = Sheet.cell(row=int(reg_number), column=5).value
            x6 = Sheet.cell(row=int(reg_number), column=6).value
            x7 = Sheet.cell(row=int(reg_number), column=7).value
            x8 = Sheet.cell(row=int(reg_number), column=8).value
            x9 = Sheet.cell(row=int(reg_number), column=9).value
            x10 = Sheet.cell(row=int(reg_number), column=10).value
            x11 = Sheet.cell(row=int(reg_number), column=11).value
            x12 = Sheet.cell(row=int(reg_number), column=12).value

           # print(x1,x2, x3, x4,x5,x6,x7,x8,x9, x10,x11,x12)


            Registrarion.set(x1)
            Name.set(x2)
            Class.set(x3)

            if x4 =="Female":
                R2.select()
            else:
                R1.select()

            DOB.set(x5)
            Date.set(x6)
            Religion.set(x7)
            Skill.set(x8)
            F_Name.set(x9)
            M_Name.set(x10)
            Father_Occupation.set(x11)
            Mother_Occupation.set(x12)

            img= (Image.open("Student Image/"+str(x1)+".jpg"))
            resized_image = img.resize((190,190))
            photo2 = ImageTk.PhotoImage(resized_image)
            lbl.config(image=photo2)
            lbl.image=photo2


# Update Button config

def Update():
        R1 = Registrarion.get()
        N1 = Name.get()
        C1 = Class.get()
        selection()
        G1 = gender
        D2 = DOB.get()
        D1 = Date.get()
        Re1 = Religion.get()
        S1 = Skill.get()
        Fathername = F_Name.get()
        Mothername = M_Name.get()
        F1 = Father_Occupation.get()
        M1 = Mother_Occupation.get()
        file=openpyxl.load_workbook("Student Data.xlsx")
        Sheet=file.active


        for row in Sheet.rows:
            if row[0].value == R1:
                name=row[0]
                print(str(name))
                reg_no_position = str(name)[14:-1]
                reg_number=str(name)[15:-1]
                print(reg_number)

                #Sheet.cell(column=1,row=int(reg_number),value=R1)
                Sheet.cell(column=2, row=int(reg_number), value=N1)
                Sheet.cell(column=3, row=int(reg_number), value=C1)
                Sheet.cell(column=4, row=int(reg_number), value=G1)
                Sheet.cell(column=5, row=int(reg_number), value=D2)
                Sheet.cell(column=6, row=int(reg_number), value=D1)
                Sheet.cell(column=7, row=int(reg_number), value=Re1)
                Sheet.cell(column=8, row=int(reg_number), value=S1)
                Sheet.cell(column=9, row=int(reg_number), value=Fathername)
                Sheet.cell(column=10, row=int(reg_number), value=Mothername)
                Sheet.cell(column=12, row=int(reg_number), value=F1)
                Sheet.cell(column=12, row=int(reg_number), value=M1)
                file.save(r"Student Data.xlsx")

                try:
                    img.save("Student Image/"+str(R1)+".jpg")
                except:
                    pass
                messagebox.showinfo("Update", "Update Sucessfully")
                Clear()










# Reset/clear button
def Clear():
    global img
    Name.set("")
    DOB.set("")
    Religion.set("")
    Skill.set("")
    F_Name.set("")
    M_Name.set("")
    Father_Occupation.set("")
    Mother_Occupation.set("")
    Class.set("Select Class")

    registration_no()
    savebutton.config(state ='normal')
    img1=PhotoImage(file='Images/profile.png')
    lbl.config(image=img1)
    lbl.image=img1

    img=''



#Exit Window
def Exit():
    root.destroy()



#gender
def selection():
    global gender
    value=radio.get()
    if value==1:
        gender = "Male"
    else:
        gender = "Female"


####Registration No.
# it is crate to automatic enter registration no.

def registration_no():
    file=openpyxl.load_workbook("Student Data.xlsx")
    sheet=file.active
    row=sheet.max_row
    max_row_value=sheet.cell(row=row, column=1).value
    try:
        Registrarion.set(max_row_value+1)
    except:
        Registrarion.set("1")




# top frames
Label(root, text="Email: mdeasinarafatrana@gmail.com", width=10, height=2,bg="pink", anchor="e").pack(side=TOP, fill=X)
Label(root, text="STUDENT REGISTRATION", width=8, height=2,bg="magenta2", fg="white", font="aral 20 bold" ).pack(side=TOP, fill=X)

# Search box to update

Search = StringVar()
Entry(root, textvariable=Search, width=15, bd=2, font="arial 20").place(x=820, y=50)

imageicon3 =PhotoImage(file="Images/search.png")
srch = Button(root, text="Search",compound=LEFT, image = imageicon3,width=122,height=30, bg = "DarkOrchid1", font = "arial 13 bold", command=search)
srch.place(x=1060, y=50)


imageicon4=PhotoImage(file="Images/refresh.png")
redresh= Button(root,image=imageicon4, bg = "SlateBlue1",command =Update)
redresh.place(x=110, y=50)

# Registration and Date

Label(root,text= "Registration No :", font= "arial 15",fg = "white", bg = "blue").place(x=30, y=150)
Label(root,text= "Date :", font= "arial 15",fg = "white", bg = "blue").place(x=500, y=150)

Registrarion = IntVar()
Date = StringVar()


reg_entry=Entry(root,textvariabl=Registrarion, width=15, font="arial 15")
reg_entry.place(x=220, y=150)

registration_no()




today = date.today()
d1 = today.strftime("%d/%m/%Y")
date_entry = Entry(root, textvariable = Date, width =15, font = "arial 15")
date_entry.place(x=630, y=150)
Date.set(d1)



#student Details

obj = LabelFrame(root, text = "Student's Details", font = 20, bd = 2, width = 900, bg = "orchid1", fg = "blue",height= 250,relief = GROOVE)
obj.place(x=30, y=200)

Label(obj, text="Full Name     :",font = "arial 15", bg = "orchid2", fg="white").place(x=30, y=20)
Label(obj, text="Date of Birth :",font = "arial 15", bg = "orchid2", fg="white").place(x=30, y=92)
Label(obj, text="Gender         :",font = "arial 15", bg = "orchid2", fg="white").place(x=30, y=162)

Label(obj, text="Class     :",font = "arial 15", bg = "orchid2", fg="white").place(x=480, y=20)
Label(obj, text="Religion  :",font = "arial 15", bg = "orchid2", fg="white").place(x=480, y=92)
Label(obj, text="Skills      :",font = "arial 15", bg = "orchid2", fg="white").place(x=480, y=162)

Name=StringVar()
name_entry = Entry(obj, textvariable = Name, width =20, font = "arial 15")
name_entry.place(x=190, y=20)

DOB=StringVar()
dob_entry = Entry(obj, textvariable = DOB, width =20, font = "arial 15")
dob_entry.place(x=190, y=92)

radio = IntVar()
R1 =Radiobutton(obj, text = "Male", variable = radio, value =1, bg ="orchid2",fg = "white",command=selection)
R1.place(x=220, y=162)

R2 = Radiobutton(obj, text = "Female", variable = radio, value =2, bg ="orchid2",fg = "white",command=selection)
R2.place(x=330, y=162)

Class= Combobox(obj, values=['1','2','3','4','5','5','7','8','9','10','11','12'], font = "Roboto 10",width =28,state="r")
Class.place(x=610, y=20)
Class.set("Select Class")

Religion=StringVar()
religion_entry = Entry(obj, textvariable = Religion, width =20, font = "arial 15")
religion_entry.place(x=610, y=92)

Skill=StringVar()
skill_entry = Entry(obj, textvariable = Skill, width =20, font = "arial 15")
skill_entry.place(x=610, y=162)



#Parent's Details
#Father
obj2 = LabelFrame(root, text = "Parent's Details", font = 20, bd = 2, width = 900, bg = "orchid1", fg = "blue",height= 220,relief = GROOVE)
obj2.place(x=30, y=465)

Label(obj2, text="Father's Name :",font = "arial 15", bg = "orchid2", fg="white").place(x=30, y=40)
Label(obj2, text="Occupation     :",font = "arial 15", bg = "orchid2", fg="white").place(x=30, y=120)

F_Name=StringVar()
f_entry = Entry(obj2, textvariable = F_Name, width =20, font = "arial 15")
f_entry.place(x=190, y=40)

Father_Occupation=StringVar()
fo_entry = Entry(obj2, textvariable = Father_Occupation, width =20, font = "arial 15")
fo_entry.place(x=190, y=120)


# Mother

Label(obj2, text="Mother's Name :",font = "arial 15", bg = "orchid2", fg="white").place(x=480, y=40)
Label(obj2, text="Occupation      :",font = "arial 15", bg = "orchid2", fg="white").place(x=480, y=120)

M_Name=StringVar()
m_entry = Entry(obj2, textvariable = M_Name, width =20, font = "arial 15")
m_entry.place(x=645, y=40)

Mother_Occupation=StringVar()
mo_entry = Entry(obj2, textvariable = Mother_Occupation, width =20, font = "arial 15")
mo_entry.place(x=645, y=120)



# image

f=Frame(root, bd=3, bg="black",height=200,width=200,relief = GROOVE)
f.place(x=1000, y=150)

img=PhotoImage(file="Images/profile.png")
lbl=Label(f, bg="black",image=img)
lbl.place(x=0, y=0)

#button

Button(root, text = "Upload", width = 19, height = 2, font = "arial 12 bold", bg = "hot pink", command=showimage).place(x=1000, y=360)

savebutton=Button (root, text = "Save", width = 19, height = 2, font = "arial 12 bold", bg = "green2",command=save)
savebutton.place(x=1000, y=450)

Button(root, text = "Reset", width = 19, height = 2, font = "arial 12 bold", bg = "cyan",command=Clear).place(x=1000, y=540)

Button(root, text = "Exit", width = 19, height = 2, font = "arial 12 bold", bg = "red2",command=Exit).place(x=1000, y=630)

#Label(obj2, text="Mother's Name :",font = "arial 15", bg = "orchid2", fg="white").place(x=30, y=92)



root.mainloop()