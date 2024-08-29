from tkinter import *
from tkinter import ttk
import openpyxl,os,random,search
import tkinter.messagebox as tmsg
from tkcalendar import DateEntry




def printf():
         
        
    def submit():
        if name.get()=='' and email.get()=='' and phone.get()=='':
             tmsg.showerror("Incomplete","Please fill all necessary entries")
        else:
            department = deptVar.get()+".xlsx"
            # filepath = "C:\\Users\\ratul\\Desktop\\python gui\\pythonproject 2\\" + department
            filepath = department
            if not os.path.exists(filepath):
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    heading = ['NAME','UID','NATIONALITY','GENDER','DOB','EMAIL','PHONE','FATHER NAME','MOTHER NAME','G .PHONE','STATE','P.O','P.S','PIN','HS SCHOOL','S SCHOOL','12TH','10TH']
                    sheet.append(heading)
                    workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            appNo = random.randint(1000,2000)
            sheet.append([name.get(),appNo,religion.get(),gender.get(),dob.get(),email.get(),phone.get(),fatherName.get(),motherName.get(),Pphone.get(),dist.get(),po.get(),ps.get(),pin.get(),hs.get(),sc.get(),hsmrk.get(),smrk.get()])
            workbook.save(filepath)
            tmsg.showinfo('Welcome',f'Application submitted Successfully.\nApplication no is {appNo}\nDO NOT SHARE IT WITH ANYONE')
    def clear():
        deptVar.set('Select Department')
        name.set('')
        gender.set('')
        religion.set('')
        fatherName.set('')
        motherName.set('')
        Pphone.set('')
        pin.set('')
        ps.set('')
        po.set('')
        dist.set('')
        hs.set('')
        sc.set('')
        hsmrk.set('')
        smrk.set('')
        email.set('')
        phone.set('')
        
       
        
    name = StringVar()
    gender = StringVar()
    religion = StringVar()
    fatherName = StringVar()
    motherName = StringVar()
    Pphone = StringVar()
    pin = StringVar()
    ps = StringVar()
    po = StringVar()
    dist = StringVar()
    sc = StringVar()
    hs = StringVar()
    sc = StringVar()
    hsmrk = StringVar()
    smrk = StringVar()
    email = StringVar()
    phone = StringVar()
    dob = StringVar()
    deptVar = StringVar()
    root1 = Toplevel()
    root1.title('Form')
    root1.geometry("680x750")
    root1.config(bg='#326273')

    

    detailFrame = LabelFrame(root1,text="Student's Details",font = ("Georgia", 10, "bold"))
    nameLabel = Label(detailFrame,text="Full Name",font='Calibri')
    nameLabel.grid(row=0,column=0,pady = 10,sticky="w")
    nameEntry = Entry(detailFrame,textvariable=name)
    nameEntry.grid(row=0,column=1,padx=20,pady = 10)

    #LIST OF DEPARTMENTS
    dept = [
        "Computer Science & Engineering","Electronics & Communication Engineering","Electrical Engineering","Artificial Intelligence & Machine Learinging","Information Technology"
    ]
     #'deptVar' is a string variable to store Department name
    deptVar.set("Select Department")
    DeptLabel = Label(detailFrame,text="Dept.",font='Calibri')
    DeptLabel.grid(row=0,column=2,pady = 10,sticky="w")
    dept = ttk.Combobox(detailFrame,values=dept,textvariable=deptVar,).grid(row=0,column=3,pady = 10)

    GenderLabel = Label(detailFrame,text="Gender",font='Calibri')
    GenderLabel.grid(row=1,column=0,pady = 10,sticky="w")
    
    optionframe = Frame(detailFrame)
    Radiobutton(optionframe,text="Male",value='male',font='Calibri',variable=gender).grid(row=1,column=1)
    Radiobutton(optionframe,text="Female",value='female',font='Calibri',variable=gender).grid(row=1,column=2)
    optionframe.grid(row=1,column=1,padx=20,pady = 10)

    ReligionLabel = Label(detailFrame,text="Nationality",font='Calibri')
    ReligionLabel.grid(row=1,column=2,pady = 10,sticky="w")
    religionEntry = Entry(detailFrame,textvariable=religion)
    religionEntry.grid(row=1,column=3,pady = 10)

    dobLabel = Label(detailFrame,text="Date Of Birth",font='Calibri')
    dobLabel.grid(row=2,column=0,pady = 10,sticky="w")
    dob = DateEntry(detailFrame, width=12, background='darkblue', foreground='white', borderwidth=2)
    # dobEntry = Entry(detailFrame,textvariable=dob)
    dob.grid(row=2,column=1,padx=20,pady = 10)

    phoneLabel = Label(detailFrame,text="Phone no.",font='Calibri')
    phoneLabel.grid(row=2,column=2,sticky="w")
    phoneEntry = Entry(detailFrame,textvariable=phone)
    phoneEntry.grid(row=2,column=3)

    emailLabel = Label(detailFrame,text="Email",font='Calibri',)
    emailLabel.grid(row=3,column=0,sticky="w")
    emailEntry = Entry(detailFrame,textvariable=email)
    emailEntry.grid(row=3,column=1)

    detailFrame.pack(ipadx=20,pady=10)


    EducationFrame = LabelFrame(root1,text="Education",font = ("Georgia", 10, "bold"))
    hs_Label = Label(EducationFrame,text="Higher Secondary Board",font='Calibri')
    hs_Label.grid(row=0,column=0,pady = 5,sticky="w",)
    hs_entry = Entry(EducationFrame,textvariable=hs)
    hs_entry.grid(row=0,column=1,pady = 10,ipadx=100)
    s_Label = Label(EducationFrame,text="Secondary Board",font='Calibri')
    s_Label.grid(row=1,column=0,sticky="w",pady = 5)
    s_entry = Entry(EducationFrame,textvariable=sc)
    s_entry.grid(row=1,column=1,ipadx=100)

    marks_frame = Frame(EducationFrame,)
    marks_frame.grid(row=2,column=0,columnspan=2)
    hsmark_Label = Label(marks_frame,text="12th marks(%)",font='Calibri')
    hsmark_Label.grid(row=0,column=0,sticky="w",pady = 10)
    hsmark_entry = Entry(marks_frame,textvariable=hsmrk)
    hsmark_entry.grid(row=0,column=1,pady = 10)
    smark_Label = Label(marks_frame,text="10th marks(%)",font='Calibri')
    smark_Label.grid(row=0,column=2,sticky="w",pady = 10)
    smark_entry = Entry(marks_frame,textvariable=smrk)
    smark_entry.grid(row=0,column=3)
   
    EducationFrame.pack(ipadx=32,pady=10)


    AdressFrame = LabelFrame(root1,text="Adress",font = ("Georgia", 10, "bold"))
    dist_label = Label(AdressFrame,text="State",font='Calibri') 
    dist_label.grid(row=0,column=0,pady = 10,sticky="w")
    dist_entry = Entry(AdressFrame,textvariable=dist)
    dist_entry.grid(row=0,column=1,ipadx=18,padx=10,pady = 10)

    po_label = Label(AdressFrame,text="P.O",font='Calibri') 
    po_label.grid(row=0,column=2,pady = 10,sticky="w")
    po_entry = Entry(AdressFrame,textvariable=po)
    po_entry.grid(row=0,column=3,ipadx=10,pady = 10)

    ps_label = Label(AdressFrame,text="P.S",font='Calibri') 
    ps_label.grid(row=1,column=0,pady = 10,sticky="w")
    ps_entry = Entry(AdressFrame,textvariable=ps)
    ps_entry.grid(row=1,column=1,ipadx=18,padx=10,pady = 10)

    pin_label = Label(AdressFrame,text="Pin Code",font='Calibri') 
    pin_label.grid(row=1,column=2,pady = 10,sticky="w")
    pin_entry = Entry(AdressFrame,textvariable=pin)
    pin_entry.grid(row=1,column=3,ipadx=10,pady = 10)

    AdressFrame.pack(ipadx=70,pady=10)

    ParentFrame = LabelFrame(root1,text="Parents Details",font = ("Georgia", 10, "bold"))
    father_label = Label(ParentFrame,text="Father's Name",font='Calibri')
    father_label.grid(row=0,column=0,pady = 10,sticky="w")
    father_entry = Entry(ParentFrame,textvariable=fatherName)
    father_entry.grid(row=0,column=1,pady = 10,padx=10)
    mother_label = Label(ParentFrame,text="Mother's Name",font='Calibri')
    mother_label.grid(row=0,column=2,pady = 10,sticky="w",padx=10)
    mother_entry = Entry(ParentFrame,textvariable=motherName)
    mother_entry.grid(row=0,column=3,pady = 10)
    phone_label = Label(ParentFrame,text="Phone No.",font='Calibri')
    phone_label.grid(row=1,column=0,pady = 10,sticky="w")
    phone_entry = Entry(ParentFrame,textvariable=Pphone)
    phone_entry.grid(row=1,column=1,pady = 10)
    ParentFrame.pack(ipadx=22,pady=10)

    button_frame = Frame(root1,borderwidth=3,relief=SUNKEN)
    button_frame.pack()

    submit_btn = Button(button_frame,text="Submit",padx = 5,bg="#070739",fg="white",font="lucida 15 bold",command=submit)
    submit_btn.grid(row=0,column=0,)

    show_btn = Button(button_frame,text = "Show",padx = 5,bg="#070739",fg="white",font="lucida 15 bold",command=search.search_window)
    show_btn.grid(row=0,column=1,)

    clearall_btn = Button(button_frame,text = "Clear All",padx = 5,bg="#070739",fg="white",font="lucida 15 bold",command=clear)
    clearall_btn.grid(row=0,column=2,)
    

    root1.mainloop()