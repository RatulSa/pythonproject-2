from tkinter import *
import form2,search
root = Tk()
root.title('Students Data Entry System')
root.geometry("600x450")
root.config(background="grey9")
root.resizable(0,0) #Disable the maximize/minimize button

mainFrame = Frame(root , borderwidth=5,relief=RAISED)
mainFrame.pack(pady=25)
clg_name = Label(mainFrame ,text="St Thomas College Of Engineering & Technology",font="Bahnschrift 19")
clg_name.pack(pady=10)
logo = PhotoImage(file="STCET-Logo.png")
logo_label = Label(mainFrame , image=logo,)
logo_label.pack(pady=30)
label1 = Label(mainFrame ,text="Student's Data Management System",font="Javanese  12")
label1.pack()

button_frame = Frame(mainFrame,borderwidth=1,relief=SUNKEN)
button_frame.pack(pady=20)
add = Button(button_frame,text="new",font="lucida 15 bold",padx=5,bg="grey",fg="white",command=form2.printf)
add.grid(row=0,column=0,ipadx=10,)
search_btn = Button(button_frame,text="search",font="lucida 15 bold",padx=5,bg="grey",fg="white",command=search.search_window)
search_btn.grid(row=0,column=1,ipadx=10,)
exit = Button(button_frame,text="exit",font="lucida 15 bold",padx=5,bg="grey",fg="white",command=exit)
exit.grid(row=0,column=2,ipadx=10,)

root.mainloop()
