#For interaction
from textwrap import fill
from tkinter import messagebox
#For File manipulation
from tkinter import filedialog
import os
import openpyxl 
#For UI 
from tkinter import *
from turtle import clear, color
def show_frameClrTxt(frame):
  def clear_text():
     entrypw.delete(0,END)
     entryuser.delete(0,END)
     
    

  clear_text()   
   
  frame.tkraise() 
def show_frame(frame):
 
 frame.tkraise() 
def verify():
            try:
                with open("credential.txt", "r") as f:
                    info = f.readlines()
                    i  = 0
                    for e in info:
                        u, p =e.split(",")
                        if u.strip() == entryuser.get() and p.strip() == entrypw.get():
                            show_frame(frame3)
                            
                            
                            i = 1
                            break
                    if i==0:
                        messagebox.showinfo("Error", "Please provide correct username and password!!")
            except:
                messagebox.showinfo("Error", "Please provide correct username and password!!")
def exit():
    root.destroy()
def button_hover(button, bg_color, default_color):
    def on_enter(e):
        button['background'] = bg_color

    def on_leave(e):
        button['background'] = default_color


    button.bind('<Enter>', on_enter)
    button.bind('<Leave>', on_leave)

# Define a function toupload
def upload_file():
    # Open a file dialog to select an Excel file
    file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
    # Check if a file was selected
    if file_path:
        # Save the file path for later use
        frame4.file_path = file_path
        # Show a message box to confirm the upload
        messagebox.showinfo("Upload", "File uploaded successfully!")

#function to modify the file
def modify_file():
    # Check if a file has been uploaded
    if hasattr(frame4, 'file_path'):
        # Load the workbook from the uploaded file
        workbook = openpyxl.load_workbook(frame4.file_path)
        # Get the active worksheet
        worksheet = workbook.active
        # Modify the worksheet by adding some data
        worksheet.append([var1.get(),var2.get(),var4.get(),var5.get(),var6.get(),var7.get(),var8.get()])
        # Save the modified workbook back to the original file
        workbook.save(frame4.file_path)
        # Show a message box to confirm the modification
        messagebox.showinfo("Modification", "File modified successfully!")
    else:
        # If no file has been uploaded, show an error message
        messagebox.showerror("Error", "Please upload a file first")
#for opening the file
def open_excel_file():
    
    if hasattr(frame4, 'file_path'):
        # Use the os module to open the file with the default program
        os.startfile(frame4.file_path)
    else:
        # If no file has been uploaded, show an error message
        messagebox.showerror("Error", "Please upload a file first")

def clear_text():
     entrypw.delete(0,END)
     entryuser.delete(0,END)
     e1.delete(0,END)
     e2.delete(0,END)
     e4.delete(0,END)
     e5.delete(0,END)
     e6.delete(0,END)
     e7.delete(0,END)
     e8.delete(0,END)
def combined_function(frame):
    show_frame(frame)
    clear_text()
root=Tk()


#root.state("zoomed")
root.maxsize(1330,700)
root.title("Result Analysis")




root.rowconfigure(0,weight=1)
root.columnconfigure(0,weight=1)


frame1=Frame(root)
frame2=Frame(root)
frame3=Frame(root)
frame4=Frame(root)
frame5=Frame(root)
frame6=Frame(root)
frame7=Frame(root)

for frame in (frame1,frame2,frame3,frame4,frame5,frame6,frame7):
    frame.grid(row=0,column=0,sticky='nsew')

    
#__________FRame1 Code
canvas_frame1 = Canvas(frame1,bg = "#ffffff",height = 700,width = 1330,bd = 0,highlightthickness = 0,relief = "ridge")
canvas_frame1.pack(fill='y')
background_img = PhotoImage(file = f"background.png")
background = canvas_frame1.create_image(670.0, 506.5,image=background_img)
img0 = PhotoImage(file = f"img0.png")
result_a=Label(frame1,text="Result Analysis",font="Calibary 28 bold")
result_a.place(x=100,y=600)
YEs = Button(frame1,text="Yes",font=("Arial",24),bg="black",fg="white",bd=5,borderwidth = 5,highlightthickness = 2,relief = "ridge",command=lambda:show_frame(frame2))
YEs.place(x = 593, y = 383,width = 149,height = 50)
#Hovering function
button_hover(YEs,'blue','black')

img1 = PhotoImage(file = f"img1.png")
No = Button(frame1,text="No",font=("Arial",24),bg="black",fg="white",borderwidth = 5,bd=5,highlightthickness = 2,command = exit,relief = "ridge")
No.place(x = 593, y = 507,width = 149,height = 50)
button_hover(No,'red','black')


#__________FRame2 Code
canvas_frame2 = Canvas(frame2,bg = "#ffffff",height = 700,width = 1330,bd = 0,highlightthickness = 0,relief = "ridge")
canvas_frame2.pack(fill='y')

background2_img = PhotoImage(file = f"Login Pagemainbg.png")
background2 =  canvas_frame2.create_image(665.0, 348.5,image=background2_img)

user=Label(frame2,text="User Id",font="calibary 20 bold",bg='black',fg='white')
user.place(x=449,y=300)
password=Label(frame2,text="Password",font="calibary 20 bold",bg='black',fg='white')
password.place(x=449,y=427)
entry0_img = PhotoImage(file = f"img_textBox0.png")
entry0_bg =  canvas_frame2.create_image(678.5, 372.5,image = entry0_img)

entryuser = Entry(frame2,bd = 0,bg = "#ffffff",highlightthickness = 0,font="Calibary 30 bold")

entryuser.place(x = 454.0, y = 349,width = 449.0,height = 50)

entry1_img = PhotoImage(file = f"img_textBox1.png")
entry1_bg =  canvas_frame2.create_image(678.5, 498.5,image = entry1_img)

entrypw = Entry(frame2,bd = 0,bg = "#ffffff",highlightthickness = 0,font="Calibary 30 bold")

entrypw.place(x = 454.0, y = 475,width = 449.0,height = 50)

frame2_btn=Button(frame2,text='Enter',font="Calibary 20 bold ",command=verify,width=5,borderwidth=3,bg='white')
frame2_btn.place(x=610,y=550)
button_hover(frame2_btn,'green','white')

#bind enter key to button 
frame2_btn.bind('<Return>',lambda event:verify())



exitb=Button(frame2,text="Exit",font="Calibary 20 bold",bg='red',fg='white',borderwidth=2,command=exit)
exitb.place(x=1150,y=250)
exitb.bind('<Return>',lambda event:exit())
#bind exit key to button
root.bind('<Escape>',lambda event:exit())

#__________FRame3 Code
canvas_frame3 = Canvas(frame3,bg = "#84dfd4",height = 700,width = 1330,bd = 0,highlightthickness = 0,relief = "ridge")
canvas_frame3.pack(fill='y')

img2 = PhotoImage(file = f"uploadb.png")
uploadb = Button(frame3,image = img2,borderwidth = 0,highlightthickness = 0,relief = "flat",command=lambda:show_frame(frame4))

uploadb.place(x = 252, y = 469,width = 727,height = 86)

img3 = PhotoImage(file = f"currentb.png")
currentd = Button(frame3,image = img3,borderwidth = 0,highlightthickness = 0,command=lambda:show_frame(frame5),relief = "flat")

logout=Button(frame3,text="logout",font="Calibary 20 bold",bg='red',fg='black',height=1,width=5,command=lambda:show_frameClrTxt(frame2),borderwidth=4)
logout.place(x=1200,y=150)
currentd.place(x = 340, y = 245,width = 555,height = 86)

background3_img = PhotoImage(file = f"background3.png")
background3 = canvas_frame3.create_image(665.0, 156.0,image=background3_img)

#______________Frame4 Code
canvas_frame4 = Canvas(frame4,bg = "#84dfd4",height = 700,width = 1330,bd = 0,highlightthickness = 0,relief = "ridge")
canvas_frame4.place(x = 0, y = 0)

img4 = PhotoImage(file = f"back1.png")
back1 = Button(frame4,image = img4,borderwidth = 0,highlightthickness = 0,command = lambda:show_frame(frame3),relief = "flat")

back1.place(x = 50, y = 643,width = 159,height = 38)

logout=Button(frame4,text="logout",font="Calibary 20 bold",bg='red',fg='black',height=1,width=5,command=lambda:show_frameClrTxt(frame2),borderwidth=4)
logout.place(x=1200,y=150)

img5 = PhotoImage(file = f"uploadexcel.png")
uploadexcel = Button(frame4,image = img5,borderwidth = 0,highlightthickness = 0,relief = "flat",command=upload_file)

uploadexcel.place(x = 498, y = 277,width = 300,height = 80)

background4_img = PhotoImage(file = f"background3.png")
background4 = canvas_frame4.create_image(665.0, 156.0,image=background4_img)


#______________Frame5 code
canvas_frame5 = Canvas(frame5,bg = "#84dfd4",height = 700,width = 1330,bd = 0,highlightthickness = 0,relief = "ridge")
canvas_frame5.pack(fill='y')

img6 = PhotoImage(file = f"img6.png")
opensheet = Button(frame5,image = img6,borderwidth = 0,highlightthickness = 0,relief = "flat",command=open_excel_file)

opensheet.place(x = 469, y = 302,width = 350,height = 38)

img7 = PhotoImage(file = f"img7.png")
modify = Button(frame5,image = img7,borderwidth = 0,highlightthickness = 0,relief = "flat",command=lambda: show_frame(frame7))

modify.place(x = 550, y = 393,width = 180,height = 38)
logout=Button(frame5,text="logout",font="Calibary 20 bold",bg='red',fg='black',height=1,width=5,command=lambda:show_frameClrTxt(frame2),borderwidth=4)
logout.place(x=1200,y=150)

img8 = PhotoImage(file = f"img8.png")
back2 = Button(frame5,image = img8,borderwidth = 0,highlightthickness = 0,command =lambda:show_frame(frame3) ,relief = "flat")

back2.place(x = 50, y = 609,width = 159,height = 38)

img9 = PhotoImage(file = f"img9.png")
toppers = Button(frame5,image = img9,borderwidth = 0,highlightthickness = 0,relief = "flat",command =lambda:show_frame(frame6))

toppers.place(x = 550, y = 484,width = 187,height = 38)

background5_img = PhotoImage(file = f"background3.png")
background5 = canvas_frame5.create_image(665.0, 156.0,image=background5_img)

#_________Frame6Code
canvas_frame6 = Canvas(frame6,bg = "#84dfd4",height = 700,width = 1330,bd = 0,highlightthickness = 0,relief = "ridge")
canvas_frame6.pack(fill='y')

img10 = PhotoImage(file = f"img10.png")
SM = Button(frame6,image = img10,borderwidth = 0,highlightthickness = 0, relief = "flat")
SM.place(x = 438, y = 337,width = 439,height = 45)

logout=Button(frame6,text="logout",font="Calibary 20 bold",bg='red',fg='black',height=1,width=5,command=lambda:show_frameClrTxt(frame2),borderwidth=4)
logout.place(x=1200,y=150)

img11 = PhotoImage(file = f"img11.png")
pex = Button(frame6,image = img11,borderwidth = 0,highlightthickness = 0,relief = "flat")

pex.place(x = 416, y = 414,width = 494,height = 45)

img12 = PhotoImage(file = f"img12.png")
LA = Button(frame6,image = img12,borderwidth = 0,highlightthickness = 0,relief = "flat")

LA.place(x = 488, y = 264,width = 349,height = 45)

img13 = PhotoImage(file = f"img13.png")
overall = Button(frame6,image = img13,borderwidth = 0,highlightthickness = 0,relief = "flat")

overall.place( x = 560, y = 134,width = 188,height = 45)

img14 = PhotoImage(file = f"img14.png")
POE = Button(frame6,image = img14,borderwidth = 0,highlightthickness = 0,relief = "flat")

POE.place( x = 405, y = 488, width = 515,height = 45)

img15 = PhotoImage(file = f"img15.png")
DSA = Button(frame6,image = img15,borderwidth = 0,highlightthickness = 0,relief = "flat")

DSA.place(x = 330, y = 561,width = 656,height = 45)

img16 = PhotoImage(file = f"img16.png")
back3 = Button(frame6,image = img16,borderwidth = 0,highlightthickness = 0,relief = "flat",command=lambda:show_frame(frame5))

back3.place(x = 50, y = 609,width = 159,height = 38)

background6_img = PhotoImage(file = f"background3.png")
background6 = canvas_frame6.create_image(665.0, 156.0,image=background6_img)

show_frame(frame1)

#Frame7__________Code

canvas_frame7 = Canvas(frame7,bg = "#84dfd4",height = 700,width = 1330,bd = 0,highlightthickness = 0,relief = "ridge")
canvas_frame7.pack(fill='y')

l1=Label(frame7,text="Enter Name of student : ")
l2=Label(frame7,text="Enter PRN number : ")
l3=Label(frame7,text="ADD subjectwise marks")
l4=Label(frame7,text="Linear Algebra : ")
l5=Label(frame7,text="Statistical Methods : ")
l6=Label(frame7,text="Principle of electronics : ")
l7=Label(frame7,text="Principle of Economics : ")
l8=Label(frame7,text="Data structures and algorithms : ")

l1.place(x = 460, y = 120,width = 200,height = 45)
l2.place(x = 460, y = 180,width = 200,height = 45)
l3.place(x = 620, y = 240,width = 200,height = 45)
l4.place(x = 460, y = 300,width = 200,height = 45)
l5.place(x = 460, y = 360,width = 200,height = 45)
l6.place(x = 460, y = 420,width = 200,height = 45)
l7.place(x = 460, y = 480,width = 200,height = 45)
l8.place(x = 460, y = 540,width = 200,height = 45)

var1=StringVar()
var2=StringVar()
var4=IntVar()
var5=IntVar()
var6=IntVar()
var7=IntVar()
var8=IntVar()

e1=Entry(frame7,textvariable=var1,font="Calibary 18 bold")
e2=Entry(frame7,textvariable=var2,font="Calibary 18 bold")
e4=Entry(frame7,textvariable=var4,font="Calibary 18 bold")
e5=Entry(frame7,textvariable=var5,font="Calibary 18 bold")
e6=Entry(frame7,textvariable=var6,font="Calibary 18 bold")
e7=Entry(frame7,textvariable=var7,font="Calibary 18 bold")
e8=Entry(frame7,textvariable=var8,font="Calibary 18 bold")

e1.place(x = 720, y = 120,width = 400,height = 45)
e2.place(x = 720, y = 180,width = 250,height = 45)

e4.place(x = 720, y = 300,width = 200,height = 45)
e5.place(x = 720, y = 360,width = 200,height = 45)
e6.place(x = 720, y = 420,width = 200,height = 45)
e7.place(x = 720, y = 480,width = 200,height = 45)
e8.place(x = 720, y = 540,width = 200,height = 45)

b1=Button(frame7,text="ADD",font=("Arial",24),bg="black",fg="white",bd=5,borderwidth = 5,highlightthickness = 2,relief = "ridge",command=modify_file)
b1.place(x=1120,y=500,height=100,width=100)


back4 = Button(frame7,image = img16,borderwidth = 0,highlightthickness = 0,relief = "flat",command=lambda:combined_function(frame5))

back4.place(x = 50, y = 609,width = 159,height = 38)
background7_img = PhotoImage(file = f"background3.png")
background7 = canvas_frame7.create_image(665.0, 156.0,image=background7_img)


root.mainloop()
