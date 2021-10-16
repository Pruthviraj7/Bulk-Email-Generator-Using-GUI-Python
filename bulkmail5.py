from tkinter import *
from email.mime import text
from tkinter import font
from PIL import ImageTk,Image 
import pandas as pd
import os
import time
import email_function
from tkinter import messagebox,filedialog
import webbrowser

class Bulk_Email:

    def __init__(self,root3):
        self.root3 = root3
        self.root3.title("Intro Page: Introduction about Nirma")
        self.root3.geometry("1000x450")
        self.root3.focus_force()
        self.root3.grab_set()
        self.root3.resizable(False,False)
        self.root3.config(bg = "grey")

        #intro
        Ititle=Label(self.root3,text = "Mission", font = ("Goudy Old Style",30,"bold"),fg="black",bg="grey").place(x=600,y=10)
        Idesc =Label(self.root3,text = """\n Nirma University places a strong focus on its students' overall development.\n Its goal is to produce not just competent professionals,\n but also decent and deserving citizens of a great country \n who will contribute to the country's overall advancement and development. 
        It strives to treat each student as a person, \n to recognise their potential, and to provide the finest preparation and\n training possible to help them achieve their professional and life objectives.""", font = ("Goudy Old Style",13),fg="black",bg="grey").place(x=150,y=50,relwidth=1)
        
        made = Label(self.root3,text = "Made By:",font = 15,fg="black",bg="grey").place(x=10,y=260)
        n1 = Label(self.root3,text = "Keyuri Kariya(19BEC051)",font = ("Goudy Old Style",15,"bold"),fg="black",bg="grey").place(x=10,y=300)
        n2 = Label(self.root3,text="Pruthvirajsinh Solanki(19BEC126)", font = ("Goudy Old Style",15,"bold"),fg="black",bg="grey").place(x=10,y=330)


        subm = Label(self.root3,text = "Mentored by:",font = 15,fg="black",bg="grey").place(x=800,y=260)
        N1 = Label(self.root3,text = "Dr. Sachin Gajjar",font = ("Goudy Old Style",15,"bold"),fg="black",bg="grey").place(x=800,y=300)
        N2 = Label(self.root3,text="Purvansh Shah", font = ("Goudy Old Style",15,"bold"),fg="black",bg="grey").place(x=800,y=330)

        #Logo
        logo=ImageTk.PhotoImage(Image.open('nirma3.png'))
        logo_label=Label(image=logo)
        logo_label.image=logo
        logo_label.grid(column=2,row=5)

        b =Button(self.root3,text="Next page",command=self.secondPage,font=(20)).place(x=500,y=400)

    
    def secondPage(self):

        self.root4=Toplevel()
        self.root4.title("Demo page")
        self.root4.geometry("1000x600")
        self.root4.resizable(False,False)
        self.root4.config(bg = "grey")

        title=Label(self.root4,text = "Project demo Page", font = ("Goudy Old Style",40,"bold"),fg="white",bg="black").place(x=0,y=0,relwidth=1)
        
        def callback(url):
            webbrowser.open_new_tab(url)

        #Create a Label to display the link
        DEsc = Label(self.root4,text = "About our project", font = ("Goudy Old Style",20,"bold"),fg="black",bg="grey").place(x=0,y=100,relwidth=1)
        Ddesc =Label(self.root4,text = """\n This project sends mail in bulk to selected person whose name is their in excel sheet.\n The main goal of the project is to send mail to every person in one go with the required message we want to send to them.\n First the sender needs to sign in their gmail account in the sender's info section.\n After that the sender can send mail by selecting appropriate excel file.\n While sending the mail the user can see in the status bar below that how many mails are sent and which one's are remaining. \nIn the above video we have demonstrated how to use the gui and how to send mail to anyone""", font = ("Goudy Old Style",13),fg="black",bg="grey").place(x=40,y=150,relwidth=1)
        link = Label(self.root4, text="Demo Video:", fg="black",bg="grey" , font= ("Goudy Old Style",18,"bold"))
        link.place(x=400,y=330)
        link1 = Label(self.root4, text="Video", fg="blue",bg="grey" , cursor="hand2",font= ("Goudy Old Style",15,"bold"))
        link1.pack(expand=True,padx=100,pady=100)
        link1.place(x=550,y=335)
        link1.bind("<Button-1>", lambda e: callback("https://drive.google.com/drive/folders/14J4zEubEKpXHAzvEVfNatsOzrkbBf8Lb?usp=sharing"))

        b1 =Button(self.root4,text="Next page",command=self.Buulk,font=(20)).place(x=450,y=550)

   
    def Buulk(self):
        #Box Size
        self.root=Toplevel()
        self.root.title("Bulk Email GUI")
        self.root.geometry("1000x600")
        self.root.resizable(False,False)
        self.root.config(bg = "grey")

        
        #self.setting_icon=ImageTk.PhotoImage(file="F:\GUI Python\Project\logos\logonirma.png")
        #Title
        title=Label(self.root,text = "Bulk Email Sender", font = ("Goudy Old Style",40,"bold"),fg="white",bg="black").place(x=0,y=0,relwidth=1)
        desc = Label(self.root, text="We use Excel file to send email in bulk.You have only two option. 1.Bulk Emails, 2.Single Emails", font=("Goudy Old Style", 17), fg="black",bg="grey")
        desc.place(x=0, y=65, relwidt=1)
        
        #button
        btn_setting=Button(self.root,command=self.setting_window,text="Sender Info",fg="black",activebackground="black").place(x=880, y = 20)

        #RadioButton
        self.var_choice=StringVar()
        single=Radiobutton(self.root,text="Single",value="single",variable=self.var_choice,bg="grey",fg="black",font=("Times New Roman",20),command=self.check_single_or_bulk).place(x=10,y=120)
        bulk=Radiobutton(self.root,text="Bulk",value="bulk",variable=self.var_choice,bg="grey",fg="black",font=("Times New Roman",20),command=self.check_single_or_bulk).place(x=150,y=120)
        self.var_choice.set("single")

        #main.......
        work = Label(self.root, text = "WORK", font = ("Times New Roman",20,"bold"),fg="white",bg="black")
        work.place(x=0,y=190,relwidth=1)
        to=Label(self.root,text="To (Email Address)",font=("Times New Roman",15,"bold"),fg="black",bg="grey").place(x=0,y=240)
        subj=Label(self.root,text="Subject",font=("Times New Roman",15,"bold"),fg="black",bg="grey").place(x=0,y=280)
        msgtxt=Label(self.root,text="Message",font=("Times New Roman",15,"bold"),fg="black",bg="grey").place(x=0,y=320)


        #Entry Box
        self.txt_to=Entry(self.root,font=("Times new Roman",10,"bold"))
        self.txt_to.place(x=300,y=240,width=400,height=30)
        self.txt_subj=Entry(self.root,font=("Times new Roman",10,"bold"))
        self.txt_subj.place(x=300,y=280,width=400,height=30)
        self.txt_msgtxt=Text(self.root,font=("Times new Roman",10,"bold"))
        self.txt_msgtxt.place(x=300,y=320,width=400,height=100)

        #BUTtons
        self.btn_browse=Button(self.root,text="Browse",command=self.browse_file, fg="black", activebackground="black", font=("bold"), cursor="hand2",
                      activeforeground="dark red",state=DISABLED)
        self.btn_browse.place(x=310, y=125, width=100)

        btn_send=Button(self.root,text="Send",command=self.send_email,fg="black", activebackground="black",font=("bold"),cursor="hand2",activeforeground="dark red").place(x=300, y=450,width=100)
        btn_clr=Button(self.root,text="Clear",command=self.clear1,fg="black", activebackground="black",font=("bold"),cursor="hand2",activeforeground="dark red").place(x=410, y=450,width=100)

        
        #Status bar
        self.status = Label(self.root, font=("Times New Roman", 15, "bold"), fg="black",bg="grey")
        self.status.place(x=0, y=500)
        self.lbl_total=Label(self.root,font=("Times New Roman",15,"bold"),fg="black",bg="grey")
        self.lbl_total.place(x=300,y=500)
        
        self.lbl_sent=Label(self.root,font=("Times New Roman",15,"bold"),fg="black",bg="grey")
        self.lbl_sent.place(x=400,y=500)

        self.lbl_left=Label(self.root,font=("Times New Roman",15,"bold"),fg="black",bg="grey")
        self.lbl_left.place(x=500,y=500)

        self.lbl_failed=Label(self.root,font=("Times New Roman",15,"bold"),fg="black",bg="grey")
        self.lbl_failed.place(x=600,y=500)

        self.check_file_exist()

    def setting_window(self):
        self.check_file_exist()
        self.root2=Toplevel()  #Same as parent window title
        self.root2.title("Mail settings")
        self.root2.geometry("1000x500")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="grey")

        title2=Label(self.root2,text="Sender Email Setting",font=("Times New Roman",48,"bold"),fg="white",bg="black").place(x=0,y=0,relwidth=1)
        desc2 = Label(self.root2,
                      text="Enter your Email id and password",
                      font=("Goudy Old Style", 25,"italic"), fg="black",
                      bg="grey").place(x=0, y=75)
        
        from_=Label(self.root2,text="Email Address", font=("Times New Roman", 15), fg="black",bg="grey").place(x=50,y=150)
        pass_=Label(self.root2,text="Password",font=("Times New Roman",15,"bold"),fg="black",bg="grey").place(x=50,y=200)        
        
        self.txt_from=Entry(self.root2,font=("Times new Roman",15,"bold"))
        self.txt_from.place(x=300,y=150,height=30)

        self.txt_pass=Entry(self.root2,font=("Times new Roman",15,"bold"),show="*")
        self.txt_pass.place(x=300,y=200,height=30)

        btn_save=Button(self.root2,text="Save",command=self.save_setting,fg="black", activebackground="black", font=("bold"), cursor="hand2",
                       activeforeground="dark red").place(x=150, y=260, width=100)
        btn_clr2=Button(self.root2,text="Clear",command=self.clear2,fg="black", activebackground="black", font=("bold"), cursor="hand2",
                       activeforeground="dark red").place(x=290, y=260, width=100) 

        self.txt_from.insert(0,self.from_)
        self.txt_pass.insert(0,self.pass_)
    
    def check_single_or_bulk(self):
        if self.var_choice.get()=="single":
            self.btn_browse.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.clear1()
        if self.var_choice.get()=="bulk":
            self.btn_browse.config(state=NORMAL)
            self.txt_to.delete(0,END)
            self.txt_to.config(state='readonly')

    def send_email(self):
        x=len(self.txt_msgtxt.get('1.0',END))
        if self.txt_to.get()=="" or self.txt_subj.get()=="" or x==1:
            messagebox.showerror("Error,All field required",parent=self.root)
        else:
            if self.var_choice.get()=="single":
                status=email_function.email_send(self.txt_to.get(),self.txt_subj.get(),self.txt_msgtxt.get('1.0',END),self.from_,self.pass_)
                if status=="s":
                    messagebox.showinfo("Email has been sent",parent=self.root)
                if status=="f":
                    messagebox.showerror("Email has not been sent",parent=self.root)
            if self.var_choice.get()=="bulk":
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=email_function.email_send(x,self.txt_subj.get(),self.txt_msgtxt.get('1.0',END),self.from_,self.pass_)
                    if status=="s":
                        self.s_count+=1
                    elif status=="f":
                        self.f_count+=1

                    self.status_bar()
                    
                messagebox.showinfo("Email has been sent",parent=self.root)
    def status_bar(self):
        self.lbl_total.config(text="Status: "+str(len(self.emails))+"=>>")
        self.lbl_sent.config(text="Sent: "+str(self.s_count))
        self.lbl_left.config(text="Left: "+str(len(self.emails)-(self.s_count+self.f_count)))
        self.lbl_failed.config(text="Failed: "+str(self.f_count))
        self.lbl_total.update()
        self.lbl_sent.update()
        self.lbl_left.update()
        self.lbl_failed.update()



    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0,END)
        self.txt_subj.delete(0,END)
        self.txt_msgtxt.delete('1.0',END)
        self.var_choice.set("single")
        self.btn_browse.config(state=DISABLED)
        self.lbl_total.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_left.config(text="")
        self.lbl_failed.config(text="")

    
    def clear2(self):
        self.txt_from.delete(0,END)
        self.txt_pass.delete(0,END)

    def check_file_exist(self):
        if os.path.exists("important.txt")==False:
            f=open('important.txt','w')
            f.write(",")
            f.close()
        f2=open('important.txt','r')
        self.credentials=[]
        for i in f2:
            self.credentials.append( [i.split(",")[0],i.split(",")[1]] )
        self.from_=self.credentials[0][0]
        self.pass_=self.credentials[0][1]
        



    def save_setting(self):
        if self.txt_from.get()=="" or self.txt_pass.get()=="":
            messagebox.showerror("Error,All field required",parent=self.root2)
        else:
            f=open('important.txt','w')
            f.write(self.txt_from.get()+","+self.txt_pass.get())
            f.close()
            messagebox.showinfo("Saved Successfully",parent=self.root2)
            self.check_file_exist()
    def browse_file(self):
        op=filedialog.askopenfile(initialdir='/',title="Select Excel File For Emails",filetypes=(("All files","*.*"),("Excel Files",".xlsx")))
        if op!=None:
            data=pd.read_excel(op.name)
            if "Email" in data.columns:
                self.emails=list(data['Email'])
                c=[]
                for i in self.emails:
                    if pd.isnull(i)==False:
                        c.append(i)
                self.emails=c
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0,END)
                    self.txt_to.insert(0,str(op.name.split("/")[-1]))
                    self.txt_to.config(state="readonly")
                    self.lbl_total.config(text="Total: "+str(len(self.emails)))
                    self.lbl_sent.config(text="Sent: ")
                    self.lbl_left.config(text="Left: ")
                    self.lbl_failed.config(text="Failed: ")

                else:
                     messagebox.showerror("Error,This file does not have any email columns",parent=self.root)
                
            else:
                messagebox.showerror("Error,Select which has email colums",parent=self.root)







root3=Tk()
obj=Bulk_Email(root3)
root3.mainloop()