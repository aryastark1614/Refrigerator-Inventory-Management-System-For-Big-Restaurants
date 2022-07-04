#A GUI Function to Add Items

def Item_Register():
    from tkinter import messagebox
    import pandas as pd
    import openpyxl
    global serial
    Name = Info1.get()
    Rack =Info2.get()
    Expiry_Date = Info4.get()
    Shelf =Info3.get()
    
    if(Name=='' or Rack=='' or Expiry_Date=='' or Shelf==''):
        messagebox.showwarning("Empty Spaces Not Allowed!!!")
        return
    else:
        wb = openpyxl.load_workbook('Project.xlsx')
        sheet = wb.active
        ser =0
        df=pd.read_excel('Project.xlsx',usecols=[0,0])
        if df.empty:            
            serial=1
        else:
            ser= df.iloc[-1]['Serial no']
            serial=ser+1
        about_book=[(serial,Name,Rack,Expiry_Date,Shelf)]
        for i in about_book:  
            sheet.append(i)
        
        wb.save('Project.xlsx')
    
        messagebox.showinfo('Success',"Item added successfully")
    
        Info1.delete(0, END)
        Info2.delete(0, END)
        Info3.delete(0, END)
        Info4.delete(0, END)
        messagebox.showinfo("Want to add more Items?","Enter the data in the blank space and press 'Submit' or press 'Quit' to exit")
    
        #root.destroy()

from tkinter import *
serial=0
def addItem(): 
    
    global Info1 ,Info2,Info3, Info4
    
    root = Tk()
    root.title("Add Items")
    root.minsize(width=400,height=400)
    root.geometry("700x500")
    
    Canvas1 = Canvas(root)
    Canvas1.config(bg="#ffff00")
    Canvas1.pack(expand=True,fill=BOTH)
    
    Name=StringVar()
    Rack=StringVar()
    Expiry_Date=StringVar()
    Shelf=StringVar()
        
    headingFrame1 = Frame(root,bg="#ff0000",bd=5)
    headingFrame1.place(relx=0.25,rely=0.03,relwidth=0.5,relheight=0.12)
    headingLabel = Label(headingFrame1, text="Add Items", bg="#FFA07A", fg='#151c52', font=('Calibre',19,'bold'))
    headingLabel.place(relx=0,rely=0, relwidth=1, relheight=1)
    
    labelFrame = Frame(root,bg="#FFA07A")
    labelFrame.place(relx=0.1,rely=0.20,relwidth=0.8,relheight=0.5)
        
    # Item_Name
    lb1 = Label(labelFrame,text="Name : ", bg="#FFA07A", fg='black',font=("calibre",14,'bold'))
    lb1.place(relx=0.05,rely=0.10, relheight=0.08)
        
    Info1 = Entry(labelFrame,justify='center',textvariable=Name,font=("calibre",14,'bold'))
    Info1.place(relx=0.4,rely=0.10, relwidth=0.56, relheight=0.09)
        
    # Rack
    lb2 = Label(labelFrame,text="Rack : ", bg="#FFA07A", fg='black',font=("calibre",14,'bold'))
    lb2.place(relx=0.05,rely=0.25, relheight=0.08)
        
    Info2 = Entry(labelFrame,textvariable=Rack,font=("calibre",14,'bold'),justify='center')
    Info2.place(relx=0.4,rely=0.25, relwidth=0.56, relheight=0.09)
    
    #Addition Expiry_Date
    lb3 = Label(labelFrame,text="Shelf :", bg="#FFA07A", fg='black',font=("calibre",14,'bold'))
    lb3.place(relx=0.05,rely=0.40, relheight=0.08)
        
    Info3 = Entry(labelFrame,textvariable=Expiry_Date,font=("calibre",14,'bold'),justify='center')
    Info3.place(relx=0.4,rely=0.40, relwidth=0.56, relheight=0.09)
    
    #Shelf
    lb4 = Label(labelFrame,text="Expiry Date:", bg="#FFA07A", fg='black',font=("calibre",14,'bold'))
    lb4.place(relx=0.05,rely=0.55, relheight=0.08)
        
    Info4 = Entry(labelFrame,textvariable=Shelf,font=("calibre",14,'bold'),justify='center')
    Info4.place(relx=0.4,rely=0.55, relwidth=0.56, relheight=0.09)
        
    #Submit Button
    SubmitBtn= Button(root,text="Submit",bg='#d4d4d4', fg='black',command= Item_Register,font=("calibre",12,'bold'))
    SubmitBtn.place(relx=0.28,rely=0.85, relwidth=0.18,relheight=0.08)
    #quit button
    quitBtn = Button(root,text="Quit",bg='#dedede', fg='black', command=root.destroy,font=("calibre",13,'bold'))
    quitBtn.place(relx=0.53,rely=0.85, relwidth=0.18,relheight=0.08)
    
    root.mainloop()
def main():
    addItem()

if __name__ == "__main__": 
    main()
