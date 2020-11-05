# import required packages
import os
import pandas as pd
from tkinter import *
import random

# setup the path of the working directory
os.chdir("E:/project")
os.getcwd()

# Creating a GUI
w=Tk()

# Setting the title of the window
w.title("Student Details")

# Frame
top_frame = Frame(w, width =450, height=50, pady=3)
top_frame.pack(side=TOP)

# Command for next button
def func():
    global value_list   
    entry_list=[[] for i in range(int(e1.get())+1)]
    value_list=[[] for i in range(int(e1.get())+1)]
    
    for i in range(int(e2.get())+1):
        value_list[0].append(StringVar())
        entry_list[0].append(Entry(center,textvariable=value_list[0][i]))
        entry_list[0][i].grid(row=0,column=i)
        value_list[0][i].set("enter_subject_name")
    
    for i in range(1,int(e1.get())+1):
        for j in range(int(e2.get())+1):
            value_list[i].append(StringVar())
            entry_list[i].append(Entry(center,textvariable=value_list[i][j]))
            entry_list[i][j].grid(row=1+i,column=j)
    
    value_list[0][0].set("Student Name")
    b2=Button(bottom,text="submit",command=makedict)
    b2.grid(row=2,column=1)

# Command for submit button
def makedict():
    global value_list,d,x
    for i in range(len(value_list[0])):
        d[value_list[0][i].get()]=[]
    for i in range(1,len(value_list)):
        for j in range(len(value_list[0])):
            if value_list[i][j].get()=='':
                d[value_list[0][j].get()].append(0)
            else:
                d[value_list[0][j].get()].append(value_list[i][j].get())

    path= StringVar()
    l=Label(bottom, text="file path:").grid(row=1,column=0)
    e3=Entry(bottom, textvariable=path)
    e3.grid(row=1, column=1)
    
    # Random number generation
    x=random.randint(10034, 699812) 
    path.set(os.getcwd()+"\marklist"+str(x))

    # Button
    b4=Button(bottom, text="Close", command=close)
    b4.grid(row=2, column=0)
    
# Command for close button to remove GUI
def close():
    w.destroy()
    
# Labels
l1= Label(top_frame, text="Enter the number of students :")
l1.grid(row=0, column=1)

l2 = Label(top_frame, text="Enter the number of subjects: ")
l2.grid(row=1, column=1)

# Declaration of Entry variables
students = StringVar()
subjects = StringVar()

# Global variable declaration
value_list=[]
d={}
x=0             

# Entry
e1 = Entry(top_frame, textvariable=students)
e1.grid(row=0, column=2)

e2 = Entry(top_frame,textvariable=subjects)
e2.grid(row=1, column=2)

# Button
b1=Button(top_frame, text="Next", command=func)
b1.grid(row=2, column=1)

# initaializing the entry variables
students.set(0)
subjects.set(0)

# Frames
center = Frame(w, width = 2*(int(e2.get())), height = 2*(int(e1.get())), pady=3)
center.pack(side=TOP)

bottom = Frame(w,width = 4, height =4, pady=3)
bottom.pack(side=TOP)

# Prevents from destroying the GUI window until it is          
mainloop()


# Converting a dictonary to a dataframe
marklist=pd.DataFrame(d)

# Reading an excel sheet (.csv file) as a dataframe
subject_credits = pd.read_csv("subject_credits.csv")


l = list(marklist.columns)
credits_copy = subject_credits.set_index('Subject_Name')


s = sum(credits_copy.loc[l[sub],'Credits'] for sub in range(1,len(l))) 
for i in range(1,len(marklist.columns)):
    if credits_copy.loc[l[i],'Maximum_marks'] == 50:
        s+=credits_copy.loc[l[i],'Credits']      

# Formula Implementation
r = []
for i in range(len(marklist)):
    sub_points = 0
    a = 0
    for j in range(1,len(marklist.columns)):
        if credits_copy.loc[l[j],'Maximum_marks'] == 50:
            c = credits_copy.loc[l[j],'Credits']*2  
            a+=1
        if marklist.iloc[i,j] in ["RA","AAA","W","ABS"]:
            p = 0
            continue
        mark=int(marklist.iloc[i,j])
        if a != 0:
            mark*=2
        if 90 <= mark <= 100:
            p = 10
        elif 80 <= mark <= 89:
            p = 9
        elif 70 <= mark <= 79:
            p = 8
        elif 60 <= mark <= 69:
            p = 7
        elif 50 <= mark <= 59:
            p = 6
        elif mark<= 49:
            p = 0     
        if a == 0:
            c = credits_copy.loc[l[j],'Credits']
        sub_points+=c*p   
    r.append(sub_points/s)    

# Appending a new column to the existing dataframe
marklist['GPA'] = r

s="marklist"+str(x)

# Writing to a new excel file
marklist.to_csv(s+".csv",index = False)
