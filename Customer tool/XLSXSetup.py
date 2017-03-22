'''
@author: 10204728
'''
import os
from tkinter import *
import Ann.GUI_Setup_SMP

root = Tk() # Tk Object
app = Ann.GUI_Setup_SMP.DemoGUI(master=root)
#start the program
app.mainloop()

'''
Input_file = open("Input.txt","r")
path = "./Output.txt"
existed = os.path.exists(path.strip())
if (existed != True):
    Output_file = open(path,"w")
    for line in Input_file:
        for word in range(len(line)):
            if line[word] != "\n":
                Output_file.write(line[word] + ",")
            else:
                Output_file.write(line[word])
    Output_file.close()

Input_file.close()
'''