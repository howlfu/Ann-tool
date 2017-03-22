#coding=utf-8
from tkinter import *
from cgitb import text
#import os
from Ann.Members import OBJofXL
from Ann.Members import OBJofXL2

from _overlapped import NULL
from distutils.cmd import Command
from openpyxl.chart.data_source import StrVal
from doctest import master
        
class DemoGUI(Frame):
    Setup_Data = object
    row_get = False
    New_Member = False
    path_in_use = './Ann_SMP.xlsx'
    page_in_use = 'CUST_HEAD'
    DTL_path_in_use = './Ann_DTL.xlsx'
    DTL_page_in_use = 'CUST_DTL'
    mylist1 = [u'重接', u'整理']
    Customer_Info = [u'姓名', u'生日', u'根數', u'電話', u'睫毛次數', u'除毛次數', u'入會日', u'最近消費日', u'上次重接日', u'預繳餘額', u'Line', u'備註', ]
    CheckVar1 = 0
    ReDo = False
    error_log = ""
    
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid()
        self.createWidgets()
        self["background"] = "pink"
        #Instance of Members setting
        try:
            self.Setup_Data = OBJofXL(self.path_in_use, self.page_in_use)
            self.Setup_Data.Check_workpage_sheetname(self.page_in_use)
            self.Setup_Data_Detail = OBJofXL2(self.DTL_path_in_use, self.DTL_page_in_use)
            self.Setup_Data_Detail.Check_workpage_sheetname(self.DTL_page_in_use)
        except:
            self.error_log = "File loading is not existed"
        self.Output_Infor(self.error_log, True)
        
    def createWidgets(self):
        #input Text
        #1
        self.Text_for_save1 = Label(self, text = u"姓名", font = 16, background = "PaleVioletRed2")
        self.Text_for_save1.grid(row=0, column=0)
        #input Entry
        self.te1 = Entry(self)
        self.te1["width"] = 20 # Byte for 50 words
        self.te1.grid(row=0, column=1, columnspan=2)
        #2
        self.Text_for_save2 = Label(self, text = u"電話", font = 16, background = "PaleVioletRed2")
        self.Text_for_save2.grid(row=1, column=0)
        #input Entry
        self.te2 = Entry(self)
        self.te2["width"] = 20 # Byte for 50 words
        self.te2.grid(row=1, column=1, columnspan=2)
        #3
        self.Text_for_save3 = Label(self, text = u"生日", font = 16, background = "PaleVioletRed2")
        self.Text_for_save3.grid(row=2, column=0)
        self.te3 = Entry(self)
        self.te3["width"] = 20 # Byte for 50 words
        self.te3.grid(row=2, column=1, columnspan=2)
        #4
        self.Text_for_save4 = Label(self, text = u"儲值", font = 16, background = "PaleVioletRed2")
        self.Text_for_save4.grid(row=3, column=0)
        self.te4 = Entry(self)
        self.te4["width"] = 20 # Byte for 50 words
        self.te4.grid(row=3, column=1, columnspan=2)
        #5
        self.Text_for_save5 = Label(self, text = u"服務", font = 16, background =  "PaleVioletRed2")
        self.Text_for_save5.grid(row=4, column=0)
        
        self.OpMenu_var1 = StringVar()
        self.CheckVar1 = IntVar()
        
        self.OpMenu_var1.set(u"睫毛")
        self.sel1 = OptionMenu(self, self.OpMenu_var1, *self.mylist1, command = self.services1)
        self.sel1["width"] = 8 # Byte for 50 words
        self.sel1.grid(row=4, column=1, columnspan=2)
        
        self.CheckBox1 = Checkbutton(self, text = u"除毛", variable = self.CheckVar1, onvalue = 1, offvalue = 0, height=2, width = 5, background =  "pink")
        self.CheckBox1.grid(row=4, column=2, columnspan=3)  
        '''    
        #6
        self.Text_for_save6 = Label(self, text = "根數", font = 16, background =  "PaleVioletRed2")
        self.Text_for_save6.grid(row=5, column=0)
        self.te5 = Entry(self)
        self.te5["width"] = 25 # Byte for 50 words
        self.te5.grid(row=5, column=1, columnspan=3)
        '''
        #7
        self.Text_for_save7 = Label(self, text = "價格", font = 16, background =  "PaleVioletRed2")
        self.Text_for_save7.grid(row=5, column=0)
        self.te6 = Entry(self)
        self.te6["width"] = 20 # Byte for 50 words
        self.te6.grid(row=5, column=1, columnspan=2)
        
        #8
        self.Text_for_save8 = Label(self, text = "Line ", font = 16, background =  "PaleVioletRed2")
        self.Text_for_save8.grid(row=6, column=0)
        self.te7 = Entry(self)
        self.te7["width"] = 20 # Byte for 50 words
        self.te7.grid(row=6, column=1, columnspan=2)
        
        #9
        self.Text_for_save9 = Label(self, text = u"備註", font = 16, background =  "PaleVioletRed2")
        self.Text_for_save9.grid(row=7, column=0)
        self.te8 = Entry(self)
        self.te8["width"] = 20 # Byte for 50 words
        self.te8.grid(row=7, column=1, columnspan=2)
        
        #Output
        #self.printWord = Label(self)
        #self.printWord["background"] = "LightGoldenrod1"
        #self.printWord["text"] = "紀錄:"
        #self.printWord["font"] = 16
        #self.printWord.grid(row=4, column=0)
        self.out = Text(self, background = "azure", foreground = "black")
        self.out["width"] = 40
        self.out["height"] = 14 
        self.out.grid(row=8, column=0, columnspan=6)        
        
        #Botton
        #1
        self.save = Button(self, text = u"存檔", font = 16, background = "lemon chiffon", foreground = "black", relief = "groove")
        self.save["width"] = 6
        self.save.grid(row=9, column=2)
        self.save["command"] =  self.saveMethod
        #2
        self.exit = Button(self, text = u"離開", font = 16, background = "lemon chiffon", foreground = "black", relief = "groove")  
        self.exit["width"] = 6
        self.exit.grid(row=9, column=3)
        self.exit["command"] =  self.exitMethod
        #3
        self.get = Button(self, text = u"尋找", font = 16, background = "lemon chiffon", foreground = "black", relief = "groove")  
        self.get["width"] = 6
        self.get.grid(row=9, column=1)
        self.get["command"] =  self.getMethod
        #4
        self.get = Button(self, text = u"清除", font = 16, background = "lemon chiffon", foreground = "black", relief = "groove")  
        self.get["width"] = 6
        self.get.grid(row=9, column=0)
        self.get["command"] =  self.clean_all
     
    def clean_all(self):
        self.clean_flag()
        self.clean_values()

    def clean_flag(self):
                # Clean all flag
        self.row_get = False
        self.New_Member = False
        self.ReDo = False
        # clean 睫毛
        if self.OpMenu_var1.get() != u'睫毛':
            self.Text_info1.destroy()
            self.text1.destroy()
            self.Text_info2.destroy()
            self.text2.destroy()
            self.Text_info3.destroy()
            self.text3.destroy()
            self.Text_info4.destroy()
            self.text4.destroy()
            self.Text_info5.destroy()
            self.text5.destroy()
            self.Text_info6.destroy()
            self.text6.destroy()
            self.Check_EyeLashes_under.destroy()
        self.OpMenu_var1.set(u"睫毛")
        #Clean member in dictionary
        for i in self.Setup_Data.Customer_data.keys():
            self.Setup_Data.Customer_data[i] = ""
        for i in self.Setup_Data_Detail.Customer_data.keys():
            self.Setup_Data_Detail.Customer_data[i] = ""
        # Reload data
        self.Setup_Data = OBJofXL(self.path_in_use, self.page_in_use)
        self.Setup_Data_Detail = OBJofXL2(self.DTL_path_in_use, self.DTL_page_in_use)
        
        
    def clean_values(self):
        self.te1.delete(0, END)
        self.te2.delete(0, END)
        self.te3.delete(0, END)
        self.te4.delete(0, END)
        self.CheckBox1.deselect()
        #self.te5.delete(0, END)
        self.te6.delete(0, END)
        self.te6.delete(0, END)
        self.te7.delete(0, END)
        self.te8.delete(0, END)
        self.out.delete(0.0, END)  
        if self.OpMenu_var1.get() != u'睫毛':
            self.text1.delete(0, END)
            self.text2.delete(0, END)
            self.text3.delete(0, END)
            self.text4.delete(0, END)
            self.text5.delete(0, END)
            self.text6.delete(0, END)
            self.Check_EyeLashes_under.deselect()
        # Reload data    
        self.Setup_Data = OBJofXL(self.path_in_use, self.page_in_use)
        self.Setup_Data_Detail = OBJofXL2(self.DTL_path_in_use, self.DTL_page_in_use) 
        
    def saveMethod(self):
        Get_Name = self.te1.get()
        Get_Phone = self.te2.get()
        Get_birth = self.te3.get()
        print("New_Member:" + str(self.New_Member)+" " +"Row_Get:" + str(self.row_get))
        if Get_Name == "" or Get_Phone == "":
            self.Output_Infor(u"請先搜尋會員\n", True)
            return 0
        else:
            #有輸入會員
            #if self.OpMenu_var1.set("睫毛"):
            #    self.Output_Infor("請選擇服務項目", True)
            #    return 0
            
            if self.New_Member == False and self.row_get != False:
                #非新會員, 新增筆數
                #self.OpMenu_var1.set("睫毛")
                self.Output_Infor(u'儲存會員資料', True)
                Member_Info = self.get_customer_info()
                self.Setup_Data.Update_Member(self.row_get, Member_Info)
                if self.OpMenu_var1.get() != u'睫毛':
                    Member_EB_Info = self.get_customer_EB_info(self.row_get) 
                    self.Setup_Data_Detail.Add_new_member_data(Member_EB_Info)
                self.clean_flag()
                
            elif self.New_Member == True and self.row_get == False:
                #找不到會員且是新會員
                if Get_Name == "" or Get_Phone == "" or Get_birth == "":
                    self.Output_Infor(u"資料不完全,至少輸入姓名,電話,生日\n", True)
                    return 0
                self.Output_Infor(u'新會員存檔', True)
                New_Member_Info = self.get_customer_info()
                self.Setup_Data.Add_new_member_data(New_Member_Info)
                #Add new member to the max_row
                #self.row_get = self.Setup_Data.workseet.max_row
                # Save ok clean
                self.clean_flag()
                
            elif self.New_Member == False and self.row_get == False:
                #沒有搜尋直接存檔
                self.getMethod()
                self.saveMethod()
            else:
                return 0
            
                
        '''path = "./log.txt"
            existed = os.path.exists(path.strip())
            if existed:
                self.out.insert(INSERT, "File is existed.")
                self.te1.delete(0,100) # 2 Bytes for one word  
            else:
                self.out.insert(INSERT, "The word: " + Get_Name + " is saved")
                self.te1.delete(0,100) # 2 Bytes for one word    
                f = open(path, "w")
                f.write(Get_Name)
                f.close()'''
                        
    def exitMethod(self):
        self.quit() 

    def getMethod(self):
        
        self.Setup_Data = OBJofXL(self.path_in_use, self.page_in_use)
        self.Setup_Data_Detail = OBJofXL2(self.DTL_path_in_use, self.DTL_page_in_use)
        
        customer_n = self.te1.get()
        customer_c = self.te2.get()
        # Search by name or phone
        if customer_n == "" and customer_c == "": 
            self.Output_Infor(u"輸入名字或電話搜尋", True)
            return 0
        elif customer_n != "" or customer_c != "":
            self.row_get = self.Setup_Data.Serch_member(customer_n, u"姓名")
            if self.row_get == False: # If can't find data
                self.row_get = self.Setup_Data.Serch_member(customer_c, u"電話")
        
        if self.row_get != False:
            #find out
            self.Setup_Data.Get_Member(self.row_get)
            #show in text
            self.clean_values()
            self.te1.insert(INSERT, self.Setup_Data.Customer_data[u'姓名']) 
            self.te2.insert(INSERT, self.Setup_Data.Customer_data[u'電話']) 
            self.te3.insert(INSERT, str(self.Setup_Data.Customer_data[u'生日'])) 
            self.te7.insert(INSERT, str(self.Setup_Data.Customer_data[u'Line'])) 
            self.te8.insert(INSERT, str(self.Setup_Data.Customer_data[u'備註'])) 
            if self.OpMenu_var1.get() != u'睫毛':
                print(self.row_get)
                self.Setup_Data_Detail.Get_Member(self.row_get - 1) #Row -1 = real PSN
                self.text1.insert(INSERT, self.Setup_Data_Detail.Customer_data[u'翹度'])
                self.text2.insert(INSERT, self.Setup_Data_Detail.Customer_data[u'根數'])
                self.text3.insert(INSERT, self.Setup_Data_Detail.Customer_data[u'粗度'])
                self.text4.insert(INSERT, self.Setup_Data_Detail.Customer_data[u'款式'])
                self.text5.insert(INSERT, self.Setup_Data_Detail.Customer_data[u'長度'])
                self.text6.insert(INSERT, self.Setup_Data_Detail.Customer_data[u'內容說明'])
                if self.Setup_Data_Detail.Customer_data[u'下睫毛'] == "Y":
                    self.Check_EyeLashes_under.select()
                    
            #show in output
            for i in self.Customer_Info:
                if self.Setup_Data.Customer_data[i] != None:
                    self.Output_Infor(i + ": "+ str(self.Setup_Data.Customer_data[i]) + "\n", False)            
        else:
            self.Output_Infor(u"找不到, 新會員請直接存檔", True)
            self.New_Member = True

    def services1(self, val):
        
        '''
        Select service
        
        '''
        if val == u'重接':
            self.ReDo = True      
        else:
            self.ReDo = False
            
        self.Setup_Data.Customer_data[u'睫毛'] = True 
        # show detail information           
        self.createWidgets2()
        

    def Output_Infor(self, string, clean_all):
        
        '''
        Print out the output data,
        arg 1 is output string,
        arg 2 is clean before print flag
        '''
        
        if clean_all == True:
            self.out.delete(0.0, END)
            self.out.insert(INSERT, string)
        else:
            self.out.insert(INSERT, string)
            
    def get_customer_info(self):
        
        Member_Info = []
        Member_Info.append(self.te8.get())  #0 備註
        Member_Info.append(self.te1.get())  #1 姓名
        Member_Info.append(self.te3.get())  #2 生日
        Member_Info.append(self.Setup_Data_Detail.Customer_data[u"根數"]) #3 根數
        Member_Info.append(self.te2.get())  #4 電話
        Member_Info.append(self.OpMenu_var1.get()) #5 次數
        Member_Info.append(self.CheckVar1.get()) #6 除毛
        Member_Info.append(self.ReDo)      #7 重接日
        prepayment = self.get_prepayment()
        if prepayment == False:
            Member_Info.append(0) #8 預繳     
        else:
            Member_Info.append(prepayment) #8 預繳            
        Member_Info.append(self.te7.get()) #9 Line
        return Member_Info
    
    def get_customer_EB_info(self, row):
        Member_EB_Info = []
        Member_EB_Info.append(row-1) # Row -1 = real PSN
        Member_EB_Info.append(self.text1.get())  #1 翹度
        Member_EB_Info.append(self.text2.get())  #2 根數
        Member_EB_Info.append(self.text3.get())  #3 粗度
        Member_EB_Info.append(self.text4.get())  #4 款式
        Member_EB_Info.append(self.text5.get())  #5 長度
        Member_EB_Info.append(self.CheckEBV.get()) #6 次數
        Member_EB_Info.append(self.text6.get())  #7 內容
        Member_EB_Info.append(self.te4.get()) #8 儲值
        Member_EB_Info.append(self.te6.get()) #9 消費金額
        prepayment = self.get_prepayment()
        if prepayment == False:
            Member_EB_Info.append(0) #10 預繳     
        else:
            Member_EB_Info.append(prepayment)
        Member_EB_Info.append(self.te8.get()) #11 備註
        return Member_EB_Info
    
    def get_prepayment(self):
        
        # Get number when there is not None or empty
        if self.te4.get() != "" and self.te4.get() != None:
            prepayment = int(self.te4.get())
        else:
            prepayment = 0
        if self.te6.get() != "" and self.te6.get() != None:
            payment = int(self.te6.get())
        else:
            payment = 0
        if self.Setup_Data.Customer_data[u"預繳餘額"]!= "" and self.Setup_Data.Customer_data["預繳餘額"] != None:
            Old_Prepayment = int (self.Setup_Data.Customer_data[u"預繳餘額"])
        else:
            Old_Prepayment = 0
        #start to caculate    
        if self.New_Member == True:
            #預繳 = 預繳金額- 消費
            if prepayment != 0:
                if payment != 0:
                    prepayment = prepayment - payment
                else:
                    return prepayment
            else:
                return False # There is no prepayment
        else:
            print("modify")
            #非新會員 新預繳+ 舊預繳 - 消費
            #print("old" + Old_Prepayment)
            if Old_Prepayment != 0:
                if prepayment != 0:
                    if payment != 0:
                        prepayment = prepayment + Old_Prepayment - payment
                    else:
                        prepayment = prepayment + Old_Prepayment
                else:
                    if payment != 0:
                        prepayment = Old_Prepayment - payment
                    else:
                        prepayment = Old_Prepayment
            else:
                if prepayment != 0:
                    if payment != 0:
                        prepayment = prepayment - payment
                    else:
                        return prepayment
                else:
                    return False # There is no prepayment
        return prepayment
            
    
    def createWidgets2(self):
        #input Text
        #1       
        self.Text_info1 = Label(self, text = u"翹度", font = 16, background = "PaleVioletRed2")
        self.Text_info1.grid(row=0, column=4)
        #input Entry
        self.text1 = Entry(self)
        self.text1["width"] = 15 # Byte for 50 words
        self.text1.grid(row=0, column=5, columnspan=2)
        #2
        self.Text_info2 = Label(self, text = u"根數", font = 16, background = "PaleVioletRed2")
        self.Text_info2.grid(row=1, column=4)
        #input Entry
        self.text2 = Entry(self)
        self.text2["width"] = 15 # Byte for 50 words
        self.text2.grid(row=1, column=5, columnspan=2)
        #3
        self.Text_info3 = Label(self, text = u"粗度", font = 16, background = "PaleVioletRed2")
        self.Text_info3.grid(row=2, column=4)
        #input Entry
        self.text3 = Entry(self)
        self.text3["width"] = 15 # Byte for 50 words
        self.text3.grid(row=2, column=5, columnspan=2)
        #4
        self.Text_info4 = Label(self, text = u"款式", font = 16, background = "PaleVioletRed2")
        self.Text_info4.grid(row=3, column=4)
        #input Entry
        self.text4 = Entry(self)
        self.text4["width"] = 15 # Byte for 50 words
        self.text4.grid(row=3, column=5, columnspan=2)
        #5
        self.Text_info5 = Label(self, text = u"長度", font = 16, background = "PaleVioletRed2")
        self.Text_info5.grid(row=4, column=4)
        #input Entry
        self.text5 = Entry(self)
        self.text5["width"] = 15 # Byte for 50 words
        self.text5.grid(row=4, column=5, columnspan=2)
        #6
        self.Text_info6 = Label(self, text = u"內容", font = 16, background = "PaleVioletRed2")
        self.Text_info6.grid(row=5, column=4)
        #input Entry    
        self.text6 = Entry(self)
        self.text6["width"] = 15 # Byte for 50 words
        self.text6.grid(row=5, column=5, columnspan=2)
        #7
        self.CheckEBV = IntVar()
        self.Check_EyeLashes_under = Checkbutton(self, text = u"下睫毛", variable = self.CheckEBV, onvalue = 1, offvalue = 0, height=2, width = 5, background =  "pink")
        self.Check_EyeLashes_under.grid(row=6, column=6) 
        

if __name__ == '__main__':
    root = Tk() # Tk Object
    root.title("Ann")
    app = DemoGUI(master=root)
    #start the program
    app.mainloop()