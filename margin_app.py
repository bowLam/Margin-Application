# -*- coding: utf-8 -*-
"""
Created on Mon Mar  7 21:04:29 2022

@author: Bowie Lam 
"""
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui
import numpy as np
import pandas as pd
import requests
import time
import ctypes 
import os
import webbrowser
import datetime
from tkinter import *
from tkinter import filedialog
from datetime import date
import math



list2 = ['cat', 'bat', 'mat', 'cat', 
         'get', 'cat', 'sat', 'pet']
list2[0:]

driver_path = r"\\192.168.35.1\fs\Derek Lung\chromedriver.exe"
driver_path
#all path 
input_path = "//input[@id='username']"
password_path = "//input[@id='password']"
clickbutton="//input[@id='btnLogin']"
prev_bal= "//body/table[contains(@class, 'subtitleGrey')][1]/tbody/tr[9]/td[contains(@class, 'subtitleBlue')]"
fund_tra ="//table[contains(@class, 'subtitleGrey')][1]/tbody/tr[10]/td[contains(@class, 'subtitleBlue')]"
cash= "//table[contains(@class, 'subtitleGrey')][1]/tbody/tr[11]/td[contains(@class, 'subtitleBlue')]"
buy="//table[contains(@class, 'subtitleGrey')][1]/tbody/tr[12]/td[contains(@class, 'subtitleBlue')]"
sell="//table[contains(@class, 'subtitleGrey')][1]/tbody/tr[13]/td[contains(@class, 'subtitleBlue')]"
hk_acc_bal="//table[contains(@class, 'subtitleGrey')][2]/tbody/tr[16]/td[contains(@class, 'Title')]"
real_tim="//div[@id='quote-header-info']/div[contains(@class, 'My(6px) Pos(r) smartphone_Mt(6px) W(100%)')]/div[contains(@class, 'D(ib) Va(m) Maw(65%) Ov(h)')]/div[contains(@class, 'D(ib) Mend(20px)')]/fin-streamer[contains(@class, 'Fw(b) Fz(36px) Mb(-4px) D(ib)')]"

# List of Tuples
# empoyees = [('jack', 34, 'Sydney') ,
#            ('Riti', 31, 'Delhi') ,
#            ('Aadi', 16, 'New York') ,
#            ('Mohit', 32,'Delhi') ,
#             ]
# # Create a DataFrame object
# empDfObj = pd.DataFrame(empoyees, columns=['Name', 'Age', 'City'])
# for i in range(len(empDfObj)):
#     if empDfObj.iloc[i,1] > 30:
#         print(empDfObj.iloc[i,0])

# folder_path = os.getcwd()
# folder_path
# driver_path = folder_path + "\\chromedriver"
# driver_path

pshk=Tk()
pshk.title("Margin Call Program")
pshk.geometry("500x150")

#testing accounts:
# m566836,m585071,m598505,m501357,m562875,m580065,m501380,m595115,m584651,m573240,m324388,m324271
#define x-path
client_name = "//html/body/table[contains(@class, 'subtitleGrey')][1]/tbody/tr[2]/td[contains(@class, 'subtitleBlue')]"

def button_click():
    record = str(acc_num.get())
    records = record.split(",")
    global editor
    global f_name_editor
    global l_name_editor
    global address_editor
    global city_editor
    global p_editor
    global edit_btn
    #looping through: 
    for i in records:
        #Path of webdriver
        driver = webdriver.Chrome(executable_path=driver_path)
        driver.get("http://192.168.5.86/intranet/index.asp")
        elem= driver.find_element_by_xpath(input_path)
        elem.send_keys("ivanlau")
        elem1= driver.find_element_by_xpath(password_path)
        elem1.send_keys("ivanlau19")
        elem3 =  driver.find_element_by_xpath(clickbutton)
        elem3.click()
        driver.get("http://192.168.5.86/intranet/FS/FSClientNew.asp?func=View&Accode=" + i)
        elem_name=driver.find_element_by_xpath(client_name).text
        prev_bala=driver.find_element_by_xpath(prev_bal).text.replace(",","")
        fund_traa=driver.find_element_by_xpath(fund_tra).text.replace(",","")
        casha=driver.find_element_by_xpath(cash).text.replace(",","")
        buya=driver.find_element_by_xpath(buy).text.replace(",","")
        sella=driver.find_element_by_xpath(sell).text.replace(",","")
        hk_acc_bala=driver.find_element_by_xpath(hk_acc_bal).text.replace(",","")
        # print(fund_traa)
        # print(call_amount)
        driver.get("http://192.168.5.86/intranet/FS/FSClientDetail.asp?func=View&Accode="+i+"&Data=Position")
        post_stock="//html/body/table[contains(@class, 'normal')][1]/tbody/tr/td[1]"
        quant="//html/body/table[contains(@class, 'normal')]/tbody/tr/td[5]"
        ratio="//html/body/table[contains(@class, 'normal')]/tbody/tr/td[8]"
        code=[]
        quan=[]
        rat=[]
        for e in driver.find_elements_by_xpath(post_stock)[1:-1]:
            code.append(e.text)
        for r in driver.find_elements_by_xpath(quant)[1:-1]:
            quan.append(float(r.text.replace(",","")))
        for t in driver.find_elements_by_xpath(ratio)[1:-1]:
            rat.append(float(t.text.replace("%","")))     
        data={
            'Code':code,
            'Quan':quan,
            'Ratio':rat
            }
        df= pd.DataFrame(data)
        df= df.sort_values(by="Ratio", ascending=True)
        real_time=[]
        for v in df["Code"].to_list():
            driver.get("https://hk.finance.yahoo.com/quote/"+v+"?p="+v+"&.tsrc=fin-srch")
            real_time.append(float(driver.find_element_by_xpath(real_tim).text.replace(",","")))
        df["Price"]=real_time
        df["Mark_Val"]=df["Quan"]*df["Price"]*(df["Ratio"]/100)
        margin_value=round(df["Mark_Val"].sum()*7.8,2)
        if float(hk_acc_bala) > 0:
            call_amount= float(prev_bala)+float(fund_traa)+float(casha)+float(hk_acc_bala)+margin_value
            df["C_Amount"]=call_amount
            print(call_amount)
        else:
            call_amount= float(prev_bala)+float(fund_traa)+float(casha)+margin_value  
            df["C_Amount"]=call_amount
            print(call_amount)
        print(df)

        editor = Tk()
        editor.title(i)
        editor.geometry("400x300")
        # 	Create Global Variables for text box names

        # global state_editor
        # global zipcode_editor
        # Create Text Boxes
        f_name_editor = Entry(editor, width=30)
        f_name_editor.grid(row=0, column=1, padx=20, pady=(10, 0))
        l_name_editor = Entry(editor, width=30)
        l_name_editor.grid(row=1, column=1)
        address_editor = Entry(editor, width=30)
        address_editor.grid(row=2, column=1)
        city_editor = Entry(editor, width=30)
        city_editor.grid(row=3, column=1)
        p_editor = Entry(editor, width=30)
        p_editor.grid(row=4, column=1)        
        state_editor = Entry(editor, width=30)
        state_editor.grid(row=5, column=1)
        	
        # Create Text Box Labels
        f_name_label = Label(editor, text="Prev Bal")
        f_name_label.grid(row=0, column=0, pady=(10, 0))
        l_name_label = Label(editor, text="Fund Transfer")
        l_name_label.grid(row=1, column=0)
        address_label = Label(editor, text="Cash")
        address_label.grid(row=2, column=0)
        city_label = Label(editor, text="HK Stock Balance")
        city_label.grid(row=3, column=0)
        p_label = Label(editor, text="Real time Margin Value")
        p_label.grid(row=4, column=0)
        state_label = Label(editor, text="Margin Call in HKD")
        state_label.grid(row=5, column=0)
        f_name_editor.insert(0, prev_bala)
        l_name_editor.insert(0, fund_traa)
        address_editor.insert(0, casha)
        city_editor.insert(0, hk_acc_bala)
        p_editor.insert(0, margin_value)
        state_editor.insert(0, round(call_amount,2))
        edit_btn = Button(editor, text="Close", command=editor.destroy)
        edit_btn.grid(row=13, column=0, columnspan=2, pady=10, padx=10, ipadx=145)
        if call_amount<0:
            list_label = Label(editor, text="Format: Stock Price Shares-Cut")
            list_label.grid(row=6, column=0, pady=(10, 0))
            list1= Listbox(editor, height=6, width=30)
            list1.grid(row=6, column=1, rowspan=6)
            sbl=Scrollbar(editor)
            sbl.grid(row=6,column=2, rowspan=6)
            list1.configure(yscrollcommand=sbl.set)
            sbl.configure(command=list1.yview)
            org_sell_cash=0
            for i in range(len(df)):
                tem_mkt=df.iloc[i, 1]*df.iloc[i,3]
                if tem_mkt >= abs(df.iloc[i, 5])/(1-df.iloc[i,2]/100):
                    shares_cut=round(float((df.iloc[i, 5]/(1-df.iloc[i,2]/100))/(df.iloc[i,3]*7.8)),2)
                # if shares_cut <= df.iloc[i,1]:
                # define list box
                    list1.insert("end",str(df.iloc[i, 0])+" "+str(df.iloc[i,3])+" "+str(shares_cut))
                    # =============================================================================
                    # attach scrollbar to the list
                    break
                elif tem_mkt < abs(df.iloc[i, 5])/(1-df.iloc[i,2]/100):
                    list1.insert("end",str(df.iloc[i, 0])+" "+str(df.iloc[i,3])+" "+str(df.iloc[i,1]))
                    cash_come=tem_mkt*7.8
                    org_sell_cash+=cash_come
                    print(org_sell_cash)
                    df.iloc[i,4]=0
                    # df= df.iloc[i+1:,:]
                    margin_value=round(df["Mark_Val"].sum()*7.8,2)
                    if float(hk_acc_bala) > 0:
                        call_amount= float(prev_bala)+float(fund_traa)+float(org_sell_cash)+float(hk_acc_bala)+margin_value+float(casha)
                        df["C_Amount"]=call_amount
                        print(call_amount)
                        print(df)
                    else:
                        call_amount= float(prev_bala)+float(fund_traa)+float(org_sell_cash)+margin_value+float(casha)
                        df["C_Amount"]=call_amount
                        print(call_amount)
                        print(df)
                    

        driver.close()                    
                    # money_cut=shares_cut*df.iloc[i,3]*7.8
                    # print(money_cut)
                # if float(hk_acc_bala) > 0:
                #     df=df.iloc[i+1:,:]
                #     print(df)
                #     call_amount= float(prev_bala)+float(fund_traa)+float(casha)+float(hk_acc_bala)
                #     +round(df["Mark_Val"].sum()*7.8,2)+abs(money_cut)+first_force_sell
                #     print(call_amount)
                #     df["C_Amount"]=call_amount
                #     # df= df.iloc[i:,:]
                #     # first_force_sell+=tem_mkt
                #     break
                # else:
                #     df=df.iloc[i+1:,:]
                #     print(df)
                #     call_amount= float(prev_bala)+float(fund_traa)+float(casha)+round(df["Mark_Val"].sum()*7.8,2)
                #     +abs(money_cut)+first_force_sell
                #     print(call_amount)
                #     df["C_Amount"]=call_amount
                #     break

                
                # else:
                #     first_force_sell=0
                #     if float(hk_acc_bala) > 0:
                #         while call_amount<0:
                #             call_amount= float(prev_bala)+float(fund_traa)+float(casha)+float(hk_acc_bala)
                #             +margin_value+tem_mkt+first_force_sell
                #             df["C_Amount"]=call_amount
                #             df= df.iloc[i:,:]
                #             first_force_sell+=tem_mkt
                        
                #     else:
                #         call_amount= float(prev_bala)+float(fund_traa)+float(casha)+margin_value+first_force_sell
                #         df["C_Amount"]=call_amount
                    
                # break

     
    
#Create excel path searching 
today= date.today()
def open():
    global editor
    global f_name_editor
    global l_name_editor
    global address_editor
    global city_editor
    global p_editor
    global edit_btn
    pshk.filename = filedialog.askopenfilename(initialdir=r"\\192.168.35.1\fs\Felix Kwan", 
                                               title = "Select a File",
                                               filetypes=(("Excel Files", "*.xlsx"),))
    # askopenfilename(filetypes=(("Mp3 Files", "*.mp3"),))
    acc_num_ex.insert(0,pshk.filename)
    df2=pd.read_excel(pshk.filename)
    acc_n=df2[df2['Des'].str.contains(str(today.strftime("%d/%m/%Y")),na=False)]["CODE"]
    for i in acc_n.values:
        #Path of webdriver
        #Path of webdriver
        driver = webdriver.Chrome(executable_path=driver_path)
        driver.get("http://192.168.5.86/intranet/index.asp")
        elem= driver.find_element_by_xpath(input_path)
        elem.send_keys("ivanlau")
        elem1= driver.find_element_by_xpath(password_path)
        elem1.send_keys("ivanlau19")
        elem3 =  driver.find_element_by_xpath(clickbutton)
        elem3.click()
        driver.get("http://192.168.5.86/intranet/FS/FSClientNew.asp?func=View&Accode=" + i)
        elem_name=driver.find_element_by_xpath(client_name).text
        prev_bala=driver.find_element_by_xpath(prev_bal).text.replace(",","")
        fund_traa=driver.find_element_by_xpath(fund_tra).text.replace(",","")
        casha=driver.find_element_by_xpath(cash).text.replace(",","")
        buya=driver.find_element_by_xpath(buy).text.replace(",","")
        sella=driver.find_element_by_xpath(sell).text.replace(",","")
        hk_acc_bala=driver.find_element_by_xpath(hk_acc_bal).text.replace(",","")
        # print(fund_traa)
        # print(call_amount)
        driver.get("http://192.168.5.86/intranet/FS/FSClientDetail.asp?func=View&Accode="+i+"&Data=Position")
        post_stock="//html/body/table[contains(@class, 'normal')][1]/tbody/tr/td[1]"
        quant="//html/body/table[contains(@class, 'normal')]/tbody/tr/td[5]"
        ratio="//html/body/table[contains(@class, 'normal')]/tbody/tr/td[8]"
        code=[]
        quan=[]
        rat=[]
        for e in driver.find_elements_by_xpath(post_stock)[1:-1]:
            code.append(e.text)
        for r in driver.find_elements_by_xpath(quant)[1:-1]:
            quan.append(float(r.text.replace(",","")))
        for t in driver.find_elements_by_xpath(ratio)[1:-1]:
            rat.append(float(t.text.replace("%","")))     
        data={
            'Code':code,
            'Quan':quan,
            'Ratio':rat
            }
        df= pd.DataFrame(data)
        df= df.sort_values(by="Ratio", ascending=True)
        real_time=[]
        for v in df["Code"].to_list():
            driver.get("https://hk.finance.yahoo.com/quote/"+v+"?p="+v+"&.tsrc=fin-srch")
            real_time.append(float(driver.find_element_by_xpath(real_tim).text.replace(",","")))
        df["Price"]=real_time
        df["Mark_Val"]=df["Quan"]*df["Price"]*(df["Ratio"]/100)
        margin_value=round(df["Mark_Val"].sum()*7.8,2)
        if float(hk_acc_bala) > 0:
            call_amount= float(prev_bala)+float(fund_traa)+float(casha)+float(hk_acc_bala)+margin_value
            df["C_Amount"]=call_amount
            print(call_amount)
        else:
            call_amount= float(prev_bala)+float(fund_traa)+float(casha)+margin_value  
            df["C_Amount"]=call_amount
            print(call_amount)
        print(df)

        editor = Tk()
        editor.title(i)
        editor.geometry("400x300")
        # 	Create Global Variables for text box names

        # global state_editor
        # global zipcode_editor
        # Create Text Boxes
        f_name_editor = Entry(editor, width=30)
        f_name_editor.grid(row=0, column=1, padx=20, pady=(10, 0))
        l_name_editor = Entry(editor, width=30)
        l_name_editor.grid(row=1, column=1)
        address_editor = Entry(editor, width=30)
        address_editor.grid(row=2, column=1)
        city_editor = Entry(editor, width=30)
        city_editor.grid(row=3, column=1)
        p_editor = Entry(editor, width=30)
        p_editor.grid(row=4, column=1)        
        state_editor = Entry(editor, width=30)
        state_editor.grid(row=5, column=1)
        	
        # Create Text Box Labels
        f_name_label = Label(editor, text="Prev Bal")
        f_name_label.grid(row=0, column=0, pady=(10, 0))
        l_name_label = Label(editor, text="Fund Transfer")
        l_name_label.grid(row=1, column=0)
        address_label = Label(editor, text="Cash")
        address_label.grid(row=2, column=0)
        city_label = Label(editor, text="HK Stock Balance")
        city_label.grid(row=3, column=0)
        p_label = Label(editor, text="Real time Margin Value")
        p_label.grid(row=4, column=0)
        state_label = Label(editor, text="Margin Call in HKD")
        state_label.grid(row=5, column=0)
        f_name_editor.insert(0, prev_bala)
        l_name_editor.insert(0, fund_traa)
        address_editor.insert(0, casha)
        city_editor.insert(0, hk_acc_bala)
        p_editor.insert(0, margin_value)
        state_editor.insert(0, round(call_amount,2))
        edit_btn = Button(editor, text="Close", command=editor.destroy)
        edit_btn.grid(row=13, column=0, columnspan=2, pady=10, padx=10, ipadx=145)
        if call_amount<0:
            list_label = Label(editor, text="Format: Stock Price Shares-Cut")
            list_label.grid(row=6, column=0, pady=(10, 0))
            list1= Listbox(editor, height=6, width=30)
            list1.grid(row=6, column=1, rowspan=6)
            sbl=Scrollbar(editor)
            sbl.grid(row=6,column=2, rowspan=6)
            list1.configure(yscrollcommand=sbl.set)
            sbl.configure(command=list1.yview)
            org_sell_cash=0
            for i in range(len(df)):
                tem_mkt=df.iloc[i, 1]*df.iloc[i,3]
                if tem_mkt >= abs(df.iloc[i, 5])/(1-df.iloc[i,2]/100):
                    shares_cut=round(float((df.iloc[i, 5]/(1-df.iloc[i,2]/100))/(df.iloc[i,3]*7.8)),2)
                # if shares_cut <= df.iloc[i,1]:
                # define list box
                    list1.insert("end",str(df.iloc[i, 0])+" "+str(df.iloc[i,3])+" "+str(shares_cut))
                    # =============================================================================
                    # attach scrollbar to the list
                    break
                elif tem_mkt < abs(df.iloc[i, 5])/(1-df.iloc[i,2]/100):
                    list1.insert("end",str(df.iloc[i, 0])+" "+str(df.iloc[i,3])+" "+str(df.iloc[i,1]))
                    cash_come=tem_mkt*7.8
                    org_sell_cash+=cash_come
                    print(org_sell_cash)
                    df.iloc[i,4]=0
                    # df= df.iloc[i+1:,:]
                    margin_value=round(df["Mark_Val"].sum()*7.8,2)
                    if float(hk_acc_bala) > 0:
                        call_amount= float(prev_bala)+float(fund_traa)+float(org_sell_cash)+float(hk_acc_bala)+margin_value+float(casha)
                        df["C_Amount"]=call_amount
                        print(call_amount)
                        print(df)
                    else:
                        call_amount= float(prev_bala)+float(fund_traa)+float(org_sell_cash)+margin_value+float(casha)
                        df["C_Amount"]=call_amount
                        print(call_amount)
                        print(df)
                    

        driver.close()    

#define label
l1= Label(pshk, text="Account Number (add , if > 1 acc)")
l1.grid(row=0, column=0)
l2= Label(pshk, text="'PSHK_FSST_TOP100_ddmmyyyy_Amt'")
l2.grid(row=1, column=0)

#define input
# acc_text= StringVar()
acc_num = Entry(pshk, width=20)
acc_num.grid(row=0, column=1)
acc_num_ex = Entry(pshk, width=20)
acc_num_ex.grid(row=1, column=1)

#define button
search_btn = Button(pshk, text="Search", command=button_click, bg="red")
search_btn.grid(row=0, column=2, columnspan=3, pady=10, padx=10)
my_btn = Button(pshk, text="Eat Excel File", command=open, bg="orange")
my_btn.grid(row=1, column=2, columnspan=3, pady=10, padx=10)
close_btn = Button(pshk, text="Close", command=pshk.destroy,bg="yellow")
close_btn.grid(row=3, column=2, columnspan=2, pady=10, padx=10)


pshk.mainloop()

