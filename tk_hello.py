#!/usr/bin/python
# -*- coding: UTF-8 -*-
 
#import tkinter
#top = tkinter.Tk()
# 进入消息循环
#top.mainloop()

import pandas as pd

from tkinter import *
from tkinter import messagebox
import datetime


def date_convert(dates): #定义转化日期戳的函数,dates为日期戳
    delta=datetime.timedelta(days=dates)
    today=datetime.datetime.strptime('1899-12-30','%Y-%m-%d')+delta#将1899-12-30转化为可以计算的时间格式并加上要转化的日期戳
    return datetime.datetime.strftime(today,'%Y-%m-%d')#制定输出日期的格式

def getYearMonth(my_date):
    # print('Type: {}'.format(type(s)))
    if isinstance(my_date, str):
        #new_date = datetime.strptime(my_date, "%Y-%m-%d")
        #dateinfo = '{}-{}'.format((new_date.year), new_date.month)
        if '-' not in my_date:
            dateinfo = 'NAN'
        else:
            dateinfo = my_date.split("-")[1]+"-"+my_date.split("-")[2]
    elif isinstance(my_date, datetime.datetime):
        dateinfo = '{}-{}'.format((my_date.year), my_date.month)
    else:
        #raise Exception('date format error please check {}'.format(my_date))
        dateinfo = 'NAN'
    print(dateinfo)
    return dateinfo
#初始化Tk()
myWindow = Tk()
v=IntVar()
#设置标题
myWindow.geometry('380x250')
myWindow.title('商务数据处理工具')
#标签控件布局
Label(myWindow, text="input path").grid(row=0)
Label(myWindow, text="output path").grid(row=1)
#Entry控件布局
entry1=Entry(myWindow)
entry2=Entry(myWindow)
entry1.grid(row=0, column=1)
entry2.grid(row=1, column=1)

def format_number(x):
    value = re.compile(r'^\s*[-+]*[0-9]+\.*[0-9]*\s*$')
    if value.match(str(x)): #不是数字
        return x
    else:
        # print('x2:>>>', str(x))
        return 0

def chart_maker(writer, sheetname, charttype, sheetdata):
    workbook = writer.book
    worksheet = writer.sheets[sheetname]
    # Create a chart object.
    chart = workbook.add_chart({'type': charttype})
    length2 = len(sheetdata.index) + 1
        
    categories = '={}!$A$2:$A${}'.format(sheetname, length2)
    values = '={}!$B$2:$B${}'.format(sheetname, length2)
    # Configure the series of the chart from the dataframe data.
    chart.add_series({
        'categories': categories,
        'values':     values,
        'name':       'chart data',
    })
    if charttype == 'pie':
        chart.set_style(10)
    else:
        chart.set_style(11)
    # Insert the chart into the worksheet.
    worksheet.insert_chart('D2', chart)


def anaysis_month_data():

    choice = v.get()
    if choice!= 0:
        messagebox.showinfo("error", "error choose {}".format(choice))
        return 0

    #excel_file = r'C:\Users\wuh17\Downloads\sourcedata\iot_month_product_data.xlsx'
    excel_file = entry1.get()
    com_data = pd.read_excel(excel_file)

    key_list = []
    name_list = []
    num_atv_list = []    
    num_update_list = []

    #busssis = com_data[r'商务']
    #ourbus = list(dict.fromkeys(busssis))
    #for name in ourbus:
      #  busdata = com_data.loc[com_data[r'商务']== name]
       # keylist = []
        #for index, row in busdata.iterrows():
         #   keyvalue = row[r'项目'] + row[r'品牌商']
       # break
    for index, row in com_data.iterrows():
        key_list.append(str(row[r'项目']) + str(row[r'品牌商']) + str(row[r'商务']) + str(row[r'ID']))
    com_data['prokey'] = key_list

    print(len(key_list))

    single_pro_name = list(dict.fromkeys(key_list))
    for singlekey in single_pro_name:
        onetype = com_data.loc[com_data[r'prokey']== singlekey]
        num_atv = 0
        num_update = 0
        for index, row in onetype.iterrows():
            name_worker = row[r'商务']
            num_atv += (row[r'激活数量'])    
            num_update += (row[r'升级数量'])
        name_list.append(name_worker)
        num_atv_list.append(num_atv)
        num_update_list.append(num_update)
    #com_data[r'累计激活'] = num_atv_list
    #com_data[r'累计升级'] = num_update_list

    # Define a dictionary containing Students data 
    data = {'商务': name_list,
            'pro_key': single_pro_name, 
            r'累计激活': num_atv_list, 
            r'累计升级': num_update_list} 
       
       
    # Convert the dictionary into DataFrame 
    df = pd.DataFrame(data)
    df.to_excel("output.xlsx")
    messagebox.showinfo("data generate", "success!")

def bussis_analysis():
    print('hello')


def project_analysis():
    excel_file = entry1.get()
    #xl = pd.ExcelFile(r'C:\Users\wuh17\Downloads\sourcedata\calucate_project\Projects--2020.xlsx')
    all_data =  pd.ExcelFile(excel_file)
    sheet_names = all_data.sheet_names  # see all sheet names

    for name in sheet_names:
        if '20' in name:
            # read a specific sheet to DataFrame
            databuff = all_data.parse(name)

    #print(databuff.index)
    #df.groupby('state')['name'].nunique().plot(kind='bar')
    # gk = databuff.groupby(r'商务').size().reset_index(name='counts')
    # people_data = gk.sort_values('counts', ascending=False)


    #databuff['YearMonth']= databuff[r'新建项目时间'].apply(lambda x:x.date())
    databuff['YearMonth']= databuff[r'新建项目时间'].apply(lambda x: getYearMonth(x))
    people_data = databuff.groupby([r'商务', 'YearMonth']).size().reset_index(name='counts')
    #print(databuff['YearMonth'])

    
    location_raw = databuff.groupby([r'项目地区', 'YearMonth']).size().reset_index(name='counts')
    location_data =  location_raw.sort_values('counts', ascending=False)

    system_raw = databuff.groupby([r'系统', 'YearMonth']).size().reset_index(name='counts')
    system_data = system_raw.sort_values('counts', ascending=False)

    project_state_data = databuff.groupby([r'项目状态', 'YearMonth']).size().reset_index(name='counts')
    #project_state_data = project_state_raw.sort_values('counts', ascending=False)
    
    project_state_data = project_state_data.replace(r'新建项目', r'市场机会')
    project_state_data = project_state_data.replace(r'需求分析', r'项目立项')
    project_state_data = project_state_data.replace([r'项目开发',r'集成调试',r'内部测试'], r'项目开发')
    project_state_data = project_state_data.replace([r'客户验收', r'交付完成'], r'客户交付')
    project_state_data = project_state_data.groupby([r'项目状态', 'YearMonth']).sum()
    print(project_state_data)
    #project_state_data = project_state_raw.sort_values('counts', ascending=False)
    
    Contract_signing_raw = databuff.groupby([r'合同是否签订', 'YearMonth']).size().reset_index(name='counts')
    Contract_signing_data = Contract_signing_raw.sort_values('counts', ascending=False)

    chip_maker_raw = databuff.groupby([r'芯片厂商', 'YearMonth']).size().reset_index(name='counts')
    chip_maker_data = chip_maker_raw.sort_values('counts', ascending=False)
    
    with pd.ExcelWriter('Iot_project_statistics.xlsx') as writer:  
        people_data.to_excel(writer, sheet_name=r'项目数', index=False)
        location_data.to_excel(writer, sheet_name=r'地区统计', index=False)
        system_data.to_excel(writer, sheet_name=r'产品平台与行业', index=False)
        project_state_data.to_excel(writer, sheet_name=r'项目状态')
        Contract_signing_data.to_excel(writer, sheet_name=r'商务数据', index=False)
        chip_maker_data.to_excel(writer, sheet_name=r'芯片厂商', index=False)
        
        # chart_maker(writer, r'项目数', 'column',people_data)               
        # chart_maker(writer, r'地区统计', 'pie',location_data)
        #chart_maker(writer, r'产品平台与行业', 'pie',system_data)               
        #chart_maker(writer, r'项目状态', 'pie',project_state_data)
        #chart_maker(writer, r'商务数据', 'pie',Contract_signing_data)               
       # chart_maker(writer, r'芯片厂商', 'column',chip_maker_data)
        
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()        
    messagebox.showinfo("data generate", r" success!")

def agreement_analysis():
    excel_file = entry1.get()
    #xl = pd.ExcelFile(r'C:\Users\wuh17\Downloads\sourcedata\agreement\raw.xlsx')
    all_data =  pd.ExcelFile(excel_file)
    databuff = all_data.parse(r'合同')
    # iterating the columns 
    for col in databuff.columns: 
        if 'BD' in col:
            name = col
            break
    #agreement_data = databuff['BD'].str.contains('2020', re.IGNORECASE).groupby(databuff[r'合同归属']).sum()
    #df = df.groupby(['category']).filter(lambda x: len(x) >= 5)
   # agreement_raw = databuff.groupby([r'签署月份']).filter(lambda x: '2020' in str(x)).reset_index()
    databuff[r'先收费'] = databuff[r'先收费'].apply(format_number)
    agreement_pro_count = databuff.groupby([name,r'签署月份']).size().reset_index(name='counts')
    agreement_data_money = databuff.groupby([name,r'签署月份'])[r'先收费'].sum()

    with pd.ExcelWriter('agreement_analysis.xlsx') as writer:  
        agreement_pro_count.to_excel(writer, sheet_name=r'合同数统计', index=False)
        agreement_data_money.to_excel(writer, sheet_name=r'合同金额统计')
        writer.save()
    messagebox.showinfo("data generate", "success!")    
    #print(databuff['YearMonth'])    

 
def data_analysis_dispatch():
    value = v.get()
    fun_list = [anaysis_month_data, bussis_analysis, project_analysis, agreement_analysis]
    print(value)
    (fun_list[value])()
    
Label(myWindow,text='选择一种需要处理的数据').grid(row=3, column=1, sticky=W, padx=5, pady=5)
data_options=[(r'设备激活&升级数',0),(r'商务对账单',1),(r'物联网项目数统计',2),('物联网合同数以及金额统计',3)]
for lan,num in data_options:
    Radiobutton(myWindow, text=lan, value=num,  variable=v).grid(row=num+4, column=1, sticky=W, padx=5, pady=5)


#Quit按钮退出；Run按钮打印计算结果
Button(myWindow, text='Quit', command=myWindow.quit).grid(row=num+5, column=4,sticky=W, padx=5, pady=5)
Button(myWindow, text='Run', command=data_analysis_dispatch).grid(row=num+5, column=0, sticky=W, padx=5, pady=5)



#进入消息循环
myWindow.mainloop()







