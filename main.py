#!/usr/bin/env python
# coding: utf-8

# In[50]:

from pathlib import Path
import pandas as pd
import win32com.client as win32
from pathlib import Path
import os
import datetime
from win10toast import ToastNotifier
import six
import appdirs
import packaging.requirements


# In[51]:


# assign File variables
Excel_file_path = Path.cwd() / "US80 Roscoe.xlsx"
Email_file_Path = Path.cwd() / "Email_Python New US80.xlsx"
ATTACHMENT_DIR = Path.cwd() / "US80_Attachments"
ATTACHMENT_DIR.mkdir(exist_ok=True)


# In[52]:


Email_file_Path


# In[53]:


Excel_file_path


# In[54]:


# Load financial data into dataframe
data = pd.read_excel(Excel_file_path)


# In[55]:


data['Delivery date']=data['Delivery date'].dt.strftime("%d-%m-%Y")


# In[56]:


data.head()


# In[57]:


column_name = data.columns[12]
unique_values = data[column_name].unique()


# In[58]:


len(unique_values)


# In[59]:


# Load email distribution list into dataframe
email_list = pd.read_excel(Email_file_Path)
email_list.shape


# In[60]:


unique_ven = pd.DataFrame(unique_values,columns=[column_name])


# In[61]:


Final_mail_list = pd.DataFrame()
for u in unique_values:
    df=email_list[email_list[column_name].str[:]==u]
    Final_mail_list = Final_mail_list.append(df)


# In[62]:


Final_mail_list.shape


# In[63]:


# Iterate over email distribution list & send emails via Outlook App
outlook = win32.Dispatch("outlook.application")
for index, row in Final_mail_list.iterrows():
    mail = outlook.CreateItem(0)
    mail.To = row["List"]
    mail.Subject = row['Subject']
    mail.Body = row['mail body']
    attachment_path = str(ATTACHMENT_DIR / f"US80_Roscoe_Forecast_{row[column_name]}.xlsx")
    mail.Attachments.Add(Source=attachment_path)
    mail.Send()


# In[ ]:

Final_mail_list.to_excel(Path.cwd() / "Final_email_list_US80.xlsx",index=False)


# In[ ]:


print("All The Mails were sent to Vendors on",datetime.datetime.now().strftime("%d-%b-%Y %H-%M-%S"))


# In[ ]:


n = ToastNotifier()
n.show_toast("US80 Email Notification", f"{len(unique_values)} Emails Sent !!", duration=10)

