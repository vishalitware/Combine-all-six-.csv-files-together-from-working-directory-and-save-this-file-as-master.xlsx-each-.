
# coding: utf-8

# In[3]:


import glob,os
import  pandas as pd


# In[4]:


from pandas import DataFrame, ExcelWriter
writer = ExcelWriter(r'E:\model test\Master.xlsx')
for filename in glob.glob(os.path.join('E:/model test/','*.csv')):
    df_csv =pd.read_csv(filename)
    
    (_, f_name) = os.path.split(filename)
    (f_short_name,_) = os.path.splitext(f_name)
    
    df_csv.to_excel(writer, f_short_name, index=False)
    

writer.save()


# In[ ]:


get_ipython().system('pip install pdfkit')


# In[ ]:


import pdfkit


# In[ ]:


df = pd.read_excel(r'E:\model test\Master.xlsx',2)


# In[ ]:


df.head()


# In[ ]:


import win32com.client


# In[ ]:


from win32com import client
xlApp = client.Dispatch("Excel.Application")
books = xlApp.Workbooks.Open('E:/model test/Master.xlsx',3)
ws = books.Worksheets[0]


# In[ ]:


def to_pdf(fname):
    save_pdf = os.path.splitext(fname)[0] + '.pdf'
    excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    book = excel.Workbooks.Open(Filename = fname)
    book.ExportAsFixedFormat(c.xlTypePDF,save_pdf)
    sheet = None
    book = None
    excel.Quit()
    excel = None


# In[85]:


os.getcwd()

