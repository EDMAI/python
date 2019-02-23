
# coding: utf-8

# In[42]:


import pandas as pd


# # 네이버

# In[43]:


#skiprows 필요없는 첫 번째 행 제거
#encoding?

n_data=pd.read_csv('C:\\Users\\jungh\\Desktop\\데이터분석\\파이썬 강의\\강의data\\캠페인보고서,--.csv',engine='python',skiprows=1,encoding='utf-8')
n_data


# In[44]:


n_day=n_data['일별'].str.replace('.','-')
n_day


# In[45]:


n_day.str[:-1]


# In[46]:


n_data['일별']=n_day.str[:-1]


# In[47]:


n_data


# In[48]:


n_imp=n_data['노출수'].str.replace(',','')
n_imp


# In[49]:


n_imp.astype(int)


# In[50]:


n_data['노출수']=n_imp.astype(int)


# In[51]:


n_data


# In[52]:


n_cost=n_data['총비용(VAT포함,원)'].str.replace(',','')
n_cost


# In[53]:


n_data['총비용(VAT포함,원)']=n_cost.astype(int)


# In[54]:


#완성된 데이터
n_data


# # 다음

# In[55]:


d_data=pd.read_csv('C:\\Users\\jungh\\Desktop\\데이터분석\\파이썬 강의\\강의data\\추이보고서_---_20190201-20190220.csv',engine='python')


# In[56]:


d_data


# In[57]:


d_day=n_data['일별'].str.replace('.','-')
d_day


# In[58]:


d_data['날짜']=d_day


# In[59]:


d_data


# # FaceBook

# In[61]:


f_data=pd.read_csv('C:\\Users\\jungh\\Desktop\\데이터분석\\파이썬 강의\\강의data\\제목-없음-Feb-1-2019-Feb-22-2019.csv',engine='python',encoding='utf-8',names=['일','노출','클릭(전체)','지출 금액','보고 시작','보고 종료'],skiprows=1)


# In[62]:


f_data


# # 합치기

# In[63]:


#먼저 칼럼 알아보기
n_data.columns


# In[64]:


d_data.columns


# In[65]:


g_data.columns


# In[66]:


f_data.columns


# In[67]:


#네이버 합치기
n_result=n_data[['일별','노출수','클릭수','총비용(VAT포함,원)']]
n_result


# In[68]:


n_daily=n_result.set_index('일별')
n_daily


# In[69]:


#다음
d_result=d_data[['날짜','노출수','클릭수','총비용']]
d_result


# In[70]:


d_daily=d_result.set_index('날짜')
d_daily


# In[71]:


#구글
g_result=g_data[['일','노출수','클릭수','비용']]
g_result


# In[72]:


g_daily=g_result.set_index('일')
g_daily


# In[73]:


#페이스북
f_result=f_data[['일','노출','클릭(전체)','지출 금액']]
f_result


# In[74]:


f_daily=f_result.set_index('일')
f_daily


# # 전체 보고서 합치기

# In[75]:


data=pd.concat([n_daily,d_daily,g_daily,f_daily],axis=1)
data


# In[76]:


data_f=data.fillna(0)
data_f


# In[77]:


pd.options.display.float_format = '{:,}'.format
pd.set_option('display.float_format', None)


# In[78]:


data_f


# In[79]:


data_all=data_f.applymap('{:,}'.format)


# In[ ]:


import openpyxl
wb = openpyxl.load_workbook('C:\\Users\\jungh\\Desktop\\데이터분석\\파이썬 강의\\강의data\\리포트sample.xlsx')


# In[ ]:


wb.get_sheet_names()


# In[ ]:


ws=wb.get_sheet_by_name('리포트1')


# In[ ]:


len(data_f.columns)


# In[ ]:


len(data_f)


# In[ ]:


for x in range(0,22):
    for y in range(0,12):
        ws.cell(row=3+x,column=3+y).value=data_f.iloc[x,y]
wb.save('result.xlsx')
wb.close()

