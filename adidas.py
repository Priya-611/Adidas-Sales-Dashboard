import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

df=pd.read_excel("C:\\Users\\HP\\OneDrive\\Documents\\Project(Data Science)\\Excel\\Adidas US_Sales Datasets.xlsx",sheet_name='Sheet1')

# print(df.head())
# print(df.info())
# print(df.describe())

# print(df.isnull().sum())


# sns.heatmap(df.corr(numeric_only=True),annot=True)
# plt.show()


# print(df.dtypes)


df['Invoice Date']=pd.to_datetime(df['Invoice Date'],errors='coerce')


df['Year']=df['Invoice Date'].dt.year
df['Month']=df['Invoice Date'].dt.month_name()


with pd.ExcelWriter("C:\\Users\\HP\\OneDrive\\Documents\\Project(Data Science)\\Excel\\Adidas US_Sales Datasets.xlsx",engine='openpyxl',mode='a') as writer:
      df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
