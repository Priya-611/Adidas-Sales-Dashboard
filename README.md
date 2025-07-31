#  Adidas Sales Excel Dashboard

## Project Overview
This project presents a visually interactive Excel dashboard based on Adidas U.S. sales data. The dashboard enables users to explore key performance indicators (KPIs) and sales insights using clean visualizations. It offers a business-friendly view of operational and financial metrics.


---


## Dataset Source
- **Source**: [Adidas Sales Dataset on Kaggle](https://www.kaggle.com/datasets/heemalichaudhari/adidas-sales-dataset)
- **Sheet Used**: Sheet1
- **Preprocessing Done**: Cleaned and modified using Python


---


## Project Workflow

1. **Data Collection**  
   Retrieved the dataset from Kaggle and loaded Sheet1.

2. **Data Cleaning (Python)**  
   - Parsed and formatted `Invoice Date` column  
   - Extracted **Year** and **Month** from the invoice date  
   - Saved the cleaned data to a new sheet called **"Cleaned_Data"** in the same Excel file  
   - [See Python script: `adidas.py`](adidas.py)

3. **KPI Formulation**  
   - Calculated total orders, total sales, total units sold, max profit, and average operating margin

4. **Dashboard Design (Excel)**  
   - Built an interactive dashboard using charts, slicers, icons, and formatting
   - Applied custom formatting (e.g., millions format, colored indicators)


---


## Key KPIs

- **Total Orders**
- **Total Sales**
- **Total Units Sold**
- **Maximum Profit**
- **Average Operating Margin**


---


## Analytical Objectives

The dashboard aims to derive insights from various perspectives across different business dimensions. The specific objectives visualized in the dashboard are:

1. **Operating Profit by Sales Method**  
   Analyze which sales methods (In-store, Online, Outlet) generate the highest profit.

2. **Operating Profit Distribution by Year**  
   Observe how profit is distributed across different years to spot annual performance trends, for 2020 and 2021.

3. **Monthly Operating Profit by Retailer**  
   Track monthly profit patterns across various retailers to identify seasonal effects or top performers.

4. **Units Sold by Product**  
   Understand which product types or categories contribute most to the total units sold.

5. **Top 20 States by Units Sold**  
   Highlight the geographic regions with the highest product movement based on units sold.

6. **Average Sales by Retailer**  
   Compare average revenue figures among retailers to identify consistent performers.

7. **Yearly Total Sales by Region**  
   Visualize how each region contributes to total sales annually, using an area chart to emphasize trends, for 2020 and 2021.


---


## Features of the Dashboard

- **Year-wise Sales by Region**: Area chart for trend visualization  
- **Top KPIs**: Clearly defined with icons and numerical insights  
- **Custom Formatting**: Units shown in millions, margins with percent symbols  
- **Interactive Elements**: Filters/slicers to explore by year, product, region  


---


## Tools Used
- **Python**: `pandas`, `matplotlib`, `seaborn` for data cleaning and initial analysis  
- **Excel**: Dashboard creation, custom KPIs, visualization, and formatting


---


## 📎 File Structure
├── Adidas US_Sales Datasets.xlsx
│ ├── Sheet1 (Raw Data)
│ └── Cleaned_Data (Modified using Python)
├── adidas.py (Python script for preprocessing)
└── README.md (This file)


