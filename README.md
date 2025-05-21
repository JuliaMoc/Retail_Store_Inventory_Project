# Sales & Stock Insights: A Complete Excel BI Solution

The Retail Store Inventory project was created as a portfolio piece to demonstrate how Excel – often underestimated – can be used as a powerful tool for business intelligence, data storytelling, and operational insight, even with large datasets (70K+ rows). The aim was to simulate a real-world retail environment by creating a dynamic dashboard and an automated daily report. These tools highlight key performance indicators (KPIs) that can support retailers in optimizing their daily operations and overall performance.

### Real-World Use Case Simulation

In a real retail organization, Excel is often still the go-to tool for many daily processes – especially in small to mid-sized businesses or operational roles. This project replicates common workflows and challenges from such environments, including:

- Cleaning messy exports from outdated systems
- Handling inconsistent formats (e.g. currency/date issues)
- Manually integrating external factors like weather or holidays
- Creating easy-to-read reports for non-technical stakeholders

### Data Sources

The dataset used in this project was sourced from Kaggle (Store Inventory Demand Forecasting (Synthetic)). 
This open-source dataset provides synthetic yet realistic daily inventory and sales data across multiple stores and product categories.

### Data Processing in Excel

1️⃣ Data Cleaning
- Imported with Power Query
- Fixed corrupted date/price formats (Excel misinterpreted some price fields as dates)
- Used blank cells instead of 0s for cleaner analysis and to avoid bias
- Applied find & replace to convert price formats (e.g., . to ,)
- Converted all date/number/currency columns to proper types

2️⃣ Calculated Fields
Added new calculated columns using formulas:
- Units Total, Total Sales, Total Sales incl. Ordered Units
- Diff. Total Sales, % Diff. Total Sales
- Diff. Competitor Price, % Diff. Competitor Price

3️⃣ Design & Layout of the table
- Applied table formatting with filters
- Used traffic light color coding for:
    - Discount: Red for max discounts (>20%)
    - Competitor Price Diff: Gradient scale for visual contrast
    
4️⃣ Analysis & Visualizations
Pivot Tables & Interactive Dashboard (also in Excel)
- KPI overview: Total Sales, Total Quantity, and Average Sell-Through Rate
- Total Sales comparison: 2022 vs. 2023
- Sales breakdown: by Retail Store, Category, and Weather Condition
- Interactive Button: generates a Daily KPI Report (also available as PDF)
- Dynamic Charts: all charts can be filtered by year and month

  <img width="1305" alt="Dashboard" src="https://github.com/user-attachments/assets/18ba5cc8-7c60-4683-858e-15b1ad7f0cc3" />


5️⃣ Automated Daily KPI Report (with Macros/VBA)
- Filtered sales view by date & store
- Aggregated sales by category (incl. % differences)
- Highlights key discount activity
- Displays weather condition context

### Extra Features & Ideas for Expansion

Here are some ideas that could take this project even further:
- Predictive Add-on
    - Integrate time series forecasting models in Excel using Python (via Excel-Python Add-ins or exporting data to Jupyter)
- Connecting to live APIs (e.g. weather, product data)
    - Pull live weather data from an API into Excel for dynamic forecasting
- Scenario Simulation
    - Add "What-if" sliders to test how discounts or weather changes affect sales

### Tools & Tech Used

- Microsoft Excel (Power Query, Pivot Tables, Conditional Formatting)
- VBA for automation
- Kaggle Dataset (synthetic inventory data)

### Get Started

- Download the Excel file from this repo
- Enable Macros to use the automated report
- Filter, explore, and analyze the data
- Try connecting your own dataset!

### Let's Connect

Found this useful or want to contribute?
- ⭐ Star the repo
- 📬 Raise an issue
- 🔁 Fork the project
- 💬 Share your version!
