# Data Analysis using MS Excel - Pivot-Tables, Data-Modelling and Dashboard Creation.

Recently, I analysed a sales and customer dataset using Excel to draw insights that could drive actionable steps for stakeholders. The dataset, provided by a U.S. company, was stored in an Excel table and covered orders from 2014 to 2017. 

Let me walk you through the steps I took and the thought process behind each one:


**Data Exploration:**

At this stage, I typically review the raw data to understand its structure and the relationships between different columns. This includes descriptive information such as the number of rows and columns.

![image](https://github.com/user-attachments/assets/d779a58b-63fe-4240-bd4b-affa69f00287)

**Data Cleaning & Transformation:**

To import the table into Power Query, I selected the "Get Data" > "From Table/Range" option, found under the Data tab.

![image](https://github.com/user-attachments/assets/341f18a4-aab2-4058-a45f-3f93201a2182)


![image](https://github.com/user-attachments/assets/958cf8ca-cb08-4967-aa98-3aeae7a67b2e)

### Transformation Steps in Power Query

**•	Ensuring Data Quality:** I enabled the Column Quality feature to help identify and address any errors in each column.

**•	Splitting the Column by Delimiter:** Using the "Split Column" option found in the Home tab, I separated the Date from the Time in the Date column.

**•	Removing Irrelevant Columns:** I removed unnecessary columns to reduce the model size and improve report performance.

**•	Rectifying Locale Issues:** While transforming the Date column, I encountered an error when changing the column data type to Date. This was resolved by right-clicking the column header, selecting "Change Type" > "Using Locale," and choosing the appropriate locale for the date format.

**•	Removing Null Values:** I ensured that missing or null values were removed to avoid unexpected results during analysis.

**•	Choosing the Appropriate Data Type:** I made sure that the correct data types were selected for all columns to ensure data integrity.

Once the transformation was complete, I loaded the cleansed data into the Excel sheet and added the table to the data model.

![image](https://github.com/user-attachments/assets/f378d353-1469-46f4-bbfe-185fbb123d63)

![image](https://github.com/user-attachments/assets/47514567-9bc4-46ed-adc1-2fb090897b6a)

**Creating a Date Table:** Once the table was added to the data model, I created a dedicated Date Table to support time intelligence functions in the analysis. The Date Table is automatically populated with the earliest and latest dates by scanning through the date columns in the fact table.

![image](https://github.com/user-attachments/assets/15413d7c-f4b1-46af-9753-26f22730c715)

![image](https://github.com/user-attachments/assets/e4cb2be4-efc2-4468-b77e-87d34b7dc7bb)

I then connected *Date[Date]* to *Customer_Table[Order Date]* using a One-to-Many relationship. I also connected *Date[Date]* to *Customer_Table[Ship Date]* using a One-to-Many relationship. This is an example 

of a **Role-Playing Dimension**, where a dimension table can be used multiple times within a fact table. In this case, the Date Table is acting as a role-playing dimension with different relationships being used for different date columns.

However, there's a catch — **only one relationship can be active at a time**. Since my analysis is focused on orders, I made that relationship active, as seen in the image below with the bold line, while the inactive relationship can be identified with the broken lines.

![image](https://github.com/user-attachments/assets/4c6e66db-b7f6-4d18-8118-9caa142e2e2d)

**Building the Pivot Table and Dashboard:** I created two new sheets: one to house my pivot table for in-depth analysis, and the other to build the dashboard. This approach helps to keep things organised.

![image](https://github.com/user-attachments/assets/d9cd678e-97ed-4bbf-9182-a1636f5f47db)

I created the pivot table, and built relevant measures such as Total Sales, Number of Customers, Number of Orders, etc., using optimised DAX functions as shown below.

![image](https://github.com/user-attachments/assets/fa95efe7-d7b0-4150-b803-1acc238e3ff6)

![image](https://github.com/user-attachments/assets/5da60547-1716-4451-bdcf-a3369f9f388b)

![image](https://github.com/user-attachments/assets/e39a68da-1ced-4ebe-a518-82a2624ac94c)

Once the pivot tables were created, I began mining the data by checking for patterns, trends, and anomalies before proceeding to create charts and graphs.

Before creating any visual, chart or graph, I typically first generated a pivot table for the relevant data, highlighted it, then selected the appropriate chart option that displayed the values clearly.

**Data Visualisation & Storytelling:**

While creating graphs and charts for the dashboard, I incorporated data storytelling techniques. I ensured that every chart answered a specific business question and used the most appropriate chart types. 

For instance:

•	KPI cards for high-level overviews

•	Line charts to emphasise trends over time

•	Column charts to compare sales over a time period

•	Maps to show top-performing states

I also highlighted key data points to draw attention and used data labels, colour, and design to make the dashboard and visuals more engaging for stakeholders. Overall, I focused on using simple visuals that 

were easy to interpret.

I designed the background using PowerPoint and imported it into Excel to use as the report’s background.

**Final Validation:** Once the report was complete, I validated the analysis by checking my double checking the DAX measures to ensure that the logic, filter context, and output values were correct.

**Here is the final result:**

![Excel Sales Analysis](https://github.com/user-attachments/assets/048e6838-fa12-4c6f-965c-f3d35453851a)






