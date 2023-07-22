# MS-Excel-Data-Analysis

For this project, I conducted data analysis employing Microsoft Excel and crafted a straightforward dashboard. Subsequently, I will elaborate on the entire process.

In this project, I utilized an openly accessible dataset obtained from GitHub, which contains information related to a bike store's conversion rate. The dataset is carefully designed to exclude any Personally Identifiable Information (PII). It comprises 13 columns, including unique ID, Marital Status, Gender, Income, Children, Education, Occupation, Home Owner status, Cars, Commute Distance, Region, Age, and a binary indicator for Purchased Bike (whether they made a bike purchase or not).

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/740e7e8b-6c11-4023-9877-47458e6e613d)


The initial step involved in this project was data preparation. To begin, I duplicated the raw data and pasted it into a fresh sheet within Excel, labeling it as the "Working Sheet." This duplication served as a precautionary measure, ensuring that a backup of the original data was readily available in case any errors arose during the data cleaning process.

**Verifying and eliminating duplicate entries** is a crucial aspect of data cleaning. In this project, I proceeded with this step by examining the dataset for any duplicate values and subsequently removing them.

To accomplish this, I began by selecting the entire dataset within the table. Then, I navigated to the "Data" tab in the Excel toolbar and clicked on the "Remove Duplicates" button. This action opened the "Remove Duplicates" window. Upon clicking "OK," the system automatically scanned the data for duplicates and proceeded to eliminate them. As a result, a total of 26 duplicate values were detected and subsequently deleted from the dataset.

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/f9a16adf-2012-4cd4-aa6b-061cd43fe2e0)

In order to enhance clarity and comprehension for stakeholders, I opted to **replace abbreviations with their corresponding full-forms in the dataset**. Specifically, I addressed the abbreviations "M" and "F" in the Gender column, which represent "male" and "female," as well as "M" and "S" in the Marital Status column, which stand for "married" and "single," respectively.

To achieve this, I utilized the 'Find and Replace' feature in MS Excel. Initially, I selected the Marital Status column. To access the 'Find and Replace' window.

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/d0023c30-7c78-4e6a-9ac8-c8a72d13d746)


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/168cf80c-4967-45e4-9658-c27306942c2c)

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/58a22044-25be-4412-af9d-8cfab81fea46)

During the subsequent step, I focused on the Income column, where I made necessary adjustments to ensure data integrity and consistency. Firstly, I changed the cell format from 'General' to 'Currency' to provide clear monetary representation, as it was initially left in a general format. This modification was essential to prevent potential issues in future analyses. Ambiguous formats can lead to inadvertent errors that might significantly impact the results.

Additionally, since the data did not include decimal values for income, I proceeded to eliminate the decimal places in the Income column. This action was taken to maintain uniformity and accuracy, as displaying decimal places where they were not relevant could potentially cause confusion and misinterpretation of the data.


Indeed, the **Age column** in the dataset appears to contain a wide range of different ages, which might not be ideal for creating a clean and easily interpretable dashboard. To address this issue, I decided to group the ages into distinct age brackets. This step would help in presenting a clearer visual representation of the data trends on the dashboard.

By categorizing the ages into broader groups or ranges, we can effectively summarize the data and highlight patterns without overwhelming the dashboard with numerous individual age values. This simplification will enable stakeholders to grasp the insights and trends more easily and make better-informed decisions based on the data.


**I categorized the ages into three Age Brackets: "Adolescent" (age < 31), "Middle Aged" (31 <= age < 54), and "Old" (age >= 54).**

To achieve this, I employed the IF statement on the Age Column in Excel:

For the "Adolescent" bracket: I used the formula =IF(L2 < 31, "Adolescent", "Invalid"). This means that if the value in cell L2 is less than 31, it is labeled as "Adolescent"; otherwise, it is marked as "Invalid."

For the "Middle Aged" bracket: The formula used was =IF(L2 >= 31, "Middle Aged", IF(L2 < 54, "Adolescent", "Invalid")). The syntax is similar to the previous one, but in the event that the first condition is not met (age is not greater than or equal to 31), it checks whether the age is less than 54. If true, it is labeled as "Adolescent"; otherwise, it defaults to "Invalid."

For the "Old" bracket: I applied the formula =IF(L2 >= 54, "Old", IF(L2 >= 31, "Middle Aged", "Adolescent")). Again, the structure is the same, with three nested IF statements. If the age is greater than or equal to 54, it is labeled as "Old." If not, it checks whether the age is greater than or equal to 31 to be categorized as "Middle Aged." If neither of these conditions is met, it defaults to "Adolescent."

The use of nested IF statements allowed for efficient age grouping, ensuring a clearer presentation of data on the dashboard.


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/7c09354f-096c-4964-aeff-a5bf0aadc671)


**Pivot Tables**

Upon completing the initial data cleaning process, my focus shifted to retaining only the relevant data necessary for building the intended dashboard. To accomplish this efficiently, I opted to employ Pivot Tables, which provide an interactive and efficient means of summarizing extensive data without altering the original worksheet.

To proceed, I created a new sheet and designated it as "Pivot Table." Within this sheet, I crafted four Pivot Tables in total. These tables allowed me to consolidate and organize the essential data elements essential for the dashboard creation, facilitating a more streamlined and effective visualization process.

**Average Income In Relation To Gender and its Effect on Sale:**
The first pivot table that I created was the average income of the people in relation to their gender and how it affected the bike sale accordingly. For this, I used the following steps

**Pivot Table 1**

In the new sheet, click Insert, then Pivot Table. A dialog box appears titled “PivotTable from table or range”.

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/827e4de4-ccaa-498d-b5be-da8e84706610)

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/0c03304b-b33f-408d-bbf4-841024e25538)

For the initial pivot table, I proceeded by dragging the "Income" field into the "Values" section and configured the 'Value Field Setting' to calculate the average income. This allowed the pivot table to display the average income values. Next, I placed the "Purchased Bike" field into the "Column" area and the "Gender" field into the "Rows" area as illustrated below:

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/fd4d9bd1-14f0-4b05-b667-3b1eab6e2312)

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/cb003a26-5cf9-4b06-9435-a991028afdb7)

This pivot table provides insights into the average income of male and female customers and their bike purchase behavior. Evidently, the table reveals that, in this specific scenario, the average income of male customers surpasses that of female customers. Moreover, it becomes apparent that the average income of customers plays a crucial role in influencing their purchasing decisions. As observed in the pivot table, those who purchased the bike exhibited higher average income compared to those who did not make a purchase, though it's essential to acknowledge that income might be just one of several factors affecting customers' buying choices.

Once the Pivot Table is ready, a visual representation or chart can be generated using the data, which will later be utilized to create the Dashboard. The process of creating the visual is straightforward and involves the following steps:

Choose any cell within the Pivot Table that has been created.
Proceed to the "Insert" tab on the toolbar.
From there, select the option "Recommended Charts."


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/a3f8b989-03f7-4550-9582-851557bc0f95)


Within the dialog box, opt for the "Column" option, and choose a visually appealing and neat column chart that suits your preferences. Once selected, the column chart will be added to the sheet. To further customize the chart's appearance according to your liking, utilize the '+' icon located at the top-right of the chart to access the formatting options. This allows you to modify the chart's style and layout, ensuring it aligns with your desired aesthetics.

![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/13dd104e-7f0d-472c-b0ab-81af8d717a53)

**Pivot Table 2**

The second pivot table examines the influence of customers' commute distance from their homes to their workplaces or other significant locations on bike purchases. The process of creating this pivot table follows the same steps as explained in the previous section (PivotTable 1). The subsequent actions are as follows:

Drag the "Commute Distance" field into the "Rows" area.
Place the "Purchased Bike" field into both the "Columns" and "Values" areas.
In the "Values" area, modify the 'Value Field Settings' to "Count."


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/851a7f9b-bd52-4355-9893-ecfc8472194e)


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/a8118449-be4d-4c77-b4bb-dae70bceac1e)

I encountered a challenge with this pivot table due to a row named "10 Miles+." Excel was uncertain about the correct order for this row, resulting in its placement in the second position. This incorrect positioning led to a detrimental impact on the visual representation, causing an abrupt dip in the graph, which was visually misleading.

To address this issue, I manually resolved it by returning to the Working Sheet. There, I selected the Commute Distance column and utilized the 'Find and Replace' feature to change all instances of "10 Miles+" to "More than 10 miles." This manual adjustment rectified the order, ensuring the correct representation of the data in the pivot table and preserving the visual integrity of the graph.


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/88ec92a4-cb3f-4c99-9d28-47b9c14d28c7)

**Pivot Table 3**

The third pivot table examines the connection between the customer's age and their buying behavior. This pivot table was added to the sheet using the same procedure as explained earlier (PivotTable 1). The subsequent steps are as follows:

Within the PivotTable Fields, drag the "Age Bracket" field into the "Rows" area.
Place the "Purchased Bike" field into both the "Columns" and "Values" areas.
In the "Values" area, modify the 'Value Field Settings' to "Count."


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/e39f2fe8-c24b-4b5f-8689-04f1f8811c2a)


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/3877bbeb-e420-4790-ae2e-878dc5b10ce4)


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/21515508-3647-4d00-880a-8c211947e60c)

**Creating dashboard**


To construct the ultimate dashboard, I initiated the process by generating a new sheet titled "Dashboard." Then, I proceeded to copy each visual that I previously created in the Pivot Table sheet and diligently pasted them into the Dashboard sheet, one by one. Following this, I meticulously formatted the sheet to eliminate all grids, fostering a clear and visually appealing display. This meticulous approach resulted in the creation of an organized and informative dashboard that effectively communicates the analyzed data to the stakeholders.

To eliminate the grid lines from the cells, navigate to the Toolbar and locate the "View" tab. Within this tab, find the "Gridlines" option and uncheck the corresponding box. By doing so, the gridlines will be removed, resulting in a cleaner and more streamlined appearance for the dashboard. This adjustment enhances the visual clarity, ensuring that the focus remains on the data and the visuals without any distracting grid lines.



After removing the grid lines, the next step involved creating a title section for the dashboard. I designated the dashboard as "Bike Sales Dashboard." To enhance the visual appeal and coherency of the dashboard, I manually adjusted all the charts, ensuring they were both visually appealing and informative. To achieve a polished and well-aligned appearance, I utilized the "Align" option in the Page Layout tab. This feature allowed me to precisely align the charts, ensuring their boundaries matched perfectly and contributed to an aesthetically pleasing layout. The meticulous alignment and adjustments helped in presenting the data in a clear and organized manner, enhancing the overall effectiveness of the Bike Sales Dashboard.


Once the visual aspect of the dashboard was complete, I sought to enhance its interactivity and user engagement by incorporating additional filters. These filters allowed users and stakeholders to explore the data more dynamically and gain further insights. To achieve this, I utilized the "Insert Slicers" option.

To access this feature, I navigated to the PivotChart Analyze tab, which can be found under the Toolbar. From there, I selected the "Insert Slicers" option, which enabled me to add interactive filters to the dashboard. These slicers provided users with the ability to filter and view specific data subsets, making the dashboard more engaging and informative. With the inclusion of these filters, users could extract valuable additional information from the Bike Sales Dashboard, contributing to a more interactive and user-friendly data exploration experience.

<img width="1339" alt="Screenshot 2023-07-22 at 11 14 14 AM" src="https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/352d71cb-3146-4ca6-97cc-09567102199b">


<img width="259" alt="Screenshot 2023-07-22 at 11 08 02 AM" src="https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/d34ecd9c-dad8-424a-b8bc-b5dc5d4579a1">


In the "Insert Slicers" window, I had the opportunity to choose the options for which I wanted to insert filters. In my case, I opted to include filters for Marital Status, Region, and Education. By selecting these specific options, users could easily interact with the dashboard and explore the data based on various marital statuses, different regions, and levels of education. This increased level of interactivity provided stakeholders with a more comprehensive understanding of the bike sales data from diverse perspectives, further enriching their data exploration experience.

To synchronize the filters across all the charts in the dashboard, we need to link the slicers we created to all the pivot tables. To achieve this, follow these steps:

Start by selecting a slicer that you want to link to the pivot tables.
Proceed to the "Slicer" tab, which can be found under the Toolbar.
By linking the slicer to the pivot tables, changes made through the slicer will now apply uniformly to all the charts in the dashboard. This ensures consistent and coordinated data filtering across the entire dashboard, enhancing the overall user experience and data exploration capabilities.


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/1d732466-1ea3-47ab-a4bf-f9e622f90e57)


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/5b1ab67b-2520-4517-a58a-74d386a5305c)

After selecting all the PivotTables we created, click the "OK" button to establish the connection between the slicer and the charts.

By doing so, all the charts in the dashboard become linked to the specific slicer, allowing them to respond to filter changes made through the slicer.

To ensure consistency and interactivity, repeat the above steps for the other two slicers as well.

With this linkage in place, applying different filters through the slicers will cause all the visuals in the dashboard to adjust accordingly. This synchronization across the charts grants users a seamless and dynamic experience while exploring the data from various angles, enriching their understanding and insights.

**Dashboard outcome**


![image](https://github.com/Nidhiwa/MS-Excel-Data-Analysis/assets/88158951/086f3962-fea8-42e3-b1b6-8095925092f0)








