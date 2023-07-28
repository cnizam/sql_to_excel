# sql_to_excel
Today, many companies still use excel spreadsheets for report sharing. Since the SQL query results are exported to Excel as raw data, a series of time-consuming operations are required to format the table and display the data in the appropriate structure. This repository has been prepared to automate these processes. 

The following operations are applied to the raw data in the created Excel Table:

- The data type of the column is determined (integer, float, percentage, object, date) and appropriate formatting is done.
- Title background and font are colored.
- Sequential blue-white coloring is applied to increase line distinctiveness.
- Table border is added.
- Column widths are adjusted automatically.
- Output adjustment is made so that the table fits horizontally on a page.

More than one sql query can be converted at once.
