# Excel and Python: One is the Microsoft licensed application, and the other is the open-source object-oriented languange. So what are they put together for?

How to leverage Excel as an independent development environment (IDE) for python
 
Common Python libraires needed: pandas, numpy to initiate the development 
## Background
When your code development is ready to move in implementation, where you would like it to be? For a small-scope of project, we might just want to create a business-friendly development process to deliver the outputs in the Excel files. Like most programmers, I have tried different "code editors" to install python libraries, create a virtual envrionment prior to the start of off-line development. With current technology advancement, database tools, cloud platforms and any specific tools used in the data pipelines start to share a lot of common ground. That said, the streamlined process is more easier to achieve in these days. The repository is to show you how you can use Excel as the data table, python eidtor, the output files.
## Use case: Output layout conversion from the multidimensional strcuture to the tabular form
Oracle Hyperion Essbase is highly reliable tool for financial analysts to structure their finance data. The Excel add-in, Smart View is required to access the data hosted in Essbase. Although querying data on Essbase by hierarchies makes the financial reports easily read, data professionals, like data scientists and programmers would find it quite different from the usual scripting experience. All the data fetching, table joins, computations, and more data manupulation steps are transformed in selection options in many drop-downs, built in the financial-analysis popular tool. Due to the infrastructure divergence, there are challenges to access Essbase by open-source programming languange, such as python. The good news is: Python in Excel is available to Enterprise and Business users.
## Example to show business users that our advanced analytics based on python is integrated in Excel
1. Get your code editor to complete one conversion: Budget file
2. Find out the location of python in Excel

   You don't need to install python by the version anymore.

   You don't have to create virtual environment for each project anymore.

   You don't need to install common libraries anymore.
4. Re-pivot the data-read step to the Excel spreadsheet, just like selecting Excel cells
5. Remove the data export step. Instead, click the small icon to make your output shown on the Excel spreadsheet
6. Insert Python on the Excel cell or Open Python editor
   
   You can do either like Excel function or the code editor. All we have to do is to drop the whole edited chunk of python script in.

## Reference
https://support.microsoft.com/en-us/office/get-started-with-python-in-excel-a33fbcbe-065b-41d3-82cf-23d05397f53d
