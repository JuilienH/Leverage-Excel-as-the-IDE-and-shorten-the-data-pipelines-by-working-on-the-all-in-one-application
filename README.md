# Excel and Python: One is the Microsoft licensed application, and the other is the open-source object-oriented languange. So what are they put together for?

How to leverage Excel as an independent development environment (IDE) for python
 
Common Python libraires needed: pandas, numpy to initiate the development 
## Background
When your code development is ready to move in implementation, where you would like it to be? For a small-scope of project, we might just want to create a business-friendly development process to deliver the outputs in the Excel files. Like most programmers, I have tried different "code editors" to install python libraries, create a virtual envrionment prior to the start of off-line development. With current technology advancement, database tools, cloud platforms and any specific tools used in the data pipelines start to share a lot of common grounds. That said, the streamlined process is more easier to achieve in these days. The repository is to show you how you can use Excel as the data table, python eidtor, and the output files.
## Use case: Output layout conversion from the multidimensional strcuture to the tabular form
Oracle Hyperion Essbase is highly reliable tool for financial analysts to structure their finance data. The Excel add-in, Smart View is required to access the data hosted in Essbase. Although querying data on Essbase by hierarchies makes the financial reports easily read, data professionals, like data scientists and programmers would find it quite different from the usual scripting experience. All the data fetching, table joins, computations, and more data manupulation steps are transformed in selection options in many drop-downs, built in the financial-analysis popular tool. Due to the infrastructure divergence, there are challenges to access Essbase by open-source programming languange, such as python. The good news is: Python in Excel is available to Enterprise and Business users.
## Example to show business users that advanced analytics based on python is integrated in Excel
1. Get your code editor to complete one conversion script: Budget file

Prepare to move the script to production, and the platform is Excel!

2. Find out the location of python in Excel
![Screenshot 2025-04-15 112747](https://github.com/user-attachments/assets/d1d0c246-bc0b-4387-8598-0f2ddbb08103)

Python Editor in Excel looks much similar to a coding envrionment:

![Screenshot 2025-04-15 112811](https://github.com/user-attachments/assets/6acc2b09-53b2-4562-ac53-bfadd9f4a10a)

   You don't need to install python by the version anymore.

   You don't have to create virtual environment for each project anymore. Initiation botton shows all the pre-loaded python libraires. If you have specific pre-requisite libraires to be installed, the Initiation panel is where you can insert more library imports.
   
![Screenshot 2025-04-15 112853](https://github.com/user-attachments/assets/6784c364-79f1-4483-8c25-bcb3dea3e85d)

   You don't need to install common libraries anymore.
   
Now we need to slightly edit our python script that works in the code editor by doing the following steps:

4. Re-pivot the data-read step to the Excel spreadsheet, just like selecting Excel cells. The pandas python command ``read_excel`` is no longer needed. Instead, the common Excel cell selection is baked in the python script to make the pandas data frame.
   
   ``df_example = xl("'Essbase output file'!B375:F739",header=None)``
   
6. Remove the data export step. Instead, click the small icon to make your output shown on the Excel spreadsheet
   
   You can work on either the Excel function or the python script. All we have to do is to drop the whole edited chunk of python script in either places. If you are familiar with python development on your local desktop, you might frequently save the outputs in a folder by using the pandas python command ``to_csv``. Instead, we just need to make sure the output is a data frame as follows: 

   ``pd.DataFrame(df_example)``
7. The most fun part is to switch between Python Object and Excel Value in Excel.
   
   First, Python Object should be shown by default as we are working in python until this step.
   ![Screenshot 2025-04-15 145113](https://github.com/user-attachments/assets/ee74222d-9a07-4aa5-a7bf-81069f60f772)

   Second, simply click the drop-down to pick Excel Value.
  ![Screenshot 2025-04-15 145715](https://github.com/user-attachments/assets/a377f375-d541-4c99-b7e2-95bf3ce8734b)

   The magic moment is here: The output has been in Excel for review.
  ![Screenshot 2025-04-15 145716](https://github.com/user-attachments/assets/b7e8c995-2b59-4cba-b631-b011b352cf0d)

## Reference
https://support.microsoft.com/en-us/office/get-started-with-python-in-excel-a33fbcbe-065b-41d3-82cf-23d05397f53d
