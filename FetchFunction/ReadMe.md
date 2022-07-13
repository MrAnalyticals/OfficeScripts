


**Scrape CSV web data with Excel Office Scripts without using Interfaces**

In this repo I use the fetch function to retrieve csv data from a URL. The code did not need to use any interfaces to handle the returned data. 
Also included is a routine of VBA that uses the QueryTables.Add method to retrieve data from a URL.
The URL in the demo is :  
https://raw.githubusercontent.com/treselle-systems/customer_churn_analysis/master/WA_Fn-UseC_-Telco-Customer-Churn.csv

There is a 5million cells and 5MB size limit when using the fetch function in Excel Office Scripts. 

This URL tests this. It returns the “too big” notification: 

https://raw.githubusercontent.com/MrAnalyticals/OfficeScripts/main/FetchFunction/FR%20Sales.csv

This URL has 100,000 rows and is able to be fetched: 

https://raw.githubusercontent.com/MrAnalyticals/OfficeScripts/main/FetchFunction/MX%20Sales100000.csv

Ref:
Read or write to large ranges using the Excel JavaScript API - Office Add-ins | Microsoft Docs

See this reference for an example of using the JSON return type with the fetch function. 
Use external fetch calls in Office Scripts - Office Scripts | Microsoft Docs
https://docs.microsoft.com/en-us/office/dev/scripts/resources/samples/external-fetch-calls


Here is the example VBA code script: 
![image](https://user-images.githubusercontent.com/47678539/178653157-78654652-4d69-478c-a56e-60a50ce6713f.png)
 

Size Limits documentation: 
Resource limits and performance optimization for Office Add-ins - Office Add-ins | Microsoft Docs
https://docs.microsoft.com/en-gb/office/dev/add-ins/concepts/resource-limits-and-performance-optimization#excel-add-ins

![image](https://user-images.githubusercontent.com/47678539/178653205-5eb8f3c7-ed55-4aa6-bce7-3371999efed1.png)

The Office Script Example FetchingCSVv3 : 

YouTube Demo Video:  
**YouTube Video Script**

Microsoft recently announced new Get Data functionality for Excel Online. It is being rolled out this year 2022. In the meantime you can use Office Scripts to perform webscraping to obtain data from online resources like GitHub.  In this video I demo Office Scripts code that uses the Fetch function to obtain a csv file without using an Interface – an artifice that structures returned data.
I also show a VBA script that performs the same action for comparison purposes. 
Whilst Office Scripts runs online, VBA does not and is therefore restricted to working only on your machine. With the ever-expanding role of cloud tools, data centres and the internet through mobile devices and other machines, onPemise solutions, like VBA, are becoming non-relevant. VBA was popular because of the simplicity of writing Visual Basic. You could automate a business process and save your business time and money. But if your business is working online with collaborative tools like Teams and Excel Online, VBA is not able to assist there. Office Scripts is.
I provide a link to the code used in this video at the end of this vide as well as in its description. 
Let’s run the VBA code first. We can see it is getting data successfully. 
Now let’s run the Office Script equivalent code. We can see it creates a new sheet to place the retrieved data into. There is, also, a notification of the number of rows contained in the data and the number of cells used to hold the data. These values are important for keeping track of size limits more on of which I will discuss later. 
Here is the same Office Script running but this time using a url to a GitHub repository.
The repos in that GitHub account are returned to the worksheet.
In Excel online we cannot, as yet create queries or connect to external data. If you first create those queries using Excel desktop and then upload that workbook to SharePoint you will be able to visualise the Query and Connections pane. But the functionality to refresh those queries and connections does not exist. Using the fetch function in Office Scripts allows you to obtain data from remote URLS despite this lack of functionality in Excel Online. To demo this I remove the values from some cells and attempt to refresh the query.
So, lets run through the Office Script code :
Talkthrough video.
As mentioned earlier here is the information about the limits using the fetch function within Excel. You can see the screenshot was taken from a Excel Add-ins help page and this is one thing I want to mention. Office Scripts and Excel Add-Ins both reference the same Excel Object model. That is, they are both programming Excel but one uses an IDE and one uses Excel itself. So, in this regard Office Scripts is the same as VBA because VBA, too, used an inbuilt IDE. I wonder if the code pane in Office Scripts will be built upon to be more advanced and even become an embedded version of Visual Studio Code. 
It is the Fetch function (not the async main method) that is preventing the operation of the button operating in Excel Desktop. 
Here is the VBA which uses the QueryTables.Add method to obtain data from a url.
In my next video I will show how to use DropBox.com as part of a solution that uses Office Scripts to fetch data that bypasses the 5million cells and 5MB size limit. It uses a method called chunking.  




