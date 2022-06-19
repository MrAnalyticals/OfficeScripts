
**Microsoft Webinar – Using Excel Office Scripts with Power Automate**
Register here to see the webinar: https://msft.it/6046bYD9o
See the QnA below.

**References**
Excel Office Scripts Tutorial 
Pass data to scripts in an automatically-run Power Automate flow - Office Scripts | Microsoft Docs
https://docs.microsoft.com/en-gb/office/dev/scripts/tutorials/excel-power-automate-trigger 

**Video Script**
The Microsoft Excel Office Scripts Team has created a very clear webinar that introduces Excel Office scripts and also gives a very insightful code demo. 
The demo shows how to integrate your Office Script with a Power Automate Flow passing a parameter into and out of the Excel Office Script.
You need to register to see the webinar at this URL:
https://msft.it/6046bYD9o
In the GitHub repo, for this video, I have posted the QnA for that webinar. 
To further expand your knowledge about Office Scripts Microsoft has also provided a detailed Tutorial here: 
I have put the url in the description of this video.
https://docs.microsoft.com/en-gb/office/dev/scripts/tutorials/excel-power-automate-trigger 

For those knowledgeable about Azure and Azure Functions one of the things that is not very well known, about Office Scripts, is that it can replace the Azure Functions service. That is, you can have your code run as part of  a Flow triggered by an event. This is exactly what Azure Functions does. 
Hope you enjoy this video. Check out my business website too.




**Promotional**
Analysis Cloud YouTube video: https://youtu.be/Tzek7PmFlhk

Analysis Cloud Business Website: https://www.analysis.ie
The Irish Cloud Company

 

 

**Webinar QnA**

•	Scripts are saved to OneDrive --- is there an ability to save scripts to a file so other users can easily run the scripts as well?
You can share an Office Script in a workbook so that other users who have access to that workbook can view and run the script. We are also currently working on a feature that allows you to save a script to a SharePoint site so that members of that site will have access to that script as well.
•	Can we use power automate to trigger an excel app to process a spreadsheet and then ingest resulting file into a sql data table?
Yes. You can first create an Office Script from Excel on the web. That script can process the workbook and return whatever result you want. Then from Power Automate side, you can trigger that script through the "Run script" action against your chosen workbook. You can then potentially add a SQL action (https://docs.microsoft.com/en-us/connectors/sql/) to ingest the results from the Run script. Hope you can find more information regarding this from this webinar, or from here: https://docs.microsoft.com/en-us/office/dev/scripts/tutorials/excel-power-automate-manual.
•	I'm an O365 personal user
Office Scripts is not currently available for personal licenses.
•	Is the Automate button only available in Excel online? Not native desktop for O365 users?
We are working on taking the feature to the different platforms. Triggering a script from a button is available on win32 Excel and more features will be lighting up in the future.
•	Hello, Is Scripts essentially Web version of Macros?
Office Scripts will likely never be a full replacement for VBA. However, if you are looking to create automations with more granular sharing permissions and want to run them cross-platform or as a part of a cloud Flow in Power Automate, Office Scripts may be the right tool for you. Over the years we have seen a rise in collaboration in the workplace. By storing Office Scripts in the cloud and providing users with sharing management options, we want to make Office Scripts easy and safe to share with others. Though Office Scripts currently only works online, we also intend to expand to support running scripts cross-platform so that users can automate their tasks wherever they are. By contrast, VBA is focused on individual workflows and is only supported on Desktop. You can find more about the differences between Office Scripts and VBA here: https://docs.microsoft.com/office/dev/scripts/resources/vba-differences
•	why do i not have the Automate tab on my excel?
Office Scripts is available with a commercial or educational Microsoft 365 license with access to the Microsoft 365 Office desktop apps.
•	what if you don't see the automate tab on your ribbon?
Office Scripts is available with a commercial or educational Microsoft 365 license with access to the Microsoft 365 Office desktop apps.
•	I dont see that i have Automate in my excel
Office Scripts is available with a commercial or educational Microsoft 365 license with access to the Microsoft 365 Office desktop apps.
•	will we get example power automate workflows?
We have prepared several Power Automte templates that leverages Office Scripts and other popular Microsoft services/products. You can start from there and create flows. Please go to https://us.flow.microsoft.com/en-us/templates/ and search for "Office Scripts".
•	Is scripting available for word as well?
We are initially focused on Excel on the web, but we hope to grow beyond that in the future. Let us know where you would most like to see Office Scripts: http://aka.ms/ExcelSuggestions.
•	Is office script available to Excel files opened in desktop or only to Excel on web files?
Office Scripts that are linked to a button in Excel can be run in Excel desktop (win32) currently, and more features are coming to win32 in the near future.
•	If I don't have the Automate tab, does that mean my my employer has restricted access to this capability?
Office Scripts is available with a commercial or educational Microsoft 365 license with access to the Microsoft 365 Office desktop apps. If you meet these requirements and are still not seeing the Automate tab, it's possible that your admin has disabled the feature.
•	Can office scripts refresh power queries?
There isn't an API available today that allows you to refresh Power Query data connections. However, we understand that this is a huge scenario for many of our users, and this is something we hope to invest in soon.
•	Oof. Any plans to make them available for GCC?
We hear your feedback and are continuing to evaluate where Office Scripts will be offered in the future.
•	Recent experience for me with scripts on web does not seem to enabled for my office 365 G3 GCC tenant, is thats the case?
Yes, Office Scripts is not currently available in GCC. We hear your feedback and are continuing to evaluate where Office Scripts will be offered in the future.
•	Can power automate work with open source spreadsheets like Libre? Or import Libre into Excel first?
At this moment Office Scripts only supports Excel. You may need to import Libre into Excel first.
•	Can VBA be converted to scripts?
No, we do not currently have a VBA to Office Scripts converter.
•	I have Excel with an E5 license but I do not have an Automate tab.
It's possible that your admin has disabled the feature or there's some other problem with your environment. Please follow the steps here: https://docs.microsoft.com/en-us/office/dev/scripts/testing/troubleshooting#automate-tab-not-appearing-or-office-scripts-unavailable
•	Our org has E3, M365 Apps for enterprise -- I don't see the tab: does it need to be enabled by sysadmins?
It's possible that your admin has disabled the feature or there's some other problem with your environment. Please follow the steps here: https://docs.microsoft.com/en-us/office/dev/scripts/testing/troubleshooting#automate-tab-not-appearing-or-office-scripts-unavailable
•	Can Scripts be debugged line-by-line (like VBA)?
We do not currently support this capability; however, we welcome your suggestions at http://aka.ms/ExcelSuggestions.
•	Will there be a way to re-watch or download this webinar?
Yes, the webinar will be available on-demand with the same registration link. You will also get an email with that link to watch it on-demand.
•	Where can we find definitions of the code commands
We have examples, tutorials, and reference materials here https://docs.microsoft.com/en-us/office/dev/scripts/ https://docs.microsoft.com/en-us/javascript/api/office-scripts/overview?view=office-scripts
•	Does Scripts work with PowerAutomate desktop in any way?
Currently the "Run script" action is only available for cloud flows. But you can trigger a desktop flow from within a cloud flow containing a Run script action, so you probably can try that out.
•	How is Office Script addressing the issue similar with the bad guys using VBA in security attack in the past?
Office Scripts run in a sandbox that place limits on what the script can do (vs VBA which has full trust access to your whole local machine). There's a bit more here in the Security section: https://docs.microsoft.com/en-us/office/dev/scripts/resources/vba-differences
•	What if you add more data to the January tab? After you have recorded action
It depends on how your worksheet is structured, but you can use the Office Scripts API to reference data dynamically so that every time you run a script, it gets the entire table's data or the entire range's data. Check out our documentation here: https://docs.microsoft.com/office/dev/scripts/
Does the script have a static or dynamic reference to the workbook?
The script itself isn't tied to any specific workbook. You can run it against any workbook you open in Excel on the web. Also through the "Run script" action in Power Automate, you can run your script against any workbook you pick there.
•	is office scripts only available with web version of excel? if not, can you walk through how to add it to desktop version please?
It is only available in Excel for the web at the moment. However, we just released a feature that allows you to add a button in the workbook that runs a script, and this functionality also works in desktop!
•	That's ridiculous. The script CANNOT run against any workbook because it will FAIL.
As long as you have access to that workbook, you should be able to run the script against it. But of course if your script references some specific "named" entities like worksheet, table, or chart, you will need to make sure they exist in that workbook.
•	This is very encouraging and powerful - great presentation and PLEASE make this available for GCC tenants!
Thank you, we hear your feedback loud and clear!
•	Do end users need to have the O365 licenses, or is it just the individual who owns the process/workbooks? Example: I have the necessary license to create the form/automated process, but my client does not. Can they fill out a form online and still trigger all of the back end automation?
yes, the automated Flow would still work because the Flow will run on your behalf.
•	Can I use an office script in excel to refresh data in a worksheet/excel file stored in SharePoint with Power Query data that is hosted in another location?
There isn't an API available today that allows you to refresh Power Query data connections. However, we understand that this is a huge scenario for many of our users, and this is something we hope to invest in soon.

