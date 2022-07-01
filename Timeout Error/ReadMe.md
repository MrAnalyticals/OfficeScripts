
**Avoiding and Managing the Excel Office Scripts 2-minute timeout error**

YouTube video: https://youtu.be/b2hLdkioPGg
**Video Script**

In a previous video, I did, demoing the Knights Tour search using Office Scripts I encountered a timeout error. Let us, in this video, have a look at that issue in more detail. 
In the associated Flow it timed out after 40 minutes. 



In the following document Platform limits and requirements with Office Scripts - Office Scripts | Microsoft Docs it mentions about Data limits and “long-running script” and the 120-second timeout for synchronous Power Automate operations. That link takes you to Limits and configuration - Power Automate | Microsoft Docs where it states, for “Outbound synchronous requests”, the limit is 120 seconds. 
In this screenshot you can see the Office Scripts Action, for my flow, does, indeed, reach this limit. See the red box. Also note the 8 retries occurred message coloured yellow.

![image](https://user-images.githubusercontent.com/47678539/176802917-461bc689-89b0-4c94-8e6a-a4251c8ad367.png)


 


In Microsoft documentation Limits and configuration - Power Automate | Microsoft Docs
, in the section entitled “Retry Policy”, 

![image](https://user-images.githubusercontent.com/47678539/176802934-860487c5-f316-4b19-9ddf-0a61b3525c7c.png)
 

there you will see that in my case the 8 retries occurred because my Power Platform licensing is rated as “Medium or High” performance. The retry pattern was “exponential increasing intervals” and, so, this is why it led to a 40 minute duration for that one Office Scripts Action. 

For Office Scripts, the reliability of that HTTP service is around 99.9% and higher in various Regions across the world. Office Scripts is a Microsoft provided HTTP service and matches uptime reliability scores of other services including Azure.  
So, very rarely will you find a retry actually being necessary. But to capture that .1% or less occurrence, of failure, you can add one retry.  

So, in order to prevent unnecessary retries occurring, in our Flow, what do we do? 

Video : No retries.wmv

So, this explains why our script was attempting so many retries and the 40 minute duration. But how do we mitigate or prevent that error from occurring in the first place? 

Let’s go back to the script. I removed the outer loop, from the Office Script, as shown in this screenshot:

![image](https://user-images.githubusercontent.com/47678539/176802952-db0e7bea-cd29-49f2-af22-adccbec0ee69.png)
 

I, also, added a return value. Each time the script runs it does a check to see if a complete Tour was found. It returns a string value of “Congratulations” if successful. 


I rewrote the associated Flow using the “Do Until” Action as shown here. 
You can see I am capturing or making use of the script return value to check if a success was returned or not. In the advanced settings I have set it as 1600 loops to complete before moving on to the next action in the Flow. 
 
![image](https://user-images.githubusercontent.com/47678539/176802973-d2ec4d28-e311-4c1b-8d1b-cb223ca13df3.png)


However, there is a problem. 

“When using Office Scripts with Power Automate, each user is limited to 1,600 calls to the Run Script action per day. This limit resets at 12:00 AM UTC.” Source: Platform limits and requirements with Office Scripts - Office Scripts | Microsoft Docs

We want to run a Script that just keeps going until it finds a Knight’s Tour. We are only allowed 1600 calls to our script – in Power Automate. So, yes the Flow will run with no errors but we are unlikely to solve the mathematical problem that is the Knights Tour with only 1600 attempts.   

So what do we do? Well, in terms of Power Automate Online that’s as far as we can go. There is nothing more we can do using the Actions available there. But there are still two options open to us. 

The first is to manually click the script button in the Excel file as displayed in your browser. The script will run for as long as your browser is open. However, you, first, need to change the browser’s settings to disable sleeping tabs. Your machine must, also, remain on. Chrome has a 2-hour limit which can be switched off, too.

The second option is to open the Excel file in Excel Desktop and, then, manually click the Script button in the worksheet. The script will continue to run for as long as your workbook is open. 
For this latter option you can use Power Automate Online and Desktop Apps to automate the process. Power Automate Desktop cannot automate the opening of a browser and clicking the Script button. And there is no JavaScript function to start an Office Script, at this time.
 
Power Automate Desktop can be run in in unattended mode and you can, also, use a Virtual machine this way. 

So, let’s have a look at this second option in more detail.

And, finally, let’s add this Desktop Flow to a Cloud Flow to fully automate the running of an Office Script.


Question : Can you trigger an Excel Script by using a Graph REST API? Answer: No.
 



   

