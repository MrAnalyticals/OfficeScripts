**Analysing Prime Number Theory using Office Scripts

A demo of the "trial by division" method.**

**YouTube Video URL:** https://youtu.be/xgzwwc9dlWs


![image](https://user-images.githubusercontent.com/47678539/219172372-98f143da-81fa-4c73-9d05-f82b76e7cdd2.png)

Office Script URL: https://github.com/MrAnalyticals/OfficeScripts/blob/main/PrimeNumberTrial/PrimeGen.osts

**Script Audio
**
In this video I demonstrate the use of the trial division method to find the nth prime. 
The number n is supplied in cell B1 and that prime is returned in cell B2. 
Here I enter n as 1000 and start the script. It takes 10 seconds. Straight afterwards, I run the script again to find the 10,000th Prime.  It takes around the same duration. Then I run for n = 10000. And again it takes a similar time. 
For n = 1 million the duration is approximately 44 seconds. 
1:16
Now I open Excel Desktop and keep Excel Online open and run the script for n = 1 million. It takes over 4 and a half minutes. 
Next I run the script in Excel desktop but with Excel Online closed. The duration, this time is: just over 3 minutes. So, significantly fast. 
So, next, and this is the final test,  I additionally switch off the autosave toggle keeping Excel Online closed. The duration this time is: around 3 minutes 45 second.
So, in summary the fastest, for the n = 1 million case is using Excel Online with Excel Desktop closed. 
A surprising result, given I have previous experience of Excel Desktop running considerably faster. One possible reason for this is that in this demo I am not writing to or from the workbook at all except once to write the result. That is, it might be the case that Excel Desktop is a faster platform for those scripts that read and write to the workbook. Perhaps you the viewer would like to experiment with this idea. Feel free to use this script for that. 
It can be found at the GitHub repo address that follows.

