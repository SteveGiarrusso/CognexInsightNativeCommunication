# CognexInsightNativeCommunication
This class can be used to easily communicate over Telnet to a Cognex Insight Camera using TCP/IP and native mode communication. Use this to do things like change jobs, set a camera online/offline, and manipulate spreadsheets from a .NET application. You might need to install the Cognex Insight SDK to run it while debugging in VS. Its free and avalible from Cognex. https://support.cognex.com/en-th/downloads/detail/logistics-vision/4160/1033

Add this to your code as a .vb file or you can build it into a DLL file. 

CognexInsightNativeCOM.dll was built of this code
Cognex.Insight comes from Cognex

I built this off of the Cognex Insight DLL. I chose Native/Telnet method of communcation over the insight classes cognex gave due to the ease of tcp/ip, and using native will not kick you off Cognex Insight Explorer. If you use the Cvs classes from cognex, you cant view the spreadsheet in Insight Explorer as that instance will consume the connection you have to the camera.

Im probaly not going to develop this futher but fell free to add more. I built in a good chunk of avalible commands but not all of them. You can view the cognex native commands in the insight explorer help file

I used this to build a multi-cam insight system in a VB winforms

Tested on a Insight 9912 and a ISM1100. Will not work on a IS2000 but should work on any other insight camera
