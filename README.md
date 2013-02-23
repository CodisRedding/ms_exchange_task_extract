"# Exchange 2010 task extraction" 

1. Select the version of Exchange
2. Enter Exchange user credentials. 
This user must have impersonation permissions for all users that will have tasks extracted.)
(http://msdn.microsoft.com/en-us/library/exchange/bb204095(v=exchg.140).aspx)
3. If your running this application and you connected to the same network, you can attempt
to auto discover the exchange server. by checking Auto Discovery. (note* this can take a couple minutes). 
If you know the full domain of the exchange server uncheck the Auto Discovery checkbox and enter it in the 
server url textbox. This, by default is in the form: https:<you full domain url>/EWS/Exchange.asmx
4. Click the start button to begin the step by step process.

This application expects a CSV file with a list of users to query tasks for in the format

file ext: .csv
content:
Email
some@email1.com
some@email2.com
some@email3.com
some@email4.com

take note of the csv header column 'Email'
