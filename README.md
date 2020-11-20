# Current Process

The contractors who developed our version of Salesforce designed one of our most important reports due to its complexity. I can only run this report in Salesforce Einstein and I am not able to change the logic or even know how the information is being pulled. 

The problem is the report does not always run correctly and then we have to put in a ticket for them to fix the issue.

After running the report, then I count the unique IDs and then summarize the information in an email.

# New Process
The main issue I was having with Salesforce and not being able to build this report my self was that I had to rely on using the front end report builder of Salesforce. However, this report needs information from different places that can't be connected like in a normal database. The only tools that I had available to me was Microsoft Excel and Access. 

Luckily, with the help of the simple salesforce module, I was able to directly write SOQL queries to Salesforce and then use the Python Pandas module for the complex aspects of combining and cleaning the data. 

The end result is a report that looked just like the report the contractors created, but now I have full control of what is going on and if I need to change something I can. 

This will be part of a main program that will run all of the weekly reports for me. 