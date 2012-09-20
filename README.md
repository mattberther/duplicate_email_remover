## Origin
Outlook for Mac had been uploading multiple copies of the same message to the Exchange server. At final count, there were approximately 280,000 email messages sitting in my "Archive" folder on the server. This caused tremendous download times for resynchronizing my folders.

There are other tools out there that could purge the duplicates for me, but most did not work on Microsoft Outlook 2010. My attempt to solve this problem is in this project: a C# console application that iterates through all the messages in my "Archive" folder and removes duplicate items.

I looked for tools that could purge the duplicates for me, but had a tough time getting most of them to work on Microsoft Outlook 2010. I set out to try and solve this problem by creating a simple C# app that would iterate through my archive folder and identify and remove duplicate items.

## Method
To remove as many duplicates as possible, messages are parsed two different ways. The first key is created by examining the message id header. The second key is created by concatenating the sender email, subject, and sent time. This technique worked remarkably well. My archive folder now has less than 70,000 messages in it, which means that approximately 75% of the messages in the folder were duplicated.

## Known Issues
Currently, this code is suited to my particular situation: a named "Archive" folder, and a command line app. Offering options to support "dry runs", different folder names, and potentially even an Outlook add-in would be great in future versions.

If you run this with the default Outlook settings, you will receive a warning dialog for every message that Outlook attempts to process. When there are a substantial amount, this can become quite a nuisance. To get around this, you can start Outlook as an administrator, and go to Options on the File menu. Click the Trust Center tab, and then Trust Center Settings. On the programmatic access tab, select "Never warn me about suspicious activity". *PLEASE* make sure to turn this back on after you're done running this program.

## Warranties
Please keep in mind that there are no warranties with the code. It worked well for me; your mileage may vary.
