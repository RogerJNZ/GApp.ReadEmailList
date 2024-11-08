# Read Gmail Email List

Do you want a list of all Gmail emails and their associated size? This Apps Script can be added to a Google Sheet to generate such a list


### Installation Instructions
1. Create a new Google Sheet in Google Drive
   
2. Click on **Extenstions** then **Apps Script**

![image](https://github.com/user-attachments/assets/3a2d6b9b-e0d0-441d-b630-5d9c6a329d63)

3. Paste the **ReadEmailList.js** code into the **Code.JS** section of the Apps Script
![image](https://github.com/user-attachments/assets/aea89069-8964-472d-a11d-12b690900e58)

4. Add **appsscript.json** to reduce permissions required. https://developers.google.com/apps-script/concepts/scopes
   
5. Click **Run**

6. A new sheet will be created called **EmailList** which will be populated with you emails.
Note: The first time you run the script you will be required to give the script access to your Gmail emails
![image](https://github.com/user-attachments/assets/04fcdade-38b6-4405-9a47-388fcdc84137)


### Usage
If you have a large number of emails the script will probably exceed the script time limit before it has read all of your emails.
Simply click Run for the script to continue from the last Email it read.
By default, 50 threads (a thread includes the original email + all the replies) are read before writing to the sheet.

To find the related email in Gmail, copy the text in the Gmail Search column and paste into the Gmail search

 
