# Information system for theater workshop 

This project allow you how to take responses from a Google Form and automatically convert them into a Google Slides presentation. No manual copying and pasting required—each response can be dynamically added to slides using Google Apps Script.

# Getting started.

First, created a google form. After you created it, go the "response" and select "view in sheets". This sheet will storing all the data of the form reponse. 

Go the the sheet you just created and right-click extention and select Google App script

![image](https://github.com/user-attachments/assets/0bfcead9-c833-4867-9386-df9fffa5976a)

My code is for my school theater. It allows user the use form to report the status of different equipments in the theater. 

Here is a example of the reponse:
<img width="1347" alt="Screenshot 2024-09-27 at 10 58 40" src="https://github.com/user-attachments/assets/b3095181-f5d3-4a22-a6c3-e51edb34b987">



# Link the Google Slides 

Create a Google Slides, and copy the url of the slides. 

Go the code, and replace the url on the thrid line.

```

const presentation = SlidesApp.openByUrl("(Replace will your Google slides url)");

```
# Add a trigger to the code
Go to the google Appscript and select triggers in the left.
Select create a new trigger

<img width="618" alt="Screenshot 2024-10-07 at 10 13 51" src="https://github.com/user-attachments/assets/957a3996-c6a6-423b-9014-297fe97c0005">

This made the code run automatically once a form got submited

# Submit a form
There is a example: reponse
<img width="1357" alt="Screenshot 2024-10-07 at 10 11 10" src="https://github.com/user-attachments/assets/35a15d16-1374-463b-aae8-59f0ceb33b46">
The code will generate a google slides based on the reponse. 
<img width="1089" alt="Screenshot 2024-10-07 at 10 17 40" src="https://github.com/user-attachments/assets/01e2b864-c090-49d3-9b55-caa8f706a829">

# Contact info
If you have any question. Feel free to email me or ask a question on this page.

My email is jackfang0601@gmail.com

