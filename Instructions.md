This is a single page PHP application. the page will have following tabs.
1. Spreadsheet
2. Send QR
3. Scan QR
4. Attendance
5. Music
Spreadsheet tab will display a form which will browse to select a spreadsheet. it will be usually placed in ame folder of application. This spreadsheet will be base of all application data.
show ful error message below load spreadsheet button
in this tab also add two text bos with Whatsapp username and Password to fill username and password to use in Whatsapp API.
the tab Send QR will  display all teh records from spread sheet in tabular form with check box at first column. Fields SN, Name, Designation and Institution, Whatsapp will be displayed in table.
there will be one button 'Create and Send QR' will  create a QR code with value in SN column for Selected Row in table and will send this QR to Coresponding column with text 'Whatsapp' using  
	API FOR TEXT WA WITH VARIABLES/PARAMETERS + IMAGE/VIDEO
http://bhashsms.com/api/sendmsg.php?user=InnocentWTS&pass=********&sender=BUZWAP&phone=Mobile No&text=TEMPLATENAME&priority=wa&stype=normal&Params=param1,param2&htype=image&url=https://i.ibb.co/9w4vXVY/Whats-App-Image-2022-07-26-at-2-57-21-PM.jpg
teh QR generated will be saved in folder QR with name value of SN column ().jpg format) before sending. The API picks QR from the folder.
TEMPLATENAME will be event_invitation.
third tab Scan QR will use teh linked camera with teh compuer on which it is running . it will scan the QR from Mobile phone. it will match the number in Qr with column in spreadsheet with text 'SN' in first row.
It will make an announcement 
"Welcome" + Name column from Spreadsheet +Designation from spreasheet _ Institution from Spreadsheet (All of from same row of SN scaned.)
the Selecetd data will be displayed in table in 'Attendance' tab.
The Same data will be added in new row in a spreadsheet named 'Attendance.xlsx'in applicatin folder.
Tab music will play music files available in application folder one by one in ramdom order. When Scan QR will make announcement, the music will be paused and after announcement, it will resume again.
The default tab will be Scan QR. Make it firts tab too.
Saple Spreadsheet is in folder named invitations.xlsx