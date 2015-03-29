# mediabrowser
ASP - Media Browser, Browse and Watch media located on your web server, from anywhere.

ASP Media Browser uses ASP's file system objects to allow you to browse and play video, music files right in the web app. Each user has their own Playlist. Toggle between Mobile Only files and All Media files. 

Requires: Windows Server, IIS6 or greater, ASP enabled in IIS, ASPExec.Execute component, MySQL, MIME types set for MP4 files, FFMPEG to convert files. 

Set the default media directory path (E:\MyMedia\) and playlist directory path (C:\inetpub\wwwroot\...) in filemgr.asp.

Import the SQL dump file.
Create an Admin user. 
You can also set multiple directories and assign them to different users based on userlevel. Additional coding is required to achieve this.

Install FFMPEG (c:\ffmpeg\bin\ffmpeg.exe) in order to convert files to a "streamable" format (MP4). Users can click Convert and wait until the file is ready to play.

# Screenshots
![screenshot01](https://github.com/teklynk/mediabrowser/blob/master/mediabrowser_01.png)
![screenshot02](https://github.com/teklynk/mediabrowser/blob/master/mediabrowser_02.png)
![screenshot03](https://github.com/teklynk/mediabrowser/blob/master/mediabrowser_03.png)
![screenshot04](https://github.com/teklynk/mediabrowser/blob/master/mediabrowser_04.png)
