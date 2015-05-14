<!--#include file="includes/header.asp"-->
	<script>
		$(document).ready(function(){
			//setTimeout(function(){window.location.reload();},6000);
			$('.addtoPlaylist').click(function(){
				$('#msgArea').css({'display' : 'block'});
			});
			$('.convert').click(function(){
				$('#msgArea2').css({'display' : 'block'});
			});
			$('.theaterMode').click(function(){
				$('body').css('background-color', 'rgba(15, 15, 15, 0.99)');
				$('body').css('background-image', 'none');
			});
		});	
	</script>
	<style>
		.glyphicon-refresh-animate {
		-animation: spin .7s infinite linear;
		-webkit-animation: spin2 .7s infinite linear;
		}

		@-webkit-keyframes spin2 {
			from { -webkit-transform: rotate(0deg);}
			to { -webkit-transform: rotate(360deg);}
		}

		@keyframes spin {
			from { transform: scale(1) rotate(0deg);}
			to { transform: scale(1) rotate(360deg);}
		}
		.vidContainer {
			background: rgba(0, 0, 0, 0.40);
			padding:14px;
			border-radius:10px;
		}
	    .newStyle1
        {
        }
	</style>
<%
'check to see if logged in
if Session("logon")="true" then 'The Big IF Statement - Scroll to bottom to find where it ends

'setting up global variables
'uncomment IF statement if you want to assign different directory to diffent users based on userlevels
	'if Session("userlevel")=2 then
	'	path = Lcase("E:\Videos\") 'root media directory
	'elseif Session("userlevel")=3 then
	'	path = Lcase("E:\MYFILES") 'root media directory
	'else
		path = Lcase("E:\Media Files\")'root media directory
	'end if
	
	totalItemsToShow = 20
	
	playlistPath = Lcase("C:\inetpub\wwwroot\mediabrowser\myplaylist\"&Session("username")&"\")'playlist directory

	if request.querystring("folder")>"" then
		path = path&request.querystring("folder")
		playlistPath  = playlistPath &request.querystring("folder")
	end if

	if request.querystring("myplaylist")="true" then
	'reload page every few seconds to get progress update
		%>
		<script>
				$(document).ready(function(){
					setTimeout(function(){window.location.reload();},8000);
				});
		</script>
		<%
	end if

'Subs and Functions	
	Sub ListFolderContents
		response.write "<div class=""alert alert-info"" role=""alert"" id=""msgArea"" style=""display:none; ""><i class=""icon-refresh glyphicon-refresh-animate""></i>&nbsp;Please wait. Adding file to playlist...</div>"
		response.write "<div class=""alert alert-info"" role=""alert"" id=""msgArea2"" style=""display:none; ""><i class=""icon-refresh glyphicon-refresh-animate""></i>&nbsp;Please wait. Preparing to convert file...</div>"
		Set fs = server.CreateObject("Scripting.FileSystemObject")
		'Error Checking
		if request.querystring("myplaylist")="true" then

			If fs.FolderExists(playlistPath)=true Then
				Set MyFolder = fs.GetFolder(playlistPath)
			else	
				response.write "<div class=""alert alert-error""><a class=""close"" data-dismiss=""alert"" href=""?dismiss=1"">×</a>Playlist Not Found.</div>"&vbCrLf
				response.end()
			end if
		else
			If fs.FolderExists(path)=true Then
				Set MyFolder = fs.GetFolder(path)
			else	
				response.write "<div class=""alert alert-error""><a class=""close"" data-dismiss=""alert"" href=""?dismiss=1"">×</a>Directory Not Found.</div>"&vbCrLf
				response.end()
			end if
		end if
		
		'''Set session variable for mobileFiles
		if request.querystring("mobileFiles")="true" then
			session("mobileFiles")="true"
		elseif request.querystring("mobileFiles")="false" then
			session("mobileFiles")="false"
		end if
		
		response.write "<div class=""row"">"&vbCrLf 
		response.write "<div class=""span12"">"&vbCrLf
		
		response.write "<table class=""table table-bordered table-striped""><tbody>"&vbCrLf 
		response.write "<tr>"&vbCrLf
		
		if request.querystring("myplaylist")="true" then 
		response.write "<td colspan=""3""><h2>My Playlist</h2></td>"
		end if
		
		'''Toggle Show Mobile link based on session variable value
		if NOT request.querystring("myplaylist")="true" then 
			if session("mobileFiles")="true" then
					response.write "<span style=""float:right;""><a href=""?folder="&request.querystring("folder")&"&page="&request.querystring("page")&"&mobileFiles=false"">Show All Files</a></span>"
			elseif session("mobileFiles")="false" then
					response.write "<span style=""float:right;""><a href=""?folder="&request.querystring("folder")&"&page="&request.querystring("page")&"&mobileFiles=true"">Only Show Mobile Files</a></span>"
			end if
		end if
		
		response.write "</td>"
		response.write "</tr>"&vbCrLf
		'List Folders
			if request.querystring("folder")>"" then
				response.write "<tr>"&vbCrLf
				response.write "<td class=""fileType""><i class=""icon-arrow-left""></i></td>"&vbCrLf
				response.write "<td><a href=""#"" onclick=""history.back(); return false;"" style=""word-wrap: break-word;""> "&request.querystring("folder")&"</a></td>"&vbCrLf	
				response.write "<td class=""fileType""></td>"&vbCrLf
				response.write "</tr>"&vbCrLf
			end if

			For Each x in Myfolder.subfolders
				if instr(x.Name,".actors") then 'skip these folders
					response.write ""
				else
					if request.querystring("folder")>"" then
						response.write "<tr>"
						response.write "<td class=""fileType""><i class=""icon-black icon-folder-close""></i></td><td><a href=""?folder="&request.querystring("folder")&"/"&x.Name&"&page=1"" style=""word-wrap: break-word;""> "&x.Name&"</a></td>"&vbCrLf
						response.write "<td></td>"&vbCrLf
						response.write "</tr>"&vbCrLf
					else
						response.write "<tr>"
						response.write "<td class=""fileType""><i class=""icon-black icon-folder-close""></i></td><td><a href=""?folder="&x.Name&"&page=1"" style=""word-wrap: break-word;""> "&x.Name&"</a></td>"&vbCrLf
						response.write "<td></td>"&vbCrLf
						response.write "</tr>"&vbCrLf
					end if
				end if
			Next

		'List Files
		count=0
		tcount=0
		
			For Each t in Myfolder.files
				if session("mobileFiles")="true" then
					fileTypeLimit1=Right(t.Name,4) = ".mp4" OR Right(t.Name,4) = ".mp3"
				else
					fileTypeLimit1=Right(t.Name,4) = ".mp4" OR Right(t.Name,4) = ".avi" OR Right(t.Name,4)=".mov" OR Right(t.Name,4)=".wmv" OR Right(t.Name,4) = ".mp3" OR Right(t.Name,4)=".mkv" 
				end if
				if NOT (fileTypeLimit1) then
					response.write ""
				else
					tcount=tcount+1
				end if
			Next
			
			For Each x in Myfolder.files
				
				'Skip all files Except for these
				if session("mobileFiles")="true" then
					fileTypeLimit=Right(x.Name,4) = ".mp4" OR Right(x.Name,4) = ".mp3"
				else
					fileTypeLimit=Right(x.Name,4) = ".mp4" OR Right(x.Name,4) = ".avi" OR Right(x.Name,4)=".mov" OR Right(x.Name,4)=".wmv" OR Right(x.Name,4) = ".mp3" OR Right(x.Name,4)=".mkv" 
				end if
				
				if NOT (fileTypeLimit) then
					response.write ""
				else
				
					count=count+1
					
					'set icons for file types
					if (Right(x.Name,4)=".mp4" OR Right(x.Name,4)=".avi" OR Right(x.Name,4)=".mov" OR Right(x.Name,4)=".wmv" OR Right(x.Name,4)=".mkv" ) then
						fileIcon = "icon-black icon-film"
						fileTitle = "Video File"
						fileDevice = ""
					elseif Right(x.Name,4)=".mp3" then
						fileIcon = "icon-black icon-music"
						fileTitle = "Music File"
						fileDevice = ""
					else
						fileIcon = "icon-black icon-file"
						fileTitle = "Media File"
						fileDevice = ""
					end if
					
					if request.querystring("folder")>"" then

						response.write "<tr>"&vbCrLf
						response.write "<td class=""fileType""><i class="""&fileIcon&""" title="""&fileTitle&"""></i></td>"&vbCrLf
						response.write "<td><a href=""?playlist=add&file="&request.querystring("folder")&"/"&x.Name&""" title=""Add to Playlist"" style=""word-wrap: break-word;"" class=""addtoPlaylist""> "&x.Name&"</a></td>"&vbCrLf
						response.write "<td>"&fileDevice&"</td>"&vbCrLf
						response.write "</tr>"&vbCrLf

					elseif request.querystring("myplaylist")="true" then 'show playlist files
						
						logFile=playlistPath+x.Name+".log"
						
						'''''''Read the Log file to get the file status
						Set objFSO = server.CreateObject("Scripting.FileSystemObject")
						if objFSO.FileExists(logFile) then
							Set objFile = objFSO.OpenTextFile(logFile, 1)
								readLog = objFile.ReadAll
							objFile.Close
							Set objFSO=Nothing
							Set objFile=Nothing
						end if
						
						set fs=Server.CreateObject("Scripting.FileSystemObject")
						set f=fs.GetFile(playlistPath+x.Name)
							activeFileSize=f.Size
						set f=nothing
						set fs=nothing
						
						if (inStr(readLog,"P-Frames") OR inStr(readLog,"frame P")) then 'find this unique word in log file
							fileStatus=""
							'fileSize=""
						elseif Right(x.Name,4) = ".mp3"  then
							fileStatus=""
						else
							fileStatus="1"
							fileIcon="icon-refresh glyphicon-refresh-animate"
							fileTitle = "Encoding"

							if activeFileSize=0 then
								fileSize="(Failed)"
							else
								if activeFileSize>1500000000 then '1.5GB
									fileSize="1.5gb"
								elseif activeFileSize>1000000000 then '1GB
									fileSize="1gb"
								elseif activeFileSize>900000000 then '900mb
									fileSize="900mb"
								elseif activeFileSize>800000000 then '800mb
									fileSize="800mb"
								elseif activeFileSize>700000000 then '700mb
									fileSize="700mb"
								elseif activeFileSize>600000000 then '600mb
									fileSize="600mb"
								elseif activeFileSize>500000000 then '500mb
									fileSize="500mb"
								elseif activeFileSize>400000000 then '400mb
									fileSize="400mb"
								elseif activeFileSize>300000000 then '300mb
									fileSize="300mb"
								elseif activeFileSize>200000000 then '200mb
									fileSize="200mb"
								elseif activeFileSize>100000000 then '100mb
									fileSize="100mb"
								elseif activeFileSize>50000000 then '50mb
									fileSize="50mb"
								elseif activeFileSize>25000000 then '25mb
									fileSize="25mb"
								else
									fileSize="10mb"
								end if
								fileSize="(Encoding..."&fileSize&")"	
							end if
							
						end if
						
						response.write "<tr>"&vbCrLf
						response.write "<td class=""fileType""><i class="""&fileIcon&""" title="""&fileTitle&"""></i></td>"&vbCrLf

						if fileStatus="1" then
							response.write "<td style=""word-wrap: break-word;"" >"&x.Name&"&nbsp;<i>"&fileSize&"</i></td>"&vbCrLf
						else
							response.write "<td style=""word-wrap: break-word;"" ><a href=""?play="&x.Name&""" title=""Play File"" style=""word-wrap: break-word;"">"&x.Name&"</a></td>"&vbCrLf
						end if

						response.write "<td><a href=""?remove="&x.Name&""" title=""Remove File"" style=""word-wrap: break-word;""><i class=""icon-black icon-trash""></i></a></td>"&vbCrLf
						response.write "</tr>"&vbCrLf
					else
						response.write "<div class=""alert alert-error""><a class=""close"" data-dismiss=""alert"" href=""?dismiss=1"">×</a>No Files Found.</div>"&vbCrLf
					end if
					
					'get item and page counts
					PageNum = request.querystring("page")
					pageSet = (PageNum -1) * totalItemsToShow + totalItemsToShow 
					pageNext = "<span style=""float:left;"">Showing: ("&pageSet&" of "&tcount&") <a href=""?folder="&request.querystring("folder")&"&page="&pagenum+1&""">Next "&totalItemsToShow&" Items</a> <i class=""icon-arrow-right""></i></span>"
					
					if NOT count = (PageNum -1) * totalItemsToShow + totalItemsToShow then
						count=count
						pageSet = "All"
						pageNext = "<span style=""float:left;"">Showing: ("&pageSet&" of "&tcount&")</span>"
					else
						Exit For
					end if
					
				end if
			Next
			
			if request.querystring("folder")>"" then
				response.write pageNext
			end if
			
			'response.write count
			Call InQueue
			if (request.querystring("myplaylist")="true" AND count=<1) then
				response.write "<div class=""alert alert-error""><a class=""close"" data-dismiss=""alert"" href=""?dismiss=1"">×</a>No Files in Playlist.</div>"&vbCrLf
			end if
		
		response.write "</tbody></table>"&vbCrLf 
		response.write "</div>"
		response.write "</div>"
		Set fs = Nothing
	End Sub

	Sub InQueue
		if request.querystring("myplaylist")="true" then
			SQL="SELECT filename, filestatus, userid, id FROM filequeue WHERE userid="&session("userid")&""
			set rs=dbconn.execute(SQL)
			if NOT rs.eof then
				response.write "<div class=""row"">"&vbCrLf 
				response.write "<div class=""span12"">"&vbCrLf
				response.write "<table class=""table table-bordered table-striped""><tbody>"&vbCrLf 
				response.write "<tr>"&vbCrLf
				response.write "<td colspan=""3""><h2>Queue</h2></td>"
				
					do while not rs.eof
						response.write "<tr>"&vbCrLf
						response.write "<td style=""word-wrap: break-word;"" >"&rs("filename")&"&nbsp;<i>(In Queue)</i></td>"&vbCrLf
						response.write "<td><a href=""?myplaylist=true&convertNow="&rs("filename")&""" title=""Convert Now"" style=""word-wrap: break-word;"" class=""convert""><i class=""icon-black icon-cog""></i></a></td>"&vbCrLf
						response.write "<td><a href=""?remove="&rs("filename")&""" title=""Remove File"" style=""word-wrap: break-word;""><i class=""icon-black icon-trash""></i></a></td>"&vbCrLf
						response.write "</tr>"&vbCrLf
						rs.movenext
					loop
					
				response.write "</tbody></table>"&vbCrLf 
				response.write "</div>"
				response.write "</div>"
				rs.close
			end if
			if request.querystring("convertNow")>"" then
				vidSrc=""&server.HTMLEncode(path+request.querystring("convertNow"))&""
				vidDest=server.HTMLEncode(replace(playlistPath+request.querystring("convertNow"),Right(playlistPath+request.querystring("convertNow"),4),".mp4"))
				Destfilename=Split(vidDest,"\",len(vidDest))
				for each x in Destfilename
					Destfilearray=Array(x & "")
				next
				vidLog=Server.HTMLEncode(playlistPath+Destfilearray(0)&".log")
				Set objExecutor = Server.CreateObject("ASPExec.Execute")
				objExecutor.Application = "cmd.exe" 
				'objExecutor.Parameters = "/c c:\ffmpeg\bin\ffmpeg.exe -i """&vidSrc&""" -r 29 -s 720x480 -vcodec libx264 -y """&playlistPath+Destfilearray(0)&""" 2>"""&vidLog&""""
				'objExecutor.Parameters = "/c c:\ffmpeg\bin\ffmpeg.exe -i """&vidSrc&""" -s 720x480 -vcodec libx264 -pix_fmt yuv420p -profile:v baseline -preset slower -crf 18 -y """&playlistPath+Destfilearray(0)&""" 2>"""&vidLog&""""
				objExecutor.Parameters = "/c c:\ffmpeg\bin\ffmpeg.exe -i """&vidSrc&""" -c:v libx264 -crf 19 -preset slow -c:a aac -strict experimental -b:a 192k -ac 2 """&playlistPath+Destfilearray(0)&""" 2>"""&vidLog&""""
				objExecutor.ShowWindow = false
				sResult = objExecutor.ExecuteWinApp
				Set objExecutor=Nothing
				
				removefile=replace(request.querystring("convertNow"),"\","\\")
				
				SQL2="DELETE FROM filequeue WHERE filename="""&removefile&""" AND userid="&Session("userid")&""
				dbconn.Execute(SQL2)
				
				EndTime = Now() + (4 / (24 * 60* 60))
				
				Do While Now() < EndTime 
					'Do nothing Wait 4 Seconds before going to next step
				Loop
				
				response.redirect("?myplaylist=true")
			end if
		end if
	End Sub
	
	Sub TaskKill
		Set objExecutor = Server.CreateObject("ASPExec.Execute")
		objExecutor.Application = "cmd.exe" 
		objExecutor.Parameters = "/c taskkill /F /IM ffmpeg.exe"
		objExecutor.ShowWindow = false
		sResult = objExecutor.ExecuteWinApp
		Set objExecutor=Nothing
		response.redirect ("?myplaylist=true")
	End Sub

	Sub AddtoPlaylist
	
		On Error Resume Next '''there will be errors. so skip them!
		'Copy a file to another folder
		filepath=request.querystring("file")
		filepath=replace(filepath,"/","\")
		filename=Split(filepath,"\",4)
		TVShowFanArt=Split(filepath,"\",3)
		
		for each x in filename
			filearray=Array(x & "")
		next
		
		for each x in TVShowFanArt
			filearray2=Array(x & "")
		next
		
			Set fs=Server.CreateObject("Scripting.FileSystemObject")
				If fs.FolderExists(playlistPath)=false Then
					Set CreatePlaylistFolder = fs.CreateFolder(playlistPath)''create playlist folder if it doesn't exist
				end if
				
				if Right(playlistPath+fileArray(0),4)=".mp4" then
				
					set dummyfile=fs.CreateTextFile(playlistPath+fileArray(0)+".log")''''create dummy log file
					dummyfile.WriteLine("frame P : P-Frames : "&playlistPath&fileArray(0))'''write this line to the log file
					dummyfile.close
					set dummyfile=nothing

					fs.CopyFile path+filepath,playlistPath+fileArray(0)'''copy the video file
					
				elseif  Right(playlistPath+fileArray(0),4)=".mp3" then
				
					fs.CopyFile path+filepath,playlistPath+fileArray(0)'''copy the mp3 file
					
				else ''''if not MP4
					
					vidSrc=path+filepath
					vidDest=server.HTMLEncode(replace(playlistPath+fileArray(0),Right(playlistPath+fileArray(0),4),".mp4"))
					vidLog=Server.HTMLEncode(vidDest&".log")

					SQLfilepath=replace(filepath,"\","\\")
					SQL="INSERT INTO filequeue (filename,filestatus,userid) VALUES ("""&SQLfilepath&""",0,"&session("userid")&")"
					dbconn.Execute(SQL)

				end if
				
				EndTime = Now() + (4 / (24 * 60* 60))
				Do While Now() < EndTime 
				'Do nothing. Wait 4 Seconds before going to next step
				Loop

				fs.CopyFile replace(path+filepath,Right(path+filepath,4),"-thumb.jpg"),replace(playlistPath+fileArray(0),Right(playlistPath+fileArray(0),4),"-thumb.jpg") '''''copy TV Show thumbnail
				fs.CopyFile replace(path+filepath,Right(path+filepath,4),"-poster.jpg"),replace(playlistPath+fileArray(0),Right(playlistPath+fileArray(0),4),"-poster.jpg") '''''copy Movie poster
				fs.CopyFile replace(path+filepath,Right(path+filepath,4),"-fanart.jpg"),replace(playlistPath+fileArray(0),Right(playlistPath+fileArray(0),4),"-fanart.jpg") '''''copy Movie fanart 
				fs.CopyFile replace(path+filepath,fileArray(0),"fanart.jpg"),playlistPath+"fanart.jpg" '''''copy TV Show fanart
				fs.MoveFile playlistPath+"fanart.jpg",replace(playlistPath+fileArray(0),Right(playlistPath+fileArray(0),4),"-fanart.jpg") 'rename fanart image
			Set fs = Nothing
			
		On Error GoTo 0
		response.redirect ("?myplaylist=true")
	End Sub
	
	Sub PlayFile
		webPlayListPath = "myplaylist/"&Session("username")&"/"&request.querystring("play")
		bgImgStr=request.querystring("folder")&"/"&replace(request.querystring("play"),Right(request.querystring("play"),4),"-fanart.jpg")
		bgImgStr=Server.HTMLEncode("myplaylist/"&Session("username")&bgImgStr)
		'response.write bgImgStr
		thumbImg=request.querystring("folder")&"/"&replace(request.querystring("play"),Right(request.querystring("play"),4),"-thumb.jpg")
		thumbImg="myplaylist/"&Session("username")&thumbImg
		posterImg=request.querystring("folder")&"/"&replace(request.querystring("play"),Right(request.querystring("play"),4),"-poster.jpg")
		posterImg="myplaylist/"&Session("username")&posterImg
		
		response.write "<div class=""vidContainer"" align=""center"">"
		response.write "<span style=""word-wrap: break-word; display:block; color:#fff; text-align:left;""><i>Playing file:</i> "&request.querystring("play")&"</span><br>"
		
		'check if it is a TV Show or Movie image
		if inStr(bgImgStr,".S0") then
			thImg=thumbImg
		else
			thImg=posterImg
		end if

		%>
			<script>
				$(document).ready(function(){
					$('body').css('background', ' #333 url("<%=bgImgStr%>") repeat-y top center /cover');
				});
			</script>
		<%
		if Right(webPlayListPath,4) = ".mp4" then
			response.write "<video width=""100%"" controls poster="""&thImg&""" style=""border:solid 2px #ccc; background: rgba(0, 0, 0, 0.80);"">"
			response.write "<source src="""&webPlayListPath &""" type=""video/mp4"">"
			response.write "</video>"
		elseif Right(webPlayListPath,4) = ".mp3" then
			response.write "<audio controls>"
			response.write "<source src="""&webPlayListPath &"""  type=""audio/mpeg"">"
			response.write "</audio>"
		else
			response.write "<div class=""alert alert-error""><a class=""close"" data-dismiss=""alert"" href=""?dismiss=1"">×</a>This file can not be played in your browser. <a href="""&webPlayListPath&""" download="""&webPlayListPath&""" title=""Download File"" style=""word-wrap: break-word; color:#333;"">Download the file</a></div>"&vbCrLf
		end if
		response.write "<div style=""clear:both;""></div>"
		response.write "<div style=""margin-top:20px; float:left; width:50%; text-align:left; color:#fff;""><i class=""icon-black icon-download"" style=""margin-right:6px;""></i><a href="""&webPlayListPath&""" download="""&webPlayListPath&""" title=""Download File"" style=""word-wrap: break-word; color:#fff;"">Download File</a></div>"
		response.write "<div style=""margin-top:20px; float:right; width:50%; text-align:right;""><i class=""icon-black icon-film"" style=""margin-right:6px;""></i><a href=""#""  class=""theaterMode"" title=""Theater Mode"" style=""word-wrap: break-word; color:#fff;"">Theater Mode</a></div>"
		response.write "<div style=""clear:both;""></div>"
		response.write "</div>"
	End Sub

	Sub DeleteFile
		On Error Resume Next '''there will be errors. so skip them!
		''''Remove item from DB if exists
		removefile=replace(request.querystring("remove"),"\","\\")
		SQL2="DELETE FROM filequeue WHERE filename="""&removefile&""" AND userid="&Session("userid")&""
		dbconn.Execute(SQL2)
		
		filepath = request.querystring("remove")
		filepath = replace(filepath,"/","\")
		filename=Split(filepath,"\",2)
		
		for each x in filename
			filearray=Array(x & "")
		next
		
		Set fs=Server.CreateObject("Scripting.FileSystemObject")
		'If fs.FileExists(playlistPath+fileArray(0)) then

			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set fs=fs.GetFile(playlistPath+fileArray(0))
			fs.Delete
			
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set fs=fs.GetFile(replace(playlistPath+fileArray(0),Right(playlistPath+fileArray(0),4),"-fanart.jpg"))
			fs.Delete
			
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set fs=fs.GetFile(replace(playlistPath+fileArray(0),Right(playlistPath+fileArray(0),4),"-thumb.jpg"))
			fs.Delete
			
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set fs=fs.GetFile(replace(playlistPath+fileArray(0),Right(playlistPath+fileArray(0),4),"-poster.jpg"))
			fs.Delete
			
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set fs=fs.GetFile(playlistPath+fileArray(0)+".log")
			fs.Delete
			
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set fs=fs.GetFile(playlistPath+".DS_Store")
			fs.Delete
			
			set fs=Server.CreateObject("Scripting.FileSystemObject")
			set fs=fs.GetFile(playlistPath+"Thumbs.db")
			fs.Delete

		Set fs = Nothing

		On Error GoTo 0
		response.redirect ("?myplaylist=true")
	End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Call Subs based on what is in the querystring
	if (Request.Querystring("playlist") = "add" AND Session("logon")="true") then
		AddtoPlaylist
	elseif (Request.Querystring("play")  > "" AND Session("logon")="true") then
		PlayFile
	elseif (Request.Querystring("remove")  > "" AND Session("logon")="true") then
		DeleteFile
	elseif (Request.Querystring("taskkill")  > "" AND Session("logon")="true") then
		TaskKill
	else
		ListFolderContents
	end if
else
	response.redirect("default.asp") 'log out 
end if
%>
<!--#include file="includes/footer.asp"-->