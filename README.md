<div align="center">

## Directory Listing


</div>

### Description

To dynamically show contents in a directory (any one you want)
 
### More Info
 
You would need to get the pictures you would want for each type of file in your directory. If you have only one type just define it in the code below to shorten the script.

If you would like to see a working copy, please go to http://www15.brinkster.com/infinityprod/index.asp

Contents of a directory with the following info: name of the file/subdirectory, link to the file/subdirectory, images for the type of file

None that I am aware of


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kaustav Acharya](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kaustav-acharya.md)
**Level**          |Beginner
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kaustav-acharya-directory-listing__4-6392/archive/master.zip)

### API Declarations

If you think this script is good, please rate it appropriately. I would really like that! :D


### Source Code

```
<!DOCTYPE html public "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Directory listing of MP3s</TITLE>
</HEAD>
<BODY bgcolor="#000000" text="#400080" link="#400080" vlink="#FFFF80">
<%
Dim dirname, mypath, fso, folder, filez, FileCount
dirname = "../mp3s/"
mypath = "../mp3s/"
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(server.mappath(dirname))
Set filez = folder.Files
FileCount = folder.Files.Count
' This function takes a filename and returns the appropriate image for
' that file type based on it's extension. If you pass it "dir", it assumes
' that the corresponding item is a directory and shows the folder icon.
Function ShowImageForType(strName)
	Dim strTemp
	' Set our working string to the one passed in
	strTemp = strName
	' If it's not a directory, get the extension and set it to strTemp
	' If it is a directory, then we already have the correct value
	If strTemp <> "dir" Then
		strTemp = LCase(Right(strTemp, Len(strTemp) - InStrRev(strTemp, ".", -1, 1)))
	End If
	response.write strTemp
	' Set the part of the image file name that's unique to the type of file
	' to it's correct value and set this to strTemp. (yet another use of it!)
	Select Case strTemp
		Case "mp3"
			strTemp = "MP3"
		Case "mp2"
			strTemp = "MP3"
		Case "wav"
			strTemp = "MP3"
		Case "aiff"
			strTemp = "MP3"
		Case "html"
			strTemp = "htm"
		Case "m3u"
			strTemp = "MP3"
		End Select
	' All our logic is done... build the IMG Tag for display to the browser
	' Place it into... where else... strTemp!
	' My images are all GIFs and all start with "dir_" for my own sanity.
	' They end with one of the values set in the select statement above.
	strTemp = "<IMG SRC=""images/dir_" & strTemp & ".gif"" WIDTH=16 HEIGHT=16 BORDER=0>"
	' Set return value and exit function
	ShowImageForType = strTemp
End Function
'That's it for functions on this one!
%>
<%' Now to the Runtime code:
Dim strPath 'Path of directory to show
Dim objFSO 'FileSystemObject variable
Dim objFolder 'Folder variable
Dim objItem 'Variable used to loop through the contents of the folder
' You could just as easily read this from some sort of input, but I don't
' need you guys and gals roaming around our server so I've hard coded it to
' a directory I set up to illustrate the sample.
' NOTE: As currently implemented, this needs to end with the /
strPath = mypath
' Create our FSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
' Get a handle on our folder
Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))
' Show a little description line and the title row of our table
%>
<CENTER><FONT face="Arial" color="#9690CC">
<%= FileCount %> Contents of <B><I>
<%= strPath %></I> Not Counting Any Folders</B></CENTER></FONT><BR><BR>
<TABLE align="center" border="5" bordercolor="#010194" cellspacing="0" cellpadding="2">
	<TR bgcolor="#010194">
		<TD align="CENTER"><FONT face="Arial" color="#FFFFFF"><B>File Name:</B></FONT>
		</TD>
		<TD align="CENTER"><FONT face="Arial" color="#FFFFFF"><B>File Size (MB):</B></FONT>
		</TD>
		<TD align="CENTER"><FONT face="Arial" color="#FFFFFF"><B>Date Created:</B></FONT>
		</TD>
		<TD align="CENTER"><FONT face="Arial" color="#FFFFFF"><B>File Type:</B></FONT>
		</TD>
	</TR>
	<%
' First I deal with any subdirectories. I just display them and when you
' click you go to them via plain HTTP. You might want to loop them back
' through this file once you've set it up to take a path as input.
For Each objItem In objFolder.SubFolders
	' Deal with the bad VTI's that keep giving our visitors 404's
	If InStr(1, objItem, "_vti", 1) = 0 Then
	%>
	<TR bgcolor="#87BCFE">
		<TD align="left">
			<%= ShowImageForType("dir") %> <A href="<%= strPath & objItem.Name %>">
			<%= objItem.Name %></A>
		</TD>
		<TD align="right">
			<%= objItem.Size/1000000 'I used this to display the file size in MB, you can set it to default
			%>
		</TD>
		<TD align="left">
			<%= objItem.DateCreated 'date of creation
			%>
		</TD>
		<TD align="left">
			<%= objItem.Type 'the type of file
			%>
		</TD>
	</TR>
	<%
	End If
Next 'objItem
' Now that I've done the SubFolders, do the files!
For Each objItem In objFolder.Files
%>
	<TR bgcolor="#87BCFE">
		<TD align="left">
			<%= ShowImageForType(objItem.Name) %> <A href="<%= strPath & objItem.Name %>">
			<%= objItem.Name %></A>
		</TD>
		<TD align="right">
			<%= objItem.Size/1000000 %>
		</TD>
		<TD align="left">
			<%= objItem.DateCreated %>
		</TD>
		<TD align="left">
			<%= objItem.Type %>
		</TD>
	</TR>
	<%
Next 'objItem
Set objItem = Nothing
Set objFolder = Nothing
' All done! Kill off the object variables.
Set objFSO = Nothing
%>
</TABLE>
</BODY>
</HTML>
```

