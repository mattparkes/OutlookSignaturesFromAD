'***************************************************************************************************
'*                       OutlookSigGenerator.vbs                                                   *
'*                       Written By: Jaeden Cook, David Phelan & Matt Parkes                       *
'*                                                                                                 *
'*                       Version 3.0.0                                                             *
'*                       Original Date: 4/8/2011                                                   *
'*                                                                                                 *
'*  This script gets user info from Active Directory and uses that information to create Outlook   *
'*  Signature files in Text, HTML, and RTF formats.   										                         *
'*  Office Information is retrieved from the User's OU (set using ADSI Edit).					             *
'***************************************************************************************************

Option Explicit

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell") 
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objSysInfo
Set objSysInfo = CreateObject("ADSystemInfo")
Dim objUserQuery
objUserQuery = "LDAP://" & objSysInfo.Username
Dim objUser
Set objUser = GetObject(objUserQuery)
Dim objLocation
Set objLocation = getObject(objUser.Parent)

'=====================EDIT-HERE===============================================
Dim strCompanyName, strDisclaimer, strWebAddress
strCompanyName = "Company Name"
strWebAddress = "www.companyurl.com" 
strDisclaimer = "Disclaimer: This communication including any attachments may contain information that is either confidential or otherwise protected from disclosure and is intended"&_
      "solely for the use of the intended recipient. If you are not the intended recipient please immediately notify the sender by e-mail and delete the original transmission and"&_
      "its contents. Any unauthorised use, dissemination, forwarding, printing, or copying of this communication including any file attachments is prohibited. The recipient"&_
      "should check this email and any attachments for viruses and other defects. The Company disclaims any liability for loss or damage arising in any way from this"&_
      "communication including any file attachments."
Dim strServerSignatureFolder, strLocalSignatureFolder
strServerSignatureFolder = "\\fileserver\signatures\"
strLocalSignatureFolder = objShell.ExpandEnvironmentStrings("%AppData%") &  "\Microsoft\Signatures\"
'====================/EDIT-HERE===============================================

' Look for the signatures folder and delete all files in it. Or, if it doesnt exist then create it:
If objFSO.FolderExists(strLocalSignatureFolder) Then
	objFSO.DeleteFile (strLocalSignatureFolder & "\*.*")
Else
	objFSO.CreateFolder(strLocalSignatureFolder)
End If        

'Copy the appropriate images (based on Department) to the local signature folder (e.g companyname-finance-logo.png):
Dim LogoImage, FooterImage, ExtraImage
LogoImage =  Replace(strCompanyName & "-" & objUser.Department & "-logo.jpg", " ", "_")
If objFSO.FileExists(strServerSignatureFolder & LogoImage) Then 
	objFSO.CopyFile strServerSignatureFolder & LogoImage, strLocalSignatureFolder.Path
End If
FooterImage = Replace(strCompanyName & "-" & objUser.Department & "-footer.jpg", " ", "_")
If objFSO.FileExists(strServerSignatureFolder & FooterImage) Then 
	objFSO.CopyFile strServerSignatureFolder & FooterImage, strLocalSignatureFolder.Path
End If
ExtraImage = Replace(strCompanyName & "-" & objUser.Department & "-extra.jpg", " ", "_")
If objFSO.FileExists(strServerSignatureFolder & ExtraImage) Then 
	objFSO.CopyFile strServerSignatureFolder & ExtraImage, strLocalSignatureFolder.Path
End If

'Build the HTML File:
Dim objFile
Set objFile = objFSO.CreateTextFile(strLocalSignatureFolder & "\Default.html")

Dim chrQuote
chrQuote = chr(34)

objFile.Write "<!DOCTYPE HTML PUBLIC " & chrQuote & "-//W3C//DTD HTML 4.0 Transitional//EN" & chrQuote & ">" & vbCrLf
objFile.Write "<html><head>" & vbCrLf
objFile.Write "<meta http-equiv=Content-Type content=" & chrQuote & "text/html; charset=windows-1252" & chrQuote & ">" & vbCrLf
objFile.Write "<meta content=" & chrQuote & "MSHTML 6.00.3790.186" & chrQuote & " name=GENERATOR><style>" & vbCrLf
objFile.Write "    body {font-family: Arial;font-size: 13px;color:#000000;margin-left:0px;margin-top: 0px;padding:0px;}" & vbCrLf
objFile.Write "    .style1 {font-size: 16px; font-weight: bold; color: #002654;}" & vbCrLf
objFile.Write "    .style2 {font-family:Arial;font-size: 12px; color: #002654;}" & vbCrLf
objFile.Write "</style></head>" & vbCrLf
objFile.Write "<body bgcolor=""ffffff"" style=""PADDING: 0px; MARGIN: 0px"">" & vbCrLf
objFile.Write "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""665px"">" & vbCrLf
objFile.Write " <tr><td valign=""bottom""style=""font-family: Arial;font-size:16px;color:#002654;line-height:16px;padding-left:6px;width=""285px""><p><span class=""style1""><b>" & objUser.DisplayName & "</b></span><br></td></tr>" & vbCrLf
objFile.Write " <tr><td valign=""bottom""style=""font-family: Arial;font-size:12px;color:#002654;line-height:16px;padding-left:6px;width=""285px""><p><span class=""style2"">" & objUser.Title & "</span><br></td></tr><br>" & vbCrLf
objFile.Write " <tr><td valign=""bottom""style=""font-family:Arial;font-size:12;color:#002654;line-height:16px;padding-left:6px;width=""285px""><p>"
If objLocation.TelephoneNumber <> "" Then 			objFile.Write "P. "& objLocation.TelephoneNumber & " "
If objLocation.FacsimileTelephoneNumber <> "" Then 	objFile.Write "F.  " & objLocation.FacsimileTelephoneNumber & "<br>" & vbCrLf
If objUser.Mobile <> "" Then 						objFile.Write "M.  " & objUser.Mobile & "<br>" & vbCrLf
If objUser.Mail <> "" Then 							objFile.Write "E. " & "<a href=""mailto:"& LCase(objUser.Mail) &""" style=""text-decoration:none;color:#002654;"" target=""_blank"">"& objUser.Mail &"</a><br>" & vbCrLf
objFile.Write objLocation.Street & " " & objLocation.l & " " & objLocation.st & " " & objLocation.PostalCode & "<br>" & vbCrLf
objFile.Write "<a href=""http://"& strWebAddress &""" target=""_blank"" style=""color:#002654;text-decoration:none;""><b>"& strWebAddress &"</b></a></td>" & vbCrLf
objFile.Write "</td></tr><td valign=""bottom"" align=""right"" width=""315px""><br></td></tr>" & vbCrLf
If objFSO.FileExists(strLocalSignatureFolder & LogoImage) Then
	objFile.Write "<br><tr><td colspan=""4""><img src="& LogoImage &" width=""136"" height=""40""><br><br></td></tr>" & vbCrLf
End If
If objFSO.FileExists(strLocalSignatureFolder & FooterImage) Then
	objFile.Write "<tr><td colspan=""4""><img src="& FooterImage &" width=""273"" height=""23""><br><br></td></tr>" & vbCrLf
End If
If objFSO.FileExists(strLocalSignatureFolder & ExtraImage) Then
	objFile.Write "<tr><td colspan=""4""><img src="& ExtraImage &" width=""74"" height=""45""><br><br></td></tr>" & vbCrLf
End If
objFile.Write "<tr><td valign=""bottom""style=""font-family:Arial;font-size:10;color:#c0c0c0;line-height:1.1em;text-align:justify;padding-left:6px;width=""200px""><p>"& strDisclaimer & "</tr>" & vbCrLf
objFile.Write "</table>" & vbCrLf
objFile.Write "</body></html>"
objFile.Close

'Convert the HTML Version to RTF and Plain Text:
'WordFileConversion(strLocalSignatureFolder & "\Default.html", strLocalSignatureFolder & "\Default.rtf", wdFormatRTF)
'WordFileConversion(strLocalSignatureFolder & "\Default.html", strLocalSignatureFolder & "\Default.txt", wdFormatTextLineBreaks)

Sub WordFileConversion (inputFile, outputFile, wordFormat)
    Dim objDoc, objFile, objWord
    Const wdFormatDocument                    =  0
	Const wdFormatText                        =  2
	Const wdFormatTextLineBreaks              =  3
	Const wdFormatRTF                         =  6
	Const wdFormatHTML                        =  8
	Const wdFormatPDF                         = 17
    Set objWord = CreateObject("Word.Application")
	objWord.Visible = False
	If objFSO.FileExists(inputFile) Then
		Set objFile = objFSO.GetFile(inputFile)
	Else
		objWord.Quit
		Exit Sub
	End If
	objWord.Documents.Open objFile.Path
	Set objDoc = .ActiveDocument
	objDoc.SaveAs outputFile, wordFormat
	objDoc.Close
	objWord.Quit
End Sub
