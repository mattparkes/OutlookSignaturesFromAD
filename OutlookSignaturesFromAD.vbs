'***************************************************************************************************
'***************************************************************************************************
'*                       OutlookSigGenerator.vbs                                                   *
'*                       Written By: Jaeden Cook, David Phelan & Matt Parkes                       *
'*                                                                                                 *
'*                       Version: 3.0.1		Date: 15/05/2013                                       *
'*                                                                                                 *
'*  This script gets user info from Active Directory and uses that information to create Outlook   *
'*  Signature files in Text, HTML, and RTF formats. 											   *
'*  Office Address, Phone, Fax, etc is retrieved from the User's Parent OU (set using ADSI Edit).  *
'***************************************************************************************************
Option Explicit
Const wdFormatHTML                        =  8
Const wdFormatRTF                         =  6
Const wdFormatTextLineBreaks              =  3

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
Dim strSignatureName, strCompanyName, strDisclaimer, strWebAddress
strSignatureName = "Default"
strCompanyName = "Contoso"
strWebAddress = "www.contoso.com" 
strDisclaimer = "Disclaimer: This communication including any attachments may contain information that is either confidential or otherwise protected from disclosure and is intended"&_
      "solely for the use of the intended recipient. If you are not the intended recipient please immediately notify the sender by e-mail and delete the original transmission and"&_
      "its contents. Any unauthorised use, dissemination, forwarding, printing, or copying of this communication including any file attachments is prohibited. The recipient"&_
      "should check this email and any attachments for viruses and other defects. The Company disclaims any liability for loss or damage arising in any way from this"&_
      "communication including any file attachments."
Dim strServerSignatureFolder, strLocalSignatureFolder
strServerSignatureFolder = "\\contoso-fs-01.contoso.local\signatures$\"
strLocalSignatureFolder = objShell.ExpandEnvironmentStrings("%AppData%") &  "\Microsoft\Signatures\"
'====================/EDIT-HERE===============================================

' Look for the signatures folder and delete any existing copy of our signature. Or, if it doesnt exist then create it:
If objFSO.FolderExists(strLocalSignatureFolder) Then
	objFSO.DeleteFile (strLocalSignatureFolder & "\" & strSignatureName & ".*")
Else
	objFSO.CreateFolder(strLocalSignatureFolder)
End If        

'Copy the appropriate images (based on Department) to the local signature folder (spaces changed to underscores). e.g 'companyname-human_resources-logo.png'
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
Set objFile = objFSO.CreateTextFile(strLocalSignatureFolder & "\" & strSignatureName & ".html")

objFile.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCrLf
objFile.Write "<html><head>" & vbCrLf
objFile.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=us-ascii"">"
objFile.Write "<style>" & vbCrLf
objFile.Write "	body {font-family: Arial;font-size: 13px;color:#000000;margin-left:0px;margin-top: 0px;padding:0px;}" & vbCrLf
objFile.Write "	.style1 {font-size: 16px; font-weight: bold; color: #002654;}" & vbCrLf
objFile.Write "	.style2 {font-family:Arial;font-size: 12px; color: #002654;}" & vbCrLf
objFile.Write "</style><title></title></head>" & vbCrLf
objFile.Write "<body bgcolor=""ffffff"" style=""padding: 0px; margin: 0px"">" & vbCrLf
objFile.Write "<table cellpadding=""0"" cellspacing=""0"" border=""0"" width=""665px"">" & vbCrLf
objFile.Write " <tr><td valign=""bottom""style=""font-family: Arial;font-size:16px;color:#002654;line-height:16px;padding-left:6px;width=""285px""><span class=""style1""><b>" & objUser.DisplayName & "</b></span><br></td></tr>" & vbCrLf
objFile.Write " <tr><td valign=""bottom""style=""font-family: Arial;font-size:12px;color:#002654;line-height:16px;padding-left:6px;width=""285px""><span class=""style2"">" & objUser.Title & "</span><br></td></tr><br>" & vbCrLf
objFile.Write " <tr><td valign=""bottom""style=""font-family:Arial;font-size:12;color:#002654;line-height:16px;padding-left:6px;width=""285px"">"
If objLocation.TelephoneNumber <> "" Then 			objFile.Write "P. "& objLocation.TelephoneNumber & " "
If objLocation.FacsimileTelephoneNumber <> "" Then 	objFile.Write "F.  " & objLocation.FacsimileTelephoneNumber & "<br>" & vbCrLf
If objUser.Mobile <> "" Then 						objFile.Write "M.  " & objUser.Mobile & "<br>" & vbCrLf
If objUser.Mail <> "" Then 							objFile.Write "E. " & "<a href=""mailto:"& LCase(objUser.Mail) &""" style=""text-decoration:none;color:#002654;"" target=""_blank"">"& objUser.Mail &"</a><br>" & vbCrLf
objFile.Write objLocation.Street & " " & objLocation.l & " " & objLocation.st & " " & objLocation.PostalCode & "<br>" & vbCrLf
objFile.Write "<a href=""http://"& strWebAddress &""" target=""_blank"" style=""color:#002654;text-decoration:none;""><b>"& strWebAddress &"</b></a></td></tr>" & vbCrLf
If objFSO.FileExists(strLocalSignatureFolder & LogoImage) Then
	objFile.Write "<br><tr><td><img src="& LogoImage &" width=""136"" height=""40""><br><br></td></tr>" & vbCrLf
End If
If objFSO.FileExists(strLocalSignatureFolder & FooterImage) Then
	objFile.Write "<tr><td><img src="& FooterImage &" width=""273"" height=""23""><br><br></td></tr>" & vbCrLf
End If
If objFSO.FileExists(strLocalSignatureFolder & ExtraImage) Then
	objFile.Write "<tr><td><img src="& ExtraImage &" width=""74"" height=""45""><br><br></td></tr>" & vbCrLf
End If
objFile.Write "<tr><td valign=""bottom""style=""padding-top:1em;font-family:Arial;font-size:10;color:#c0c0c0;line-height:1.1em;text-align:justify;padding-left:6px;width=""200px"">" & strDisclaimer & "</td></tr>" & vbCrLf
objFile.Write "</table>" & vbCrLf
objFile.Write "</body></html>"
objFile.Close

'Convert the HTML Version to RTF and Plain Text:
Dim objDoc, objWord
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
If objFSO.FileExists(strLocalSignatureFolder & "\" & strSignatureName & ".html") Then
	objWord.Documents.Open strLocalSignatureFolder & "\" & strSignatureName & ".html"
	Set objDoc = objWord.ActiveDocument
	objDoc.SaveAs strLocalSignatureFolder & "\" & strSignatureName & ".rtf", wdFormatRTF
	objDoc.SaveAs strLocalSignatureFolder & "\" & strSignatureName & ".txt", wdFormatTextLineBreaks
	objDoc.Close
	objWord.Quit
End If
