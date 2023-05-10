VERSION 1.0 CLASS

BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute Items.VB_VarHelpID = -1
Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
	Dim olApp As Outlook.Application
	Dim objNS As Outlook.NameSpace
	Set olApp = Outlook.Application
	Set objNS = olApp.GetNamespace("MAPI")
	Set Items = objNS.GetDefaultFolder(olFolderInbox).Folders("DMARC").Items
End Sub

Private Sub Items_ItemAdd(ByVal item As Object)
	On Error GoTo ErrorHandler
	Dim Msg As Outlook.MailItem
	If TypeName(item) = "MailItem" Then
		Set Msg = item
		If InStr(1, Msg.To, "dmarc", 1) Then
			If Msg.Attachments.Count > 0 Then
				For Each attch In Msg.Attachments
					attch.SaveAsFile "C:\Temp\DMARC\" & attch.FileName
					Call Shell("pwsh.exe -ExecutionPolicy ByPass -File ""C:\Users\" & Environ("username") & "\Documents\Code\DMARCAudit\ScanDmarc.ps1"" -F " & attch.FileName, 0)
					Next
			End If
		End If
	End If

	ProgramExit:
		Exit Sub
	
	ErrorHandler:
		MsgBox Err.Number & " - " & Err.Description
		Resume ProgramExit
End Sub
