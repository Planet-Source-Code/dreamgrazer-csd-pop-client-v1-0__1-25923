Attribute VB_Name = "modMain"
'====================================================================
'       Main Module for csD Terminal
'
'====================================================================


Private Type UserAuth 'UDT for storing account data
 Username As String * 10
 Password As String * 12
 Level As Integer
End Type


Public Type cliDat 'UDT for storing client data
 Name As String * 100
 Username As String * 10
 AuthLevel As Integer
' 1000 - Server Admin - Access to all part
' 500 - Super Operator - Access to advanced feature
' 0 - Normal User - Access only to chatting feature
 Password As String * 12
 IP As String * 15
 IsAuth As Boolean
End Type

Global cliData() As cliDat 'Data of client
Global Userdata As UserAuth 'UDT for account info

Global NumOfClients As Integer
Global svrIp As String
Global svrName As String
Global IsServer As Boolean 'The machine is currently a server or not
Global IsConnected As Boolean 'Connected or NOT !
Global IsLogIn As Boolean 'Fill in and press login in login form
Global csMRet As String





Public Sub SendData(sData As String)
    On Error GoTo ErrH

    Dim TimeOut As Long
    
    frmTerminal.sckConnect(1).SendData sData
    
    Do Until (frmTerminal.sckConnect(1).State = 0) Or (TimeOut < 10000)
        DoEvents
        TimeOut = TimeOut + 1
        If TimeOut > 10000 Then Exit Do
    Loop
    
ErrH:
    Exit Sub
End Sub

Public Sub svrSendData(sData As String, ctrlIndex As Integer)
    On Error GoTo ErrH

    Dim TimeOut As Long
    
    frmTerminal.sckConnect(ctrlIndex).SendData sData
    
    Do Until (frmTerminal.sckConnect(ctrlIndex).State = 0) Or (TimeOut < 10000)
        DoEvents
        TimeOut = TimeOut + 1
        If TimeOut > 10000 Then Exit Do
    Loop
    
ErrH:
    Exit Sub
End Sub

Public Sub SetStatus(strStatus As String) 'Update the panel in terminal
 
 frmTerminal.txtPanel.Text = strStatus
 
End Sub


Public Function CheckPass(strUsername As String, strPass As String) As Boolean

Dim Filepath As String
Dim NumOfRecords As Integer
Dim I As Integer
Dim IsFileOpened As Boolean

On Error GoTo Erh:

Filepath = App.Path

If Right$(Filepath, 1) <> "\" Then Filepath = Filepath & "\"
Filepath = Filepath & "Pass.pwd"

NumOfRecords = FileLen(Filepath) / Len(Userdata)
Open Filepath For Random As #1 Len = Len(Userdata)

IsFileOpened = True



For I = 1 To NumOfRecords
 Get #1, I, Userdata
 If Userdata.Username = strUsername Then
  If Userdata.Password = strPass Then
   CheckPass = True
  End If
 End If
 
Next I

Close #1

IsFileOpened = False

CheckPass = False


Erh:

If IsFileOpened Then Close #1

CheckPass = False

End Function

Public Function AddAccount(strUsername As String, strPass As String, intLevel As Integer) As Boolean
 
Dim Filepath As String
Dim NumOfRecords As Integer

Dim IsFileOpened As Boolean

On Error GoTo Erh:

Userdata.Level = intLevel
Userdata.Password = strPass
Userdata.Username = strUsername


Filepath = App.Path

If Right$(Filepath, 1) <> "\" Then Filepath = Filepath & "\"
Filepath = Filepath & "Pass.pwd"

NumOfRecords = FileLen(Filepath) / Len(Userdata)
Open Filepath For Random As #1 Len = Len(Userdata)

IsFileOpened = True

Put #1, NumOfRecords + 1, Userdat

Close #1
IsFileOpened = False

AddAccount = True

Erh:

 If IsFileOpened Then Close #1
 
 AddAccount = False

End Function

Public Function RemoveAccount(strUsername As String) As Boolean

Dim Filepath As String
Dim NumOfRecords As Integer
Dim IsFileOpened As Boolean
Dim TempPath As String
Dim NumOfRecordsTemp As Integer

On Error GoTo Erh:


NumOfRecordsTemp = 1

Filepath = App.Path
TempPath = App.Path

If Right$(Filepath, 1) <> "\" Then Filepath = Filepath & "\"
Filepath = Filepath & "Pass.pwd"

If Right$(TempPath, 1) <> "\" Then TempPath = TempPath & "\"
TempPath = TempPath & "Temp.pwd"

NumOfRecords = FileLen(Filepath) / Len(Userdata)
Open Filepath For Random As #1 Len = Len(Userdata)
Open TempPath For Random As #2 Len = Len(Userdata)
IsFileOpened = True

For I = 1 To NumOfRecords
 Get #1, NumOfRecords, Userdata
 If Userdata.Username <> strUsername Then
  Put #2, NumOfRecordsTemp, Userdata
  NumOfRecordsTemp = NumOfRecordsTemp + 1
 End If
Next I

Close #1
Close #2

IsFileOpened = False

Kill Filepath

'Put the procedures for changing the name of file here

RemoveAccount = True

Erh:

If IsFileOpened Then
 Close #1
 Close #2
End If

RemoveAccount = False

End Function

Public Function ModifyAccount(strUsername As String, strPass As String, Optional intLevel As Integer, Optional strNewPass As String) As Boolean

Dim Filepath As String
Dim NumOfRecords As Integer
Dim IsFileOpened As Boolean

On Error GoTo Erh:


Filepath = App.Path

If Right$(Filepath, 1) <> "\" Then Filepath = Filepath & "\"
Filepath = Filepath & "Pass.pwd"

NumOfRecords = FileLen(Filepath) / Len(Userdata)
Open Filepath For Random As #1 Len = Len(Userdata)

IsFileOpened = True

For I = 1 To NumOfRecords
Get #1, NumOfRecords, Userdata
If Userdata.Username = srUsername Then
 If Userdata.Password = strPassword Then
  If strNewPass <> "" Then
   Userdata.Password = strNewPass
  End If
  If Str(intLevel) <> "" Then
   Userdata.Level = intLevel
  End If
  Put #1, NumOfRecords, Userdata
  End If
End If
Next I

Close #1
IsFileOpened = False

ModifyAccount = True

Erh:

If IsFileOpened Then Close #1
ModifyAccount = False

End Function

Public Function GetLevel(strUsername As String) As Integer
Dim Filepath As String
Dim NumOfRecords As Integer

Dim IsFileOpened As Boolean

On Error GoTo Erh:

Userdata.Level = intLevel
Userdata.Password = strPass
Userdata.Username = strUsername


Filepath = App.Path

If Right$(Filepath, 1) <> "\" Then Filepath = Filepath & "\"
Filepath = Filepath & "Pass.pwd"

NumOfRecords = FileLen(Filepath) / Len(Userdata)
Open Filepath For Random As #1 Len = Len(Userdata)

IsFileOpened = True

For I = 1 To NumOfRecords
 Get #1, NumOfRecords, Userdata
 If strUsername = strUsername Then
  GetLevel = Userdata.Level
 End If
Next I

Close #1
IsFileOpened = False

MsgBox "User does not exist in the database"

Erh:

If IsFileOpened Then Close #1

GetLevel = 0

End Function
' A subroutine for a custom message box
Public Function csMsgbox(strMessage As String, strTopic As String, Optional typMsgbox As String) As String
 
 frmMsgbox.lblMsg.Caption = strMessage
 
 Load frmMsgbox
 
 If strTopic <> "" Then
  frmMsgbox.lblTopic.Caption = strTopic
 End If
 
 If UCase$(typMsgbox) = "CSOKONLY" Then
  frmMsgbox.cmdOK.Visible = True
 ElseIf UCase$(typMsgbox) = "CSYESNO" Then
  frmMsgbox.cmdYes.Visible = True
  frmMsgbox.cmdNo.Visible = True
 ElseIf UCase$(typMsgbox) = "CSYESNOCANCEL" Then
  frmMsgbox.cmdYes.Visible = True
  frmMsgbox.cmdNo.Visible = True
  frmMsgbox.cmdCancel.Visible = True
 ElseIf UCase$(typMsgbox) = "CSOKCANCEL" Then
  frmMsgbox.cmdYes.Visible = True
  frmMsgbox.cmdCancel.Visible = True
 ElseIf UCase$(typMsgbox) = "CSYESCANCEL" Then
  frmMsgbox.cmdYes.Visible = True
  frmMsgbox.cmdCancel.Visible = True
 
 Else
  frmMsgbox.cmdOK.Visible = True
 End If
 
 frmMsgbox.Show 0
 
 csMsgbox = csMRet
 
End Function

Public Sub ResetData(ctrlIndex As Integer)

cliData(ctrlIndex).AuthLevel = 0
cliData(ctrlIndex).IP = ""
cliData(ctrlIndex).IsAuth = False
cliData(ctrlIndex).Name = ""
cliData(ctrlIndex).Password = ""
cliData(ctrlIndex).Username = ""

End Sub


