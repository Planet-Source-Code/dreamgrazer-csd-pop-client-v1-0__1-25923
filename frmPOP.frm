VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmPOP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPOP.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckPOP 
      Left            =   5040
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   110
   End
   Begin VB.ListBox lstHeaders 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      ItemData        =   "frmPOP.frx":02E2
      Left            =   120
      List            =   "frmPOP.frx":02E4
      TabIndex        =   4
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label cmdSetup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label cmdRefresh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label cmdDelete 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.Image imgClose 
      Height          =   210
      Left            =   5520
      Picture         =   "frmPOP.frx":02E6
      Top             =   0
      Width           =   225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "POP Email Reader"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label cmdConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   4680
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   -1200
      Picture         =   "frmPOP.frx":05C8
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   6975
   End
End
Attribute VB_Name = "frmPOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevCommand As String
Dim IsPOPConnected As Boolean
Dim MessageBody As String
Dim FromDate As String
Dim TOPReceived As Boolean
Dim TOPError As Boolean
Dim IsListing As Boolean


Private Sub cmdConnect_Click()

If cmdConnect.Caption = "Disconnect" Then
 'Disconnect from POP
 If IsPOPConnected Then
  SendPOP "QUIT"
 End If
 
 'Close the socket..of course
 sckPOP.Close
  'Inform user
 lblStatus.Caption = "Disconnected"
 
 cmdConnect.Caption = "Connect"

 Exit Sub
End If



If EmailUsername = "" Or EmailPass = "" Or POPAddress = "" Then
 csMsgbox "Information not complete.Please fill all fields.", "User Information not complete", "CSOKONLY"
 frmEmailSetup.Show 0
 Exit Sub
End If

'REname button
cmdConnect.Caption = "Disconnect"

'Connect to POP server
sckPOP.Connect POPAddress, 110



End Sub

Private Sub cmdDelete_Click()

Dim csFlags As String

'Delete the selected header

If IsConfirmDelete Then
 csFlags = csMsgbox("Are you sure to delete this message ?", "Delete Confirmation", "CSYESCANCEL")
 'If user canceled then exit la sub nih
 If csFlags = "CSCANCEL" Then Exit Sub
 
End If

SendPOP "DELE " & lstHeaders.ListIndex + 1


End Sub

Private Sub cmdRefresh_Click()
 'UPdate the headers list
 SendPOP "LIST"
 'Getting the list
 lblStatus.Caption = "Requesting for listing of messages"
End Sub

Private Sub cmdSetup_Click()
 frmEmailSetup.Show 0
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
 DragForm Me
End If
End Sub

Private Sub imgClose_Click()
Unload frmPOPMsg

Unload Me
 
End Sub


Private Sub lstHeaders_DblClick()

'Retrieve the clicked message
lblStatus.Caption = "Downloading message"

SendPOP "RETR " & lstHeaders.ListIndex + 1

End Sub

Private Sub sckPOP_Close()

'Closing call.Make sure properly closed
If sckPOP.State <> sckClosed Then sckPOP.Close


End Sub

Private Sub sckPOP_Connect()
'Connected to the server.
lblStatus.Caption = "Connected to server"
PrevCommand = "CONNECT"

End Sub

Private Sub sckPOP_ConnectionRequest(ByVal requestID As Long)
 'Why a connection request ??
 
End Sub

Private Sub sckPOP_DataArrival(ByVal bytesTotal As Long)

Dim datarec As String
Dim I As Integer
Dim strTemp As String

'Get the data
sckPOP.GetData datarec, vbString

Select Case Left$(datarec, 3)

Case "+OK" 'OK..no error

 Select Case PrevCommand
  
  Case "CONNECT"
  
   'Just connected
   IsPOPConnected = True
   'Login to server.Send username
   SendPOP "USER " & EmailUsername
   lblStatus.Caption = "Checking Username"
   
  Case "USER"
   lblStatus.Caption = "Checking password"
   SendPOP "PASS " & EmailPass

  Case "PASS"
   'User login successfully
   lblStatus.Caption = "Login successful"
   'Get the list
   
   SendPOP "LIST"
   lblStatus.Caption = "Retrieving list"
   
   
   
  Case "LIST"
   'Receiving the list of available email in current folder.
   UpdateList datarec
   
   
  Case "TOP "
   'Receiving header of the mail
   TOPReceived = True
   
   lstHeaders.AddItem FindSubject(datarec)
   
   
   
   
   
  Case "RETR"
  
  
  
   'Receiving the fullmessage
   '//// Header parser unfinished
   'Show Full Message Window
   frmPOPMsg.Show 0
   'Process message
   ProcessMessage datarec
   'Fill in subject
   frmPOPMsg.lblSubject.Caption = FromSubject
   frmPOPMsg.lblFrom.Caption = FromAddress
   'Fill in the text
   frmPOPMsg.txtMsg.Text = MessageBody
   
  Case "DELE"
   'Message marked as deleted
   lblStatus.Caption = "Message marked deleted"
   'Remove the deleted item
   lstHeaders.RemoveItem (lstHeaders.ListIndex)
   
   
  Case "RSET"
   'Undo all delete marks
   lblStatus.Caption = "Messages delete aborted"
   'UPdate the headers list
   lstHeaders.Clear
   SendPOP "LIST"
   
   
  Case "QUIT"
  
  'User Logout
  lblStatus.Caption = "Logout Successful"
  'Close socket
  sckPOP.Close
  'Rename button
  cmdConnect.Caption = "Connect"

  End Select
 
 Case "-ER" 'Error
  
  Select Case PrevCommand
  
  Case "USER"
  
  'Inform user
  lblStatus.Caption = "Username Invalid.Try Again"
  'Close connection
  SendPOP "QUIT"
  'CLose socket
  sckPOP.Close
  'Reenable button
  cmdConnect.Caption = "Connect"
  
  Case "PASS"
  
  'Inform user
  lblStatus.Caption = "Password Invalis.Try Again"
  'Close connection
  SendPOP "QUIT"
  'CLose socket
  sckPOP.Close
  'Reenable button
  cmdConnect.Caption = "Connect"
  
  Case "LIST"
  
  Case "RETR"
  
    
  lblStatus.Caption = "Failed to retrieve message.Try relogin"
  
  Case "TOP "
  
   TOPError = True
   
   
   
  
  Case "DELE"
  
  lblStatus.Caption = "Unable to delete message.Try again"
  
  Case "RSET"
  
  lblStatus.Caption = "Mailbox Reset Failed.Try Again"
 End Select
End Select


End Sub

Private Sub sckPOP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'An error occured, inform user.
csMsgbox "An error occured: [" & Str(Number) & "] Description: " & Description, "Error", "CSOKONLY"

lblStatus.Caption = "Error.Disconnected from server"

sckPOP.Close

cmdConnect.Caption = "Connect"


End Sub

Private Sub SendPOP(strData As String)

PrevCommand = Left$(strData, 4)

sckPOP.SendData strData & vbCrLf

End Sub

Private Sub ProcessMessage(strData As String)

Dim nPos As Integer

nPos = InStr(1, strData, "Date")

FromDate = Mid$(strData, nPos + 6, 15)
FromDate = szTrimCRLF(FromDate)

nPos = InStr(1, strData, "From")

FromAddress = Mid$(strData, nPos + 6, InStr(nPos + 6, strData, vbCrLf) - (nPos + 6))
FromAddress = szTrimCRLF(FromAddress)

nPos = InStr(1, strData, "Subject")

FromSubject = Mid$(strData, nPos + 8, InStr(nPos + 8, strData, vbCrLf) - (nPos + 8))

FromSubject = szTrimCRLF(FromSubject)

MessageBody = Mid$(strData, InStr(1, strData, vbCrLf & vbCrLf) + 2)


End Sub

Private Function FindSubject(strData As String) As String

Dim nPos As Integer
Dim nPos2 As Integer
Dim strBuffer As String



nPos = InStr(1, strData, "Subject") + 8
nPos2 = InStr(nPos, strData, vbCrLf)

If nPos2 > nPos Then
 strBuffer = Mid$(strData, nPos, nPos2 - nPos)
Else
 strBuffer = Mid$(strData, nPos)
End If


FindSubject = szTrimCRLF(strBuffer)

End Function

Private Sub UpdateList(strData As String)
'First remove the '+OK' mesage

Dim strBuffer As String
Dim nPos As Integer
Dim NumOfLines As Integer

Dim I As Integer

lblStatus.Caption = "Counting the number of messages"




strBuffer = Mid$(strData, 5, InStr(1, strData, "m") - 6)

NumOfLines = Val(strBuffer)


lblStatus.Caption = "Listing the meesages headers"



'Get the headers of every message



For I = 1 To NumOfLines

 SendPOP "TOP " & Str(I) & " 0"
 
 'Wait for Header to be received
 Do
  DoEvents
 Loop Until TOPReceived = True Or TOPErrror = True
 
 'Reset the flags
 If TOPReceived = True Then TOPReceived = False
 
 If TOPError = True Then
  lblStatus.Caption = "Error in one of header"
  TOPError = False
 End If
 
Next I

lblStatus.Caption = "List finished"


End Sub
