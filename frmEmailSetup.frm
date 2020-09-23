VERSION 5.00
Begin VB.Form frmEmailSetup 
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
   Picture         =   "frmEmailSetup.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDomain 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3240
      TabIndex        =   17
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CheckBox chkSavePAssword 
      BackColor       =   &H00800000&
      Caption         =   "Check1"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   1560
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdAddPOP 
      BackColor       =   &H00800000&
      Caption         =   "+"
      Height          =   255
      Left            =   5400
      MaskColor       =   &H00C00000&
      TabIndex        =   13
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdAddSMTP 
      BackColor       =   &H00800000&
      Caption         =   "+"
      Height          =   255
      Left            =   5400
      MaskColor       =   &H00C00000&
      TabIndex        =   12
      Top             =   840
      Width           =   255
   End
   Begin VB.ComboBox cmbSMTP 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Text            =   "smtp.mail.yahoo.com"
      Top             =   840
      Width           =   4095
   End
   Begin VB.ComboBox cmbPOP 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Text            =   "pop.mail.yahoo.com"
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Save Password ?"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Server"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "POP Server"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label cmdReset 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reset"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label cmdCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.Image imgClose 
      Height          =   210
      Left            =   5520
      Picture         =   "frmEmailSetup.frx":02E2
      Top             =   0
      Width           =   225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Email CLient Setup"
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
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label cmdOK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OK"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   4920
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   -1200
      Picture         =   "frmEmailSetup.frx":05C4
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   6975
   End
End
Attribute VB_Name = "frmEmailSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AppPath As String





Private Sub cmdAddPOP_Click()
Dim I As Integer

For I = 0 To cmbPOP.ListCount - 1
 If cmbPOP.List(I) = cmbPOP.Text Then Exit Sub
Next I

cmbPOP.AddItem cmbPOP.Text

'//// Put the code to save the server list here
UpdateList AppPath & "POPserver.txt", cmbPOP.Text





End Sub

Private Sub cmdAddSMTP_Click()

Dim I As Integer

For I = 0 To cmbSMTP.ListCount - 1
 If cmbSMTP.List(I) = cmbSMTP.Text Then Exit Sub
Next I

cmbSMTP.AddItem cmbSMTP.Text

'//// Put the code to save the server list here

UpdateList AppPath & "SMTPserver.txt", cmbSMTP.Text


End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()

POPAddress = Trim$(cmbPOP.Text)
SMTPAddress = Trim$(cmbSMTP.Text)
SMTPDomain = Trim$(txtDomain.Text)
EmailUsername = Trim$(txtUsername.Text)
EmailPass = Trim$(txtPass.Text)

'UPdate the lists

cmdAddPOP_Click
cmdAddSMTP_Click



'Clear the password field
If chkSavePAssword.Value = 0 Then
 txtPass.Text = ""
End If

Unload Me

End Sub

Private Function RetrieveList(strFilename As String, intRecordNumber As Integer) As String

Dim Filenum As Integer
Dim LineInFile As Integer
Dim strBuffer As String

LineInFile = 0

'Check the availability of file
 If Dir(strFilename) = "" Then
  RetrieveList = "Err"
  Exit Function
 End If

Filenum = FreeFile 'Assign the file handle

Open strFilename For Input As Filenum

Do
 Input #Filenum, strBuffer
 LineInFile = LineInFile + 1
Loop Until LineInFile = intRecordNumber



'Close the file
Close Filenum

RetrieveList = strBuffer


End Function

Private Function UpdateList(strFilename As String, strData As String) As Boolean

Dim Filenum As Integer
Dim FileOpened As Boolean


 
 
On Error GoTo Erh:

Filenum = FreeFile

Open strFilename For Append As #Filenum

FileOpened = True

Write #Filenum, strData

Close Filenum
FileOpened = False


On Error GoTo 0


Erh:

If FileOpened Then Close Filenum

End Function

Private Function CountList(strFilename As String) As Integer

Dim I As Integer
Dim Filenum As Integer
Dim FileOpened As Boolean
Dim strDummy As String

'Check the availability of file
 If Dir(strFilename) = "" Then
  CountList = 0
  Exit Function
 End If

Filenum = FreeFile

Open strFilename For Input As Filenum
FileOpened = True

Do
 Input #Filenum, strDummy
 I = I + 1
Loop Until EOF(Filenum) = True

Close Filenum
FileOpened = False

CountList = I

Erh:

If FileOpened Then Close Filenum

End Function

Private Sub cmdReset_Click()
 Form_Load
End Sub

Private Sub Form_Load()

Dim I As Integer
Dim I2 As Integer
Dim strPOP As String
Dim strSMTP As String

'Set the application path
AppPath = App.Path

If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"

If Dir(AppPath & "POPserver.txt") <> "" Then
 'Update pop servers list
 I = CountList(AppPath & "POPserver.txt")
 
 For I2 = 1 To I
  strPOP = RetrieveList(AppPath & "POPserver.txt", I2)
  cmbPOP.AddItem strPOP
 Next I2

End If

If Dir(AppPath & "SMTPserver.txt") <> "" Then
 I = CountList(AppPath & "SMTPserver.txt")
 
 For I2 = 1 To I
  strSMTP = RetrieveList(AppPath & "SMTPserver.txt", I2)
  cmbSMTP.AddItem strSMTP
 Next I2
 
End If

End Sub

Private Sub Label1_Click()

End Sub
