Attribute VB_Name = "modGUI"
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
 Private Const conHwndTopmost = -1
 Private Const conSwpNoActivate = &H10
 Private Const conSwpShowWindow = &H40
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
  ByVal dwflags As Long, ByVal dwExtraInfo As Long)
    
Public Sub DragForm(CurrentForm As Form)
    ReleaseCapture
    Call SendMessage(CurrentForm.hWnd, &HA1, 2, 0&)
End Sub



Public Function TakeSnapshot(Optional ByVal FileName As String, Optional win_hwnd As Long) As Boolean


Dim lString As String
Dim ImagePath As String

On Error GoTo Trap
'Check if the File Exist
ImagePath = App.Path
If Right$(ImagePath, 1) <> "\" Then ImagePath = ImagePath & "\"
ImagePath = ImagePath + FileName + ".bmp"

    If Dir(ImagePath) <> "" Then Exit Function

    'To get the Entire Screen
    Call keybd_event(vbKeySnapshot, 1, 0, 0)

      'To get the Active Window
    'Call keybd_event(vbKeySnapshot, 0, 0, 0)
 
    SavePicture Clipboard.GetData(vbCFBitmap), ImagePath
    

TakeSnapshot = True
Exit Function

Trap:
'Error handling
MsgBox "Error #: " & Err.Number & ", " & Err.Description

End Function

