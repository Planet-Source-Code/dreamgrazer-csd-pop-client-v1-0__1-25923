Attribute VB_Name = "modEmail"
Global SMTPAddress As String
Global POPAddress As String
Global EmailUsername As String
Global EmailPass As String
Global SMTPDomain As String
Global IsConfirmDelete As String
'Use in POP Message handling
Global FromSubject As String
Global FromAddress As String
Global FromDate As String
'Use in Email Attachments
Global FileAttached(10) As String 'Set 10 = Maximum
Global AttachedFiles As Boolean





Function szTrimCRLF(szString As String) As String

 ' Taken from VB6 Internet Programming by Carl Franklin
 
    Dim lStr As Integer
    
    lStr = Len(szString)
    
    If lStr Then
        If Right$(szString, 2) = vbCrLf Then
            szTrimCRLF = Left$(szString, lStr - 2)
        Else
            Select Case Right$(szString, 1)
                Case vbLf, vbCr
                    szTrimCRLF = Left$(szString, lStr - 1)
                Case Else
                    szTrimCRLF = szString
            End Select
        End If
    End If
        

End Function

Private Function VerifyAdd(strAddress As String) As Boolean

Dim chPos As Integer
Dim strTemp As String
Dim strExt As String

chPos = InStr(strAddress, "@")

If chPos Then
 'Separate the domain from the address
 strTemp = Mid$(strAddress, chPos + 1)
 'Now check the domain for availabality of '.'
 If Mid$(strTemp, Len(strTemp) - 2, 1) = "." Then
  'For address that are country based. e.g xxx.com.my
  strExt = Mid$(strTemp, Len(strTemp) - 5, 3) 'CHECK IT
  If UCase$(strExt) = "COM" Or UCase$(strExt) = "NET" Or UCase$(strExt) = "ORG" Or UCase$(strExt) = "EDU" Or UCase$(strExt) = "GOV" Then
   'YEs it is // Unfinished
   VerifyAdd = True
   Exit Function
  End If
  'No
   VerifyAdd = False
   Exit Sub
 ElseIf Mid$(strTemp, Len(strTemp) - 3, 1) = "." Then
  'For international based address. e.g xxx.com , xxx.net , xxx.org
  strExt = Mid$(strTemp, Len(strTemp) - 2, 3) 'CHECK IT
  If UCase$(strExt) = "COM" Or UCase$(strExt) = "NET" Or UCase$(strExt) = "ORG" Or UCase$(strExt) = "EDU" Or UCase$(strExt) = "GOV" Then
   'YEs it is // Unfinished
   VerifyAdd = True
   Exit Sub
  End If
   'No
   VerifyAdd = False
   Exit Sub
 Else
  'Unknown.Considered invalid.
  VerifyAdd = False
 End If
Else
 VerifyAdd = False 'Not even has the '@' symbol
End If


End Function

Private Function FindAdd(ByVal strMessage As String) As String

Dim strBuffer As String
Dim strTemp As String
Dim chrPos As Integer
Dim strAdd As String

strBuffer = szTrimCRLF(strMessage) 'Remove Line feeds and Carriage returns

'Put a space at the end

strBuffer = strBuffer & " "

Do

'Take a part of message separated by space
chrPos = InStr(strBuffer, " ")
strTemp = Left$(strBuffer, chrPos - 1)
strBuffer = Mid$(strBuffer, chrPos + 1)

'Check if it is an address
chrPos = InStr(strTemp, "@")

If chrPos Then
 'It is an email address ( Though not verified in terms of username@domain.extension format
 strAdd = strAdd & "," & strTemp
End If

Loop Until Trim$(strBuffer) <> ""

'Remove ',' from addlists

FindAdd = Right$(strAdd, Len(strAdd) - 1)

End Function
