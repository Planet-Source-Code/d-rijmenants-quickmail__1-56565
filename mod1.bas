Attribute VB_Name = "Module1"
Public gstrSMTPserver As String
Public gstrCurrPath As String
Public gstrSenderMail As String
Public gstrSenderName As String
Public gstrSenderIP As String
Public gintSpamWarn As Integer
Public gblnContactChanged As Boolean
Public gblnFormChange As Boolean
Public gstrContacts() As String
Public retVal As Integer

Private aDecTab(255) As Integer
Private aEncTab(63) As Byte

Option Explicit

Sub Main()

gstrCurrPath = GetSetting(App.EXEName, "Config", "Dir", "C:\")
gstrSMTPserver = GetSetting(App.EXEName, "Config", "Server")
gstrSenderMail = GetSetting(App.EXEName, "Config", "From")
gstrSenderName = GetSetting(App.EXEName, "Config", "Name")
gstrSenderIP = GetSetting(App.EXEName, "Config", "IP")
gintSpamWarn = Val(GetSetting(App.EXEName, "Config", "Warn", "1"))

Call LoadContacts

Load frmMail
Load frmOptions
Load frmContacts
Load frmNewContact
Load frmSelect

frmMail.Show
frmMail.txt_email_to.SetFocus

End Sub

Public Function FileExist(FileName As String) As Boolean
'checks weither a file exists
    On Error GoTo FileDoesNotExist
    Call FileLen(FileName)
    FileExist = True
    Exit Function
FileDoesNotExist:
    FileExist = False
End Function

Function GetDateFormat() As String
Dim fDay As String
Dim fMonth As String
Dim fTime As String
Select Case WeekDay(Date)
    Case 1: fDay = "Sun"
    Case 2: fDay = "Mon"
    Case 3: fDay = "Tue"
    Case 4: fDay = "Wed"
    Case 5: fDay = "Thu"
    Case 6: fDay = "Fri"
    Case 7: fDay = "Sat"
End Select
Select Case Month(Date)
    Case 1: fMonth = "Jan "
    Case 2: fMonth = "Feb "
    Case 3: fMonth = "Mar "
    Case 4: fMonth = "Apr "
    Case 5: fMonth = "May "
    Case 6: fMonth = "Jun "
    Case 7: fMonth = "Jul "
    Case 8: fMonth = "Aug "
    Case 9: fMonth = "Sep "
    Case 10: fMonth = "Oct "
    Case 11: fMonth = "Nov "
    Case 12: fMonth = "Dec "
End Select
fTime = Format(Time) & " +0200"
GetDateFormat = fDay & ", " & Day(Format(Date)) & " " & fMonth & Year(Format(Date, "dd/mm/yyyy")) & " " & fTime
End Function
Public Sub LoadContacts()
Dim InFIle
Dim LineIn As String
Dim strFilename As String
Dim i As Integer
ReDim gstrContacts(i)
i = 0
InFIle = FreeFile
strFilename = App.Path & "\Contacts.dat"
On Error Resume Next
Open strFilename For Input As InFIle
    While Not EOF(InFIle)
        'get next line
        Line Input #InFIle, LineIn
        If Trim(LineIn) <> "" Then
            ReDim Preserve gstrContacts(i)
            gstrContacts(i) = LineIn: i = i + 1
        End If
    Wend
Close InFIle
End Sub

Public Sub SaveContacts()
Dim InFIle
Dim strContact As String
Dim strFilename As String
Dim k As Integer
InFIle = FreeFile
strFilename = App.Path & "\Contacts.dat"
On Error Resume Next
Open strFilename For Output As InFIle
    For k = 0 To UBound(gstrContacts)
        Print #InFIle, gstrContacts(k)
    Next k
Close InFIle
End Sub

Public Function GetMailName(ByVal contact As String)
Dim pos As Integer
pos = InStr(1, contact, "<")
If pos < 2 Then Exit Function
GetMailName = Trim(Left(contact, pos - 1))
End Function

Public Function GetMailAdres(ByVal contact As String)
Dim pos As Integer
pos = InStr(1, contact, "<")
If pos < 2 Then Exit Function
GetMailAdres = Trim(Mid(contact, pos + 1, Len(contact) - pos - 1))
End Function

Function GetAttFileName(attach_str As String) As String
    Dim s As Integer
    Dim temp As String
    s = InStr(1, attach_str, "\")
    temp = attach_str
    Do While s > 0
        temp = Mid(temp, s + 1, Len(temp))
        s = InStr(1, temp, "\")
    Loop
    GetAttFileName = temp
End Function

Function EncodeAttToBase64(attach_str As String) As String
    Dim FileO       As Integer
    Dim FileBuffer() As Byte
    Dim strBase As String
       
    On Error Resume Next
    FileO = FreeFile
    Open attach_str For Binary As #FileO
        ReDim FileBuffer(0 To LOF(FileO) - 1)
        Get #FileO, , FileBuffer()
    Close #FileO
    
    strBase = StrConv(FileBuffer(), vbUnicode)
    EncodeAttToBase64 = EncodeStr64(strBase, 70) '76 per line
    strBase = ""
End Function

' ------------------------------------------------------------
'                   Base 64 Radix functions
' ------------------------------------------------------------

Private Function PadString(strData As String) As String
' Pad data string to next multiple of 8 bytes
Dim nLen As Long
Dim sPad As String
Dim nPad As Integer
nLen = Len(strData)
nPad = ((nLen \ 8) + 1) * 8 - nLen
sPad = String(nPad, Chr(nPad))
PadString = strData & sPad
End Function

Private Function UnpadString(strData As String) As String
' Strip padding
Dim nLen As Long
Dim nPad As Long
nLen = Len(strData)
If nLen = 0 Then Exit Function
nPad = Asc(Right(strData, 1))
If nPad > 8 Then nPad = 0
UnpadString = Left(strData, nLen - nPad)
End Function


Public Function EncodeStr64(encString As String, ByVal MaxPerLine As Integer) As String
' Return radix64 encoding of string of binary values
Dim abOutput()  As Byte
Dim sLast       As String
Dim B(3)        As Byte
Dim j           As Integer
Dim CharCount   As Integer
Dim iIndex      As Long
Dim Umax        As Long
Dim i As Long, nLen As Long, nQuants As Long
EncodeStr64 = ""
nLen = Len(encString)
nQuants = nLen \ 3
iIndex = 0
If MaxPerLine < 10 Then MaxPerLine = 10
Umax = nQuants + 1
Call MakeEncTab
If (nQuants > 0) Then
    ReDim abOutput(nQuants * 4 - 1)
    For i = 0 To nQuants - 1
        For j = 0 To 2
            B(j) = Asc(Mid(encString, (i * 3) + j + 1, 1))
        Next
        Call EncodeQuantumB(B)
        abOutput(iIndex) = B(0)
        abOutput(iIndex + 1) = B(1)
        abOutput(iIndex + 2) = B(2)
        abOutput(iIndex + 3) = B(3)
        CharCount = CharCount + 4
        ' insert CRLF if max char per line is reached
        If CharCount >= MaxPerLine Then
            ReDim Preserve abOutput(UBound(abOutput) + 2)
            CharCount = 0
            abOutput(iIndex + 4) = 13
            abOutput(iIndex + 5) = 10
            iIndex = iIndex + 6
            Else
            iIndex = iIndex + 4
            End If
    Next
    EncodeStr64 = StrConv(abOutput, vbUnicode)
End If
Select Case nLen Mod 3
Case 0
    sLast = ""
Case 1
    B(0) = Asc(Mid(encString, nLen, 1))
    B(1) = 0
    B(2) = 0
    Call EncodeQuantumB(B)
    sLast = StrConv(B(), vbUnicode)
    ' Replace last 2 with =
    sLast = Left(sLast, 2) & "=="
Case 2
    B(0) = Asc(Mid(encString, nLen - 1, 1))
    B(1) = Asc(Mid(encString, nLen, 1))
    B(2) = 0
    Call EncodeQuantumB(B)
    sLast = StrConv(B(), vbUnicode)
    ' Replace last with =
    sLast = Left(sLast, 3) & "="
End Select
EncodeStr64 = EncodeStr64 & sLast
End Function

Private Sub EncodeQuantumB(B() As Byte)
Dim b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte
b0 = SHR2(B(0)) And &H3F
b1 = SHL4(B(0) And &H3) Or (SHR4(B(1)) And &HF)
b2 = SHL2(B(1) And &HF) Or (SHR6(B(2)) And &H3)
b3 = B(2) And &H3F
B(0) = aEncTab(b0)
B(1) = aEncTab(b1)
B(2) = aEncTab(b2)
B(3) = aEncTab(b3)
End Sub

Private Function MakeDecTab()
' Set up Radix 64 decoding table
Dim t As Integer
Dim c As Integer
For c = 0 To 255
    aDecTab(c) = -1
Next
t = 0
For c = Asc("A") To Asc("Z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("a") To Asc("z")
    aDecTab(c) = t
    t = t + 1
Next
For c = Asc("0") To Asc("9")
    aDecTab(c) = t
    t = t + 1
Next
c = Asc("+")
aDecTab(c) = t
t = t + 1
c = Asc("/")
aDecTab(c) = t
t = t + 1
c = Asc("=")
aDecTab(c) = t
End Function

Private Function MakeEncTab()
' Set up Radix 64 encoding table in bytes
Dim i As Integer
Dim c As Integer
i = 0
For c = Asc("A") To Asc("Z")
    aEncTab(i) = c
    i = i + 1
Next
For c = Asc("a") To Asc("z")
    aEncTab(i) = c
    i = i + 1
Next
For c = Asc("0") To Asc("9")
    aEncTab(i) = c
    i = i + 1
Next
c = Asc("+")
aEncTab(i) = c
i = i + 1
c = Asc("/")
aEncTab(i) = c
i = i + 1
End Function

Private Function SHL2(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 2 bits
SHL2 = (bytValue * &H4) And &HFF
End Function

Private Function SHL4(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 4 bits
SHL4 = (bytValue * &H10) And &HFF
End Function

Private Function SHL6(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to left by 6 bits
SHL6 = (bytValue * &H40) And &HFF
End Function

Private Function SHR2(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 2 bits
SHR2 = bytValue \ &H4
End Function

Private Function SHR4(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 4 bits
SHR4 = bytValue \ &H10
End Function

Private Function SHR6(ByVal bytValue As Byte) As Byte
' Shift 8-bit value to right by 6 bits
SHR6 = bytValue \ &H40
End Function


