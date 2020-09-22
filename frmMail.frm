VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMail 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "QuickMail"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7155
   Icon            =   "frmMail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTo 
      Caption         =   "&To"
      Height          =   300
      Left            =   150
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cboPriority 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txt_email_to 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   6015
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   4965
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12118
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   5760
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdAtt 
      Caption         =   "..."
      Height          =   280
      Left            =   6720
      TabIndex        =   4
      ToolTipText     =   " Select Attachment File "
      Top             =   1440
      Width           =   375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6360
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox txt_message_text 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2280
      Width           =   7095
   End
   Begin VB.TextBox txt_subject 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   6015
   End
   Begin VB.TextBox txt_attach 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label txt_email_from 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Priority"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1875
      Width           =   645
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1470
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1110
      Width           =   885
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   400
      Width           =   645
   End
   Begin VB.Menu mnu_Send 
      Caption         =   "&Send"
   End
   Begin VB.Menu mnu_Contacts 
      Caption         =   "&Contacts"
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strEncodedAtt As String
Dim strResponse As String
Dim blnTimeOut As Boolean

Private Sub cmdTo_Click()
frmSelect.Show (vbModal)
End Sub

Private Sub mnu_Contacts_Click()
frmContacts.Show
End Sub

Private Sub Form_Activate()
Me.txt_email_from.Caption = gstrSenderMail
End Sub

Private Sub Form_Load()
Me.cboPriority.AddItem "High"
Me.cboPriority.AddItem "Normal"
Me.cboPriority.AddItem "Low"
Me.cboPriority.ListIndex = 1
Me.StatusBar1.Panels(1).Text = "Ready."
strResponse = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
SaveSetting App.EXEName, "Config", "Dir", gstrCurrPath
If gblnContactChanged = True Then Call SaveContacts
Unload frmOptions
Unload frmContacts
Unload frmNewContact
Unload frmSelect
Unload Me
End Sub

Private Sub mnu_Send_Click()
Dim k As Integer
'sender
If gstrSMTPserver = "" Then MsgBox "Please enter the SMTP-server in the options window.", vbCritical: Exit Sub
If gstrSenderMail = "" Then MsgBox "Please enter the Senders Mailadress in the options window.", vbCritical: Exit Sub
'subject
Me.txt_subject.Text = Trim(Me.txt_subject.Text)
If Me.txt_subject.Text = "" Then Me.txt_subject.Text = "None"
'attachment
Me.txt_attach.Text = Trim(Me.txt_attach.Text)
If Me.txt_attach.Text <> "" And FileExist(Me.txt_attach.Text) = False Then MsgBox "Attachment File not found.", vbCritical: Exit Sub
'mail to
If InStr(1, Me.txt_email_to.Text, "@") = 0 Or InStr(1, Me.txt_email_to.Text, ".") = 0 Then MsgBox "Please enter a valid mail adress in the 'To' field", vbExclamation:  Exit Sub
StartTrim:
Me.txt_email_to.Text = Trim(Me.txt_email_to.Text)
If Len(Me.txt_email_to.Text) < 3 Then MsgBox "Please enter only ; signs between adresses in the 'To' field", vbExclamation: Call MenusON: Exit Sub
If Left(Me.txt_email_to, 1) = ";" Or Right(Me.txt_email_to, 1) = ";" Then MsgBox "Please enter only ; signs between adresses in the 'To' field", vbExclamation: Exit Sub
For k = 1 To Len(Me.txt_email_to)
    If Mid(Me.txt_email_to, k, 2) = ";;" Then MsgBox "Please enter only one ; sign between adresses in the 'To' field", vbExclamation: Exit Sub
Next k
'content
If Me.txt_message_text.Text = "" And Me.txt_attach.Text = "" Then MsgBox "No information to send.", vbExclamation: Exit Sub

Call MenusOFF

strEncodedAtt = ""
Me.StatusBar1.Panels(1).Text = "Encoding Attachment..."
If Me.txt_attach.Text <> "" Then strEncodedAtt = EncodeAttToBase64(txt_attach.Text)
Me.StatusBar1.Panels(1).Text = "Connecting..."
Call ConnectToServer(gstrSMTPserver)
'Call MenusON
End Sub

Private Sub mnu_Options_Click()
frmOptions.Show vbModal
End Sub

Sub ConnectToServer(smtp_server As String)
On Error GoTo errHandlerConnect
Winsock1.LocalPort = 0
Winsock1.RemoteHost = gstrSMTPserver
Winsock1.RemotePort = 25
Winsock1.Connect
Exit Sub

errHandlerConnect:
Call MenusON
MsgBox "Failed sending the file." & vbCrLf & vbCrLf & Err.Description
Me.StatusBar1.Panels(1).Text = "Error."
Err.Clear
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
Me.StatusBar1.Panels(1).Text = "Connected to: " & gstrSMTPserver & "."
Call SendMail
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData strResponse
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Me.StatusBar1.Panels(1).Text = "Error: " & Description & "."
Call MenusON
End Sub

Private Function GetServerResponse(strAnswer As String, intdelay As Integer) As Boolean
Dim lngTimeStart
blnTimeOut = False
lngTimeStart = Now
Do While Left(strResponse, 3) <> strAnswer
    DoEvents
    ' wait maximum intdelay (seconds), after that: timeout
    If DateDiff("s", lngTimeStart, Now) > intdelay Then blnTimeOut = True: Exit Do
Loop
strResponse = ""
GetServerResponse = Not blnTimeOut
End Function

Sub SendMail()

Dim se_body As String
Dim se_type_content As String
Dim se_type_message As String
Dim se_type_attach As String
Dim se_priority As String
Dim se_email_from As String
Dim se_email_to As String
Dim se_subject As String
Dim se_message_text As String
Dim se_attach As String
Dim se_sender_name As String
Dim se_sender_ip As String
Dim se_next_to As String
Dim se_warn As String
Dim npos As Integer
Const boundary = "simpleboundary"
    
On Error GoTo errHandlerSendMail
    
Select Case Me.cboPriority.Text
Case "High"
    se_priority = "2"
Case "Normal"
    se_priority = "3"
Case "Low"
    se_priority = "4"
End Select

se_sender_name = gstrSenderName
If se_sender_name = "" Then se_sender_name = gstrSenderMail
se_sender_ip = gstrSenderIP
If se_sender_ip = "" Then se_sender_ip = "x.x.x.x"
se_email_from = gstrSenderMail
se_email_to = Me.txt_email_to
se_subject = Me.txt_subject
se_attach = Me.txt_attach
se_subject = Me.txt_subject
            
se_warn = vbCrLf _
    & "------------------------------------------------------------" & vbCrLf _
    & "  The senders name, mail-adress or IP could be changed" & vbCrLf _
    & "  and may not represent his true identity. If this mail" & vbCrLf _
    & "  is abusive, contains spam, or is sent with unlawfull" & vbCrLf _
    & "  intension you should report this to your provider." & vbCrLf _
    & "------------------------------------------------------------"
    
If frmOptions.chkWarn.Value = 1 Then
    se_message_text = Me.txt_message_text & vbCrLf & vbCrLf & vbCrLf & vbCrLf & se_warn
    Else
    se_message_text = Me.txt_message_text
    End If
    
se_type_message = "This is a multi-part message in MIME format." & vbCrLf & vbCrLf _
    & se_warn & vbCrLf & vbCrLf _
    & "--" & boundary & vbCrLf _
    & "Content-Type: text/plain;" & " charset=" & """" & "iso-8859-1" & """" & vbCrLf _
    & "Content-Transfer-Encoding: 7bit"
        
If Len(se_attach) > 0 Then
    se_type_attach = "--" & boundary & vbCrLf _
        & "Content-Type: application/octet-stream;" & " name=" & """" & GetAttFileName(se_attach) & """" & vbCrLf _
        & "Content-Transfer-Encoding: base64" & vbCrLf _
        & "Content-Disposition: attachment;" & " filename=" & """" & GetAttFileName(se_attach) & """" & vbCrLf _
        & vbCrLf _
        & strEncodedAtt
End If

se_body = "X-Originating-IP: [" & se_sender_ip & "]" & vbCrLf _
    & "X-Originating-Email: [" & se_email_from & " ]" & vbCrLf _
    & "X-Sender: " & se_email_from & vbCrLf _
    & "X-Priority: " & se_priority & vbCrLf _
    & "From: " & """" & se_sender_name & """" & " <" & se_email_from & ">" & vbCrLf _
    & "To: " & se_email_to & vbCrLf _
    & "Subject: " & se_subject & vbCrLf _
    & "Date: " & GetDateFormat & vbCrLf _
    & "MIME-Version: 1.0" & vbCrLf _
    & "Content-Type: multipart/mixed;" & " boundary =" & """" & boundary & """" & vbCrLf _
    & vbCrLf _
    & se_type_message & vbCrLf _
    & vbCrLf _
    & se_message_text & vbCrLf _
    & vbCrLf _
    & se_type_attach & vbCrLf _
    & "." & vbCrLf
'Debug.Print se_body
Me.StatusBar1.Panels(1).Text = "Sending message..."

'send hallo (max timeout 30 seconds)
Winsock1.SendData "HELO " & Left(se_email_from, InStr(1, se_email_from, "@") - 1) & vbCrLf
If GetServerResponse("250", 30) = False Then GoTo errHandlerSendMail

'send from
Winsock1.SendData "MAIL FROM: " & se_email_from & vbCrLf
If GetServerResponse("250", 20) = False Then GoTo errHandlerSendMail

'send all to's
Do While InStr(1, se_email_to, ";") <> 0
    npos = InStr(1, se_email_to, ";")
    se_next_to = Trim(Left(se_email_to, npos - 1))
    se_email_to = Trim(Mid(se_email_to, npos + 1))
    If se_next_to <> "" Then
        Winsock1.SendData "RCPT TO: " & se_next_to & vbCrLf '!!!!!!!!!
        If GetServerResponse("250", 20) = False Then GoTo errHandlerSendMail
    End If
Loop
If se_email_to <> "" Then Winsock1.SendData "RCPT TO: " & se_email_to & vbCrLf
If GetServerResponse("250", 20) = False Then GoTo errHandlerSendMail

'send data request
Winsock1.SendData "DATA" & vbCrLf
If GetServerResponse("354", 20) = False Then GoTo errHandlerSendMail

'send the data (max timeout = 600 seconds)
Winsock1.SendData se_body
If GetServerResponse("250", 600) = False Then GoTo errHandlerSendMail

'send quit
Winsock1.SendData "QUIT" & vbCrLf
If GetServerResponse("221", 20) = False Then GoTo errHandlerSendMail

'close
Me.StatusBar1.Panels(1).Text = "Message sent."
Call MenusON
Winsock1.Close
DoEvents
Exit Sub
    
errHandlerSendMail:
Call MenusON
If blnTimeOut = False Then
    MsgBox "Failed sending the file:" & vbCrLf & vbCrLf & Err.Description, vbCritical
    Else
    MsgBox "Failed sending the file: Timeout Error.", vbCritical
    End If
Me.StatusBar1.Panels(1).Text = "Error."
Err.Clear
Winsock1.Close
Exit Sub

End Sub

Private Sub cmdAtt_Click()
On Error Resume Next
ComDlg.InitDir = gstrCurrPath
ComDlg.Filter = "All Files (*.*)|*.*"
ComDlg.DialogTitle = " Select Attachment File"
ComDlg.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
ComDlg.ShowOpen
If Err = 32755 Then Exit Sub
gstrCurrPath = ComDlg.FileName
Me.txt_attach.Text = ComDlg.FileName
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
If Me.Width < 4000 Then Me.Width = 4000: Exit Sub
Me.txt_email_from.Width = Me.Width - 1300
Me.txt_email_to.Width = Me.Width - 1300
Me.txt_subject.Width = Me.Width - 1300
Me.txt_attach.Width = Me.Width - 1800
Me.cmdAtt.Left = Me.Width - 600
Me.txt_message_text.Width = Me.Width - 100
If Me.Height < 4000 Then Me.Height = 4000: Exit Sub
Me.txt_message_text.Height = Me.Height - 3250
End Sub

Private Sub MenusOFF()
Me.cmdTo.Enabled = False
Me.mnu_Options.Enabled = False
Me.mnu_Send.Enabled = False
Me.cmdAtt.Enabled = False
Me.txt_attach.Enabled = False
Me.txt_email_from.Enabled = False
Me.txt_email_to.Enabled = False
Me.txt_message_text.Enabled = False
Me.txt_subject.Enabled = False
Me.cboPriority.Enabled = False
End Sub

Private Sub MenusON()
Me.cmdTo.Enabled = True
Me.mnu_Options.Enabled = True
Me.mnu_Send.Enabled = True
Me.cmdAtt.Enabled = True
Me.cmdAtt.Enabled = True
Me.txt_attach.Enabled = True
Me.txt_email_from.Enabled = True
Me.txt_email_to.Enabled = True
Me.txt_message_text.Enabled = True
Me.txt_subject.Enabled = True
Me.cboPriority.Enabled = True
Me.txt_message_text.SetFocus
End Sub

Private Sub txt_attach_GotFocus()
Me.txt_attach.SelStart = 0
Me.txt_attach.SelLength = Len(Me.txt_attach.Text)
End Sub

Private Sub txt_email_to_GotFocus()
Me.txt_email_to.SelStart = 0
Me.txt_email_to.SelLength = Len(Me.txt_email_to.Text)
End Sub

Private Sub txt_subject_GotFocus()
Me.txt_subject.SelStart = 0
Me.txt_subject.SelLength = Len(Me.txt_subject.Text)
End Sub

