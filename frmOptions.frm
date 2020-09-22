VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " QuickMail Options"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Server"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   5775
      Begin VB.TextBox txt_smtp_server 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP server"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sender Settings"
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5775
      Begin VB.CheckBox chkWarn 
         Caption         =   "Add Spam Warning To Mail"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.TextBox txt_default_ip 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txt_default_name 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txt_default_from 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mail-adress"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Originating IP"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Me.txt_default_name.Text = gstrSenderName
Me.txt_default_from.Text = gstrSenderMail
Me.txt_default_ip.Text = gstrSenderIP
Me.txt_smtp_server.Text = gstrSMTPserver
Me.chkWarn.Value = gintSpamWarn
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
gstrSenderName = Me.txt_default_name.Text
gstrSenderMail = Me.txt_default_from.Text
gstrSenderIP = Me.txt_default_ip
gintSpamWarn = Me.chkWarn.Value
gstrSMTPserver = Me.txt_smtp_server.Text
frmMail.txt_email_from.Caption = gstrSenderMail
SaveSetting App.EXEName, "Config", "Name", Me.txt_default_name.Text
SaveSetting App.EXEName, "Config", "From", Me.txt_default_from.Text
SaveSetting App.EXEName, "Config", "IP", Me.txt_default_ip.Text
SaveSetting App.EXEName, "Config", "Warn", Str(gintSpamWarn)
SaveSetting App.EXEName, "Config", "Server", gstrSMTPserver
Me.Hide
End Sub

