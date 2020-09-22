VERSION 5.00
Begin VB.Form frmNewContact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " New Contact"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmNewContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtContactAdres 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtContactName 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mail Adress"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Name"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
If gblnFormChange = True Then
    frmContacts.ListContacts.List(frmContacts.ListContacts.ListIndex) = Trim(Me.txtContactName.Text) & "  <" & Trim(Me.txtContactAdres.Text) & ">"
    Else
    frmContacts.ListContacts.AddItem Trim(Me.txtContactName.Text) & "  <" & Trim(Me.txtContactAdres.Text) & ">"
    frmContacts.ListContacts.Selected(frmContacts.ListContacts.ListCount - 1) = True
    End If
gblnContactChanged = True
Me.Hide
End Sub

Private Sub Form_Activate()
If gblnFormChange = True Then
    Me.cmdOK.Enabled = True
    Me.Caption = " Change Contact"
    Me.txtContactName.Text = GetMailName(frmContacts.ListContacts.Text)
    Me.txtContactAdres.Text = GetMailAdres(frmContacts.ListContacts.Text)
    Else
    Me.cmdOK.Enabled = False
    Me.Caption = " New Contact"
    Me.txtContactName.Text = ""
    Me.txtContactAdres.Text = ""
    End If
Me.txtContactName.SetFocus
End Sub

Private Sub txtContactAdres_Change()
If Me.txtContactAdres.Text = "" Then
    Me.cmdOK.Enabled = False
    Else
    Me.cmdOK.Enabled = True
    End If
End Sub

Private Sub txtContactAdres_GotFocus()
Me.txtContactAdres.SelStart = 0
Me.txtContactAdres.SelLength = Len(Me.txtContactAdres.Text)
End Sub

Private Sub txtContactName_GotFocus()
Me.txtContactName.SelStart = 0
Me.txtContactName.SelLength = Len(Me.txtContactName.Text)
End Sub

Private Sub txtContactName_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(">") Or KeyAscii = Asc("<") Then KeyAscii = 0
If KeyAscii = 13 Then KeyAscii = 0: Me.txtContactAdres.SetFocus
End Sub

Private Sub txtContactAdres_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(">") Or KeyAscii = Asc("<") Then KeyAscii = 0
If KeyAscii = 13 Then
    KeyAscii = 0
    If Me.txtContactAdres.Text <> "" Then
        Me.cmdOK.SetFocus
        End If
    End If
End Sub

