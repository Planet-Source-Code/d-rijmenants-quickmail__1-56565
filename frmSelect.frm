VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Select Contact..."
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ControlBox      =   0   'False
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   5655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtMailTo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2760
      Width           =   5655
   End
   Begin VB.CommandButton cmdAddToMail 
      Caption         =   "&Add"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox ListSelect 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2400
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddToMail_Click()
If InStr(1, Me.txtMailTo.Text, GetMailAdres(Me.ListSelect.Text)) Then Exit Sub
If Me.txtMailTo.Text = "" Then
    Me.txtMailTo.Text = GetMailAdres(Me.ListSelect.Text)
    Else
    Me.txtMailTo.Text = Me.txtMailTo.Text & " ; " & GetMailAdres(Me.ListSelect.Text)
    End If
Me.txtMailTo.SelStart = Len(Me.txtMailTo.Text)
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
frmMail.txt_email_to.Text = Me.txtMailTo.Text
Me.Hide
End Sub

Private Sub Form_Activate()
Dim k As Integer
Me.txtMailTo.Text = frmMail.txt_email_to.Text
Me.ListSelect.Clear
Dim i As Integer
For i = 0 To UBound(gstrContacts)
    Me.ListSelect.AddItem gstrContacts(i)
Next i
If Me.ListSelect.ListCount > 0 Then
    Me.ListSelect.Selected(0) = True
    Me.cmdAddToMail.Enabled = True
    Me.cmdOK.Enabled = True
    Else
    Me.cmdAddToMail.Enabled = False
    Me.cmdOK.Enabled = False
    End If
End Sub

Private Sub ListSelect_DblClick()
Call cmdAddToMail_Click
End Sub
