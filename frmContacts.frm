VERSION 5.00
Begin VB.Form frmContacts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Contacts"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   ClipControls    =   0   'False
   Icon            =   "frmContacts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5505
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdChange 
      Caption         =   "C&hange"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   5295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ListBox ListContacts 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2790
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer
Me.ListContacts.Clear
For i = 0 To UBound(gstrContacts)
    If gstrContacts(i) <> "" Then Me.ListContacts.AddItem gstrContacts(i)
Next i
End Sub

Private Sub cmdClose_Click()
Call UpdateContacts
Me.Hide
End Sub

Private Sub cmdDel_Click()
Dim i As Integer
retVal = MsgBox("Delete " & Me.ListContacts.Text & " ?", vbQuestion + vbYesNo)
If retVal = vbNo Then Exit Sub
i = Me.ListContacts.ListIndex
Me.ListContacts.RemoveItem (Me.ListContacts.ListIndex)
If Me.ListContacts.ListCount = 0 Then
    Me.cmdDel.Enabled = False
    Me.cmdChange.Enabled = False
    Else
    Me.cmdDel.Enabled = True
    Me.cmdChange.Enabled = True
    End If
If i > 0 Then i = i - 1
If Me.ListContacts.ListCount > 0 Then Me.ListContacts.Selected(i) = True
gblnContactChanged = True
End Sub

Private Sub cmdNew_Click()
gblnFormChange = False
frmNewContact.Show (vbModal)
End Sub

Private Sub cmdChange_Click()
gblnFormChange = True
frmNewContact.Show (vbModal)
End Sub

Private Sub Form_Activate()
If Me.ListContacts.ListCount = 0 Then
    Me.cmdChange.Enabled = False
    Me.cmdDel.Enabled = False
    Else
    Me.cmdChange.Enabled = True
    Me.ListContacts.Selected(0) = True
    Me.cmdDel.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UpdateContacts
End Sub

Private Sub UpdateContacts()
Dim i As Integer
ReDim gstrContacts(Me.ListContacts.ListCount - 1)
For i = 0 To Me.ListContacts.ListCount - 1
    gstrContacts(i) = Me.ListContacts.List(i)
Next i
End Sub
