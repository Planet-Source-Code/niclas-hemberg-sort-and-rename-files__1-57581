VERSION 5.00
Begin VB.Form frmFile 
   Caption         =   "Change fileformate"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3225
   Icon            =   "frmFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   1620
      Width           =   1815
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Remove a fileformate"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   1190
      Width           =   1815
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add a new fileformate"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   760
      Width           =   1815
   End
   Begin VB.ListBox lstFile 
      Height          =   1620
      ItemData        =   "frmFile.frx":08CA
      Left            =   2040
      List            =   "frmFile.frx":08CC
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblI2 
      Caption         =   "Fileformate(s):"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblI1 
      Caption         =   "Add a new fileformate:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
If txtFile.Text = "" Then
txtFile.SetFocus
Exit Sub
End If
lstFile.AddItem (txtFile.Text)
txtFile.Text = ""
txtFile.SetFocus
End Sub

Private Sub cmdBack_Click()
Dim strPattern As String
Dim i As Integer
If lstFile.ListCount = 0 Then
strPattern = "*.*"
Else
strPattern = "*" & lstFile.List(0)
For i = 1 To (lstFile.ListCount - 1)
strPattern = strPattern & ";*" & lstFile.List(i)
Next i
End If
frmSort.File.Pattern = strPattern
txtFile.SetFocus
frmSort.Show
Me.Hide

'=================================================='

frmSort.lstWrongOrder.Clear
frmSort.lstOrder.Clear
frmSort.lstResults.Clear
frmSort.File.Refresh
frmSort.lstResults.Visible = False
Dim o As Long
Dim strAdd As String
For o = 0 To (frmSort.File.ListCount - 1)
strAdd = frmSort.File.List(o)
frmSort.lstWrongOrder.AddItem (strAdd)
Next o
frmSort.txtName.SetFocus

'=================================================='

End Sub

Private Sub cmdDel_Click()
Dim intListIndex As Integer
Dim strDel As String
txtFile.SetFocus
intListIndex = lstFile.ListIndex
strDel = lstFile.List(intListIndex)
If (intListIndex + 1) = 0 Then
Exit Sub
End If
lstFile.RemoveItem (intListIndex)
txtFile.SetFocus
End Sub

Private Sub lstFile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
cmdDel_Click
End If
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmdAdd_Click
End If
End Sub
