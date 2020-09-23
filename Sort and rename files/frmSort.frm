VERSION 5.00
Begin VB.Form frmSort 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort and rename files"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "frmSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkK 
      Caption         =   "Name 100#"
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Change visible fileformates"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Update the lists"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "Add all in an alphabetical order"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CheckBox chkPar 
      Caption         =   "Name (#)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.ListBox lstResults 
      Height          =   1425
      ItemData        =   "frmSort.frx":08CA
      Left            =   120
      List            =   "frmSort.frx":08CC
      TabIndex        =   8
      Top             =   1995
      Width           =   2415
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change name"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   2415
   End
   Begin VB.ListBox lstWrongOrder 
      Height          =   5130
      ItemData        =   "frmSort.frx":08CE
      Left            =   2640
      List            =   "frmSort.frx":08D0
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2415
   End
   Begin VB.DirListBox Folder 
      Height          =   1890
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   2415
   End
   Begin VB.FileListBox File 
      Height          =   285
      Left            =   2520
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   3600
      Width           =   2415
      Visible         =   0   'False
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox lstOrder 
      Height          =   5130
      ItemData        =   "frmSort.frx":08D2
      Left            =   5160
      List            =   "frmSort.frx":08D4
      TabIndex        =   7
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblInstrAll 
      Caption         =   "Click on the files in the right order:"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblInstrOrder 
      Caption         =   "The files in the right order:"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblInstrName 
      Caption         =   "Name on all files:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkK_Click()
txtName.SetFocus
End Sub

Private Sub chkPar_Click()
txtName.SetFocus
End Sub

Private Sub cmdAll_Click()
Dim i As Long
Dim strAdd As String
For i = 0 To (File.ListCount - 1)
strAdd = File.List(i)
lstOrder.AddItem strAdd
Next i
lstWrongOrder.Clear
txtName.SetFocus
End Sub

Private Sub cmdChange_Click()

Dim i As Long
Dim strPathNotDone As String
Dim strPathDone As String

If chkPar.Value = 0 And chkK.Value = 0 Then

  For i = 0 To (lstOrder.ListCount - 1)
    strPathNotDone = File.Path & "\" & lstOrder.List(i)
    strPathDone = File.Path & "\" & txtName.Text & " " & (i + 1) & (Right(lstOrder.List(i), 4))
    FileCopy strPathNotDone, strPathDone
    Kill strPathNotDone
  Next i

ElseIf chkPar.Value = 1 And chkK.Value = 0 Then

  For i = 0 To (lstOrder.ListCount - 1)
    strPathNotDone = File.Path & "\" & lstOrder.List(i)
    strPathDone = File.Path & "\" & txtName.Text & " (" & (i + 1) & ")" & (Right(lstOrder.List(i), 4))
    FileCopy strPathNotDone, strPathDone
    Kill strPathNotDone
  Next i

ElseIf chkPar.Value = 0 And chkK.Value = 1 Then
  
  For i = 0 To (lstOrder.ListCount - 1)
    strPathNotDone = File.Path & "\" & lstOrder.List(i)
    strPathDone = File.Path & "\" & txtName.Text & " " & (i + 1001) & (Right(lstOrder.List(i), 4))
    FileCopy strPathNotDone, strPathDone
    Kill strPathNotDone
  Next i

ElseIf chkPar.Value = 1 And chkK.Value = 1 Then

  For i = 0 To (lstOrder.ListCount - 1)
    strPathNotDone = File.Path & "\" & lstOrder.List(i)
    strPathDone = File.Path & "\" & txtName.Text & " (" & (i + 1001) & ")" & (Right(lstOrder.List(i), 4))
    FileCopy strPathNotDone, strPathDone
    Kill strPathNotDone
  Next i

End If

File.Refresh
AddResultat
lstResults.Refresh
lstResults.Visible = True
txtName.Text = ""
txtName.SetFocus

End Sub

Private Sub AddResultat()
Dim i As Long
Dim strAdd As String
For i = 0 To (File.ListCount - 1)
strAdd = File.List(i)
lstResults.AddItem strAdd
Next i
End Sub

Private Sub cmdFile_Click()
txtName.SetFocus
frmFile.Show
Me.Hide
End Sub

Private Sub cmdRefresh_Click()
lstWrongOrder.Clear
lstOrder.Clear
lstResults.Clear
File.Refresh
lstResults.Visible = False
Add
txtName.Text = ""
txtName.SetFocus
End Sub

Private Sub Drive_Change()
Folder.Path = Drive.Drive
End Sub

Private Sub Folder_Change()
File.Path = Folder.Path
lstWrongOrder.Clear
Add
End Sub

Private Sub Form_Load()
Folder.Path = App.Path
File.Path = Folder.Path
lstWrongOrder.Clear
Add
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub lstWrongOrder_Click()
Dim lngListIndex As Long
Dim strAdd As String
lngListIndex = lstWrongOrder.ListIndex
strAdd = lstWrongOrder.List(lngListIndex)
lstOrder.AddItem strAdd
lstWrongOrder.RemoveItem (lngListIndex)
txtName.SetFocus
End Sub

Private Sub Path_Click()
File.Path = Folder.Path
lstWrongOrder.Clear
Add
End Sub

Private Sub Add()
Dim i As Long
Dim strAdd As String
For i = 0 To (File.ListCount - 1)
strAdd = File.List(i)
lstWrongOrder.AddItem (strAdd)
Next i
End Sub

Private Sub lstOrder_Click()
Dim lngListIndex As Long
Dim strAdd As String
lngListIndex = lstOrder.ListIndex
strAdd = lstOrder.List(lngListIndex)
lstWrongOrder.AddItem strAdd
lstOrder.RemoveItem (lngListIndex)
txtName.SetFocus
End Sub

Private Sub opt0_Click()
txtName.SetFocus
End Sub

Private Sub opt00_Click()
txtName.SetFocus
End Sub

Private Sub opt000_Click()
txtName.SetFocus
End Sub

Private Sub opt0000_Click()
txtName.SetFocus
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
cmdChange_Click
End If
End Sub
