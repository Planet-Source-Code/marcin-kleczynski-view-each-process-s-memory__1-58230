VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProcesses 
   Caption         =   "Process Memory Viewer"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Memory Usage"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblItems 
      AutoSize        =   -1  'True
      Caption         =   "0 Processes running."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1485
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()
  cmdRefresh.Enabled = False
    lvwMain.ListItems.Clear
      If GetWindowsVersion = "NT" Then
        GetProcessesNT
      ElseIf GetWindowsVersion = "9X" Then
        GetProcesses9X
      Else
        MsgBox "Your windows version is not supported"
          End
      End If
    lblItems.Caption = lvwMain.ListItems.Count & " Processes running."
  cmdRefresh.Enabled = True
End Sub

Private Sub Form_Load()
  lvwMain.ListItems.Clear
    If GetWindowsVersion = "NT" Then
      GetProcessesNT
    ElseIf GetWindowsVersion = "9X" Then
      GetProcesses9X
    Else
      MsgBox "Your windows version is not supported"
        End
    End If
  lblItems.Caption = lvwMain.ListItems.Count & " Processes running."
End Sub

Private Sub Form_Resize()
  lvwMain.Left = 120
  lvwMain.Top = 120
  
  lvwMain.Width = frmProcesses.Width - (120 * 3)
  lvwMain.Height = frmProcesses.Height - (120 * 3) - (cmdRefresh.Height * 2) - 120

  cmdRefresh.Top = lvwMain.Top + lvwMain.Height + 120
  cmdRefresh.Left = (lvwMain.Left + lvwMain.Width) - cmdRefresh.Width
  
  lblItems.Top = cmdRefresh.Top
  
  'Divide columns by a denomiter of 5 (easier)
  lvwMain.ColumnHeaders.Item(1).Width = (lvwMain.Width / 5) * 2 - 110
  lvwMain.ColumnHeaders.Item(2).Width = (lvwMain.Width / 5) * 1 - 110
  lvwMain.ColumnHeaders.Item(3).Width = (lvwMain.Width / 5) * 2 - 110
End Sub
