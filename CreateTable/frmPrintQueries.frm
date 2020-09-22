VERSION 5.00
Begin VB.Form frmPrintQueries 
   Caption         =   "PRINT QUERIES"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2805
      Left            =   315
      TabIndex        =   5
      Top             =   630
      Width           =   2805
      Begin VB.ListBox lstTables 
         Height          =   2535
         Left            =   45
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   180
         Width           =   2715
      End
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   2430
      TabIndex        =   4
      Top             =   360
      Width           =   555
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1755
      TabIndex        =   1
      Top             =   3645
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   675
      TabIndex        =   0
      Top             =   3645
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Select Queries to Print"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1905
   End
   Begin VB.Label lblTables 
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   4185
      Width           =   1905
   End
End
Attribute VB_Name = "frmPrintQueries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    Call GetQueries
End Sub

'**************************************************************************
'                       GET DATABASE OBJECTS AND DISPLAY
'**************************************************************************
Private Function GetQueries()
Dim i As Integer
    
    Dim QueryObj1 As DAO.QueryDef
    i = 0
    lblTables.Caption = "Wait...."
    lstTables.Clear
     
    For Each QueryObj1 In obDB.QueryDefs
          lstTables.AddItem QueryObj1.Name, i
          i = i + 1
    Next QueryObj1
    If lstTables.ListCount > 0 Then lstTables.ListIndex = 0
    lblTables.Caption = "Total Queries = " & obDB.QueryDefs.Count
    Set QueryObj1 = Nothing
End Function
Private Sub chkAll_Click()
    If chkAll.Value = vbChecked Then
        lstTables.Enabled = False
    Else
        lstTables.Enabled = True
    End If
End Sub

Private Sub cmdPrint_Click()
Dim fs, a, i
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(App.Path & "\Queries.htm") Then
        Kill (App.Path & "\Queries.htm")
    End If
    Set a = fs.CreateTextFile(App.Path & "\Queries.htm", True)
    a.WriteLine "<html>"
    
    a.WriteLine "<Head><u> <h4>" & obDB.Name & "</h4></u></head>"
    a.WriteLine "<body>"
    For i = 0 To lstTables.ListCount - 1
        If chkAll.Value = vbChecked Then
            a.WriteLine "<br><br><b><u>" & lstTables.List(i) & "</u></b>"
            a.WriteLine "<br><br>" & obDB.QueryDefs(Trim(lstTables.List(i))).SQL
        Else
            If lstTables.Selected(i) Then
                a.WriteLine "<br><br><b>" & lstTables.List(i) & "</b>"
                a.WriteLine "<br><br>" & obDB.QueryDefs(Trim(lstTables.List(i))).SQL
            End If
        End If
    Next i
    a.WriteLine "</body> </HTml>"
    a.Close
    
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & App.Path & "\Queries.htm", vbMaximizedFocus
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

