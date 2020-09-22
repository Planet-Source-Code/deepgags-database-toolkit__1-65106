VERSION 5.00
Begin VB.Form frmPrintTables 
   Caption         =   "PRINT TABLES"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2805
      Left            =   360
      TabIndex        =   9
      Top             =   585
      Width           =   2805
      Begin VB.ListBox lstTables 
         Height          =   2535
         Left            =   45
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   180
         Width           =   2715
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters to Print"
      Height          =   1455
      Left            =   360
      TabIndex        =   5
      Top             =   3420
      Width           =   2805
      Begin VB.CheckBox chkIndexes 
         Caption         =   "Indexes"
         Height          =   240
         Left            =   135
         TabIndex        =   8
         Top             =   1125
         Width           =   915
      End
      Begin VB.CheckBox chkFld 
         Caption         =   "Field Properties"
         Height          =   240
         Left            =   135
         TabIndex        =   7
         Top             =   742
         Width           =   1500
      End
      Begin VB.CheckBox chkTab 
         Caption         =   "Table Properties"
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   360
         Width           =   1680
      End
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "All"
      Height          =   195
      Left            =   2385
      TabIndex        =   4
      Top             =   270
      Width           =   555
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   4950
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   4950
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Select Tables to Print"
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   270
      Width           =   1905
   End
   Begin VB.Label lblTables 
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   5490
      Width           =   1905
   End
End
Attribute VB_Name = "frmPrintTables"
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
    
    Dim tableObj1 As DAO.TableDef
    i = 0
    lblTables.Caption = "Wait...."
    lstTables.Clear
     
    For Each tableObj1 In obDB.TableDefs
          lstTables.AddItem tableObj1.Name, i
          i = i + 1
    Next tableObj1
    If lstTables.ListCount > 0 Then lstTables.ListIndex = 0
    lblTables.Caption = "Total Tables = " & obDB.TableDefs.Count
    Set tableObj1 = Nothing
End Function
Private Sub chkAll_Click()
    If chkAll.Value = vbChecked Then
        lstTables.Enabled = False
    Else
        lstTables.Enabled = True
    End If
End Sub

Private Sub cmdPrint_Click()
Call PrintTables
'Dim fs, a, i
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    If fs.FileExists(App.Path & "\Queries.htm") Then
'        Kill (App.Path & "\Queries.htm")
'    End If
'    Set a = fs.CreateTextFile(App.Path & "\Queries.htm", True)
'    a.WriteLine "<html>"
'
'    a.WriteLine "<Head><u> <b>" & obDB.Name & "</b></u></head>"
'    a.WriteLine "<body>"
'    For i = 0 To lstTables.ListCount - 1
'        If chkAll.Value = vbChecked Then
'            a.WriteLine "<p><b>" & lstTables.List(i) & "</b></p>"
'            a.WriteLine "<p>" & obDB.QueryDefs(Trim(lstTables.List(i))).SQL & "</p>"
'        Else
'            If lstTables.Selected(i) Then
'                a.WriteLine "<p><b>" & lstTables.List(i) & "</b></p>"
'                a.WriteLine "<p>" & obDB.QueryDefs(Trim(lstTables.List(i))).SQL & "</p>"
'            End If
'        End If
'    Next i
'    a.WriteLine "</body> </HTml>"
'    a.Close
'
'    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & App.Path & "\Queries.htm", vbMaximizedFocus
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Sub PrintTables()
Dim fs, a, i
Dim fld As DAO.Field
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(App.Path & "\Tables.htm") Then
        Kill (App.Path & "\Tables.htm")
    End If
    Set a = fs.CreateTextFile(App.Path & "\Tables.htm", True)
    a.WriteLine "<html>"
    a.WriteLine "<Head> <H4><u>" & obDB.Name & "</u></H4></head>"
    a.WriteLine "<body>"
    
    For i = 0 To lstTables.ListCount - 1
        If chkAll.Value = vbChecked Then
                If chkTab.Value = vbChecked Then
                     a.WriteLine "<br><b><u>" & i + 1 & ". " & lstTables.List(i) & "</u></b><br>"
                     a.WriteLine "<br> Date Created : " & obDB.TableDefs(lstTables.List(i)).DateCreated & "<br>"
                Else
                     a.WriteLine "<br><b><u>" & i + 1 & ". " & lstTables.List(i) & "</u></b><br>"
                End If
                'Printing the fields of table
                a.WriteLine "<Table>"
                a.WriteLine "<tr>"
                a.WriteLine "<td width=150><u> Field Name </u></td>"
                a.WriteLine "<td width=100><u> Type </u></td>"
                a.WriteLine "<td width=100 align=right><u> Size </u></td>"
                a.WriteLine "</tr><br>"
                For Each fld In obDB.TableDefs(Trim(lstTables.List(i))).Fields
                    a.WriteLine "<tr>"
                    a.WriteLine "<td width=150>" & fld.Name & "</td>"
                    a.WriteLine "<td width=100>" & GetFldTypeName(fld.Type) & "</td>"
                    a.WriteLine "<td  align=right>" & fld.Size & "</td>"
                    a.WriteLine "</tr>"
                Next
                a.WriteLine "</table>"
        Else
            If lstTables.Selected(i) Then
                If chkTab.Value = vbChecked Then
                     a.WriteLine "<br><b><u>" & i + 1 & ". " & lstTables.List(i) & "</u></b><br>"
                     a.WriteLine "<br> Date Created : " & obDB.TableDefs(lstTables.List(i)).DateCreated & "<br>"
                Else
                     a.WriteLine "<br><b><u>" & lstTables.List(i) & "</u></b><br>"
                End If
                    'Printing the fields of table
                    a.WriteLine "<Table>"
                    a.WriteLine "<tr>"
                    a.WriteLine "<td width=150> Field Name </td>"
                    a.WriteLine "<td width=100> Type </td>"
                    a.WriteLine "<td width=100 align=right> Size </td>"
                    a.WriteLine "</tr><br>"
                    For Each fld In obDB.TableDefs(Trim(lstTables.List(i))).Fields
                        a.WriteLine "<tr>"
                        a.WriteLine "<td width=150>" & fld.Name & "</td>"
                        a.WriteLine "<td width=100>" & GetFldTypeName(fld.Type) & "</td>"
                        a.WriteLine "<td  align=right>" & fld.Size & "</td>"
                        a.WriteLine "</tr>"
                    Next
                    a.WriteLine "</table>"
                End If
        End If
    Next
    a.WriteLine "</body> </HTml>"
    a.Close
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & App.Path & "\Tables.htm", vbMaximizedFocus

End Sub
