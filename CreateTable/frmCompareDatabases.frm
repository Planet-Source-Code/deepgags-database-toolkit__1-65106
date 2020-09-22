VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCompareDatabases 
   Caption         =   "COMPARE DATABASE"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   12600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   8595
      Left            =   90
      TabIndex        =   10
      Top             =   180
      Width           =   12405
      Begin VB.CommandButton cmdRemove 
         Caption         =   "<<<  Remove"
         Height          =   375
         Left            =   10830
         TabIndex        =   9
         Top             =   7560
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add   >>>"
         Height          =   375
         Left            =   495
         TabIndex        =   8
         Top             =   7515
         Width           =   1185
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "REFRESH"
         Height          =   375
         Left            =   9510
         TabIndex        =   6
         Top             =   8115
         Width           =   1245
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT"
         Height          =   375
         Left            =   10770
         TabIndex        =   7
         Top             =   8115
         Width           =   1245
      End
      Begin VB.TextBox txtPass1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1380
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1020
         Width           =   6120
      End
      Begin VB.TextBox txtFile1 
         Height          =   315
         Left            =   1350
         TabIndex        =   0
         Top             =   390
         Width           =   8460
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "...."
         Height          =   315
         Index           =   0
         Left            =   9855
         TabIndex        =   1
         Top             =   405
         Width           =   375
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Compare Files"
         Height          =   375
         Left            =   8505
         TabIndex        =   3
         Top             =   1035
         Width           =   1695
      End
      Begin VB.ListBox lstFile1 
         Height          =   5520
         Left            =   480
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1935
         Width           =   5715
      End
      Begin VB.ListBox lstFile2 
         Height          =   5520
         Left            =   6915
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1935
         Width           =   5175
      End
      Begin VB.Label Label5 
         Caption         =   "Password"
         Height          =   195
         Left            =   1380
         TabIndex        =   14
         Top             =   750
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "DATABSE PATH TO COMPARE WITH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1380
         TabIndex        =   13
         Top             =   150
         Width           =   3750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   12270
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "EXTRA OBJECTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6870
         TabIndex        =   12
         Top             =   1665
         Width           =   2400
      End
      Begin VB.Label Label4 
         Caption         =   "MISSING OBJECTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   495
         TabIndex        =   11
         Top             =   1665
         Width           =   2400
      End
   End
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmCompareDatabases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim obDAO1 As DAO.Workspace, obDB1 As DAO.Database
Private Sub cmdBrowse_Click(Index As Integer)
On Error GoTo Errhandler
    
    cdgFile.Filter = "MS Access (*.mdb)|*.mdb|"
    cdgFile.ShowOpen
    txtFile1.Text = cdgFile.FileName
    
Exit Sub
    
Errhandler:
End Sub

Private Sub cmdCompare_Click()
    
    Dim PassString1 As String, PassString2 As String
    
    Set obDAO1 = DAO.DBEngine.Workspaces(0)
    
    On Error GoTo Errmsg
    Set obDB1 = obDAO1.OpenDatabase(txtFile1.Text, False, False, ";pwd=" & txtPass1 & "")
    
    lstFile1.Clear
    lstFile2.Clear
       
    Call Checktables(obDB, obDB1, 1)
    Call Checktables(obDB1, obDB, 2)
    Call CheckQueries(obDB, obDB1, 1)
    Call CheckQueries(obDB1, obDB, 2)
    
'    obDB1.Close
'    obDAO1.Close
'    Set obDB1 = Nothing
'    Set obDAO1 = Nothing
    
Exit Sub
Errmsg:
        MsgBox "You need to provide a password.", vbInformation
        txtPass1.SetFocus
    
End Sub

Private Function Checktables(DBObject1 As DAO.Database, DBObject2 As DAO.Database, flag As Integer) As Boolean
    Dim tableObj1 As DAO.TableDef
    Dim TableObj2 As DAO.TableDef
    Dim TabFound As Boolean
    
    For Each tableObj1 In DBObject1.TableDefs
        TabFound = False
        For Each TableObj2 In DBObject2.TableDefs
            Debug.Print tableObj1.Name & "= " & TableObj2.Name
             If tableObj1.Name = TableObj2.Name Then
                Call CheckFields(tableObj1, TableObj2, flag)
                TabFound = True
                Exit For
            End If
        Next TableObj2
        If TabFound = False Then
            If flag = 1 Then
                lstFile2.AddItem "Table  - " & tableObj1.Name
            Else
                lstFile1.AddItem "Table  - " & tableObj1.Name
            End If
        End If
    Next tableObj1
    Set tableObj1 = Nothing
    Set TableObj2 = Nothing
End Function
Private Function CheckFields(DBField1 As DAO.TableDef, DBField2 As DAO.TableDef, flag As Integer) As Boolean
    Dim FieldObj1 As DAO.Field
    Dim FieldObj2 As DAO.Field
    Dim TabFound As Boolean
    For Each FieldObj1 In DBField1.Fields
        TabFound = False
        For Each FieldObj2 In DBField2.Fields
             If FieldObj1.Name = FieldObj2.Name Then
                TabFound = True
                Exit For
            End If
        Next FieldObj2
        If TabFound = False Then
            If flag = 1 Then
                lstFile2.AddItem "Field  - " & DBField1.Name & "-->" & FieldObj1.Name
            Else
                lstFile1.AddItem "Field  - " & DBField1.Name & "-->" & FieldObj1.Name
            End If
        End If
    Next FieldObj1
    Set FieldObj1 = Nothing
    Set FieldObj2 = Nothing
End Function

Private Function CheckQueries(DBObject1 As DAO.Database, DBObject2 As DAO.Database, flag As Integer) As Boolean
    Dim QueryObj1 As DAO.QueryDef
    Dim QueryObj2 As DAO.QueryDef
    Dim TabFound As Boolean
    For Each QueryObj1 In DBObject1.QueryDefs
        TabFound = False
        For Each QueryObj2 In DBObject2.QueryDefs
            Debug.Print QueryObj1.Name & "= " & QueryObj2.Name
             If QueryObj1.Name = QueryObj2.Name Then
                'Call CheckFields(QueryObj1, QueryObj2, flag)
                TabFound = True
                Exit For
            End If
        Next QueryObj2
        If TabFound = False Then
            If flag = 1 Then
                lstFile2.AddItem "Query  - " & QueryObj1.Name
            Else
                lstFile1.AddItem "Query  - " & QueryObj1.Name
            End If
        End If
    Next QueryObj1
    Set QueryObj1 = Nothing
    Set QueryObj2 = Nothing
End Function

Private Sub cmdPrint_Click()
    Dim fs As New FileSystemObject, st As TextStream, FileName As String
    FileName = InputBox("Enter File Name")
    If FileName <> "" Then
        fs.CreateTextFile (App.Path & "\" & FileName & ".txt")
        Set st = fs.OpenTextFile(App.Path & "\" & FileName & ".txt", ForAppending)
        st.WriteLine "//MISSING OBJECTS" & vbNewLine
        For i = 0 To lstFile1.ListCount - 1
            st.WriteLine (lstFile1.List(i))
        Next i
        st.WriteLine vbNewLine & "//EXTRA OBJECTS" & vbNewLine
        For i = 0 To lstFile2.ListCount - 1
            st.WriteLine (lstFile2.List(i))
        Next i
        st.Close
        Set fs = Nothing
        MsgBox "File Created"
    End If
End Sub

Private Sub cmdRefresh_Click()
    cmdCompare_Click
End Sub

Private Sub cmdRemove_Click()
Dim str As String, arr() As String
    str = (lstFile2.Text)
    arr = Split(str, " -")
    If Trim(arr(0)) = "Field" Then
        arr = Split(arr(1), "-->")
        obDB.TableDefs(Trim(arr(0))).Fields.Delete (Trim(arr(1)))
    ElseIf Trim(arr(0)) = "Query" Then
        obDB.QueryDefs.Delete (Trim(arr(1)))
    ElseIf Trim(arr(0)) = "Table" Then
        obDB.TableDefs.Delete (Trim(arr(1)))
    End If
    lstFile2.RemoveItem (lstFile2.ListIndex)
    MsgBox arr(0) & " " & arr(1) & " removed  from database."
    
End Sub

Private Sub cmdAdd_Click()
Dim str As String, arr() As String
Dim QuerySQL As String
    str = (lstFile1.Text)
    arr = Split(str, " -")
    If Trim(arr(0)) = "Field" Then
        arr = Split(arr(1), "-->")
        Call AddField(arr(0), arr(1), obDB, obDB1)
    ElseIf Trim(arr(0)) = "Query" Then
        Call TransferQuery(obDB1, Trim(arr(1)), obDB)
    ElseIf Trim(arr(0)) = "Table" Then
        Call TransferTable(obDB1, Trim(arr(1)), obDB)
    End If
    lstFile1.RemoveItem (lstFile1.ListIndex)
    MsgBox arr(0) & " " & arr(1) & " created in the database."
    
End Sub
