VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "DATABASE TOOLKIT"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   -1440
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   2064
      ButtonWidth     =   2619
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Open"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Refresh"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Import"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Export"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Change Password"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Compare Database"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Retreive Password"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New Table"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New Query"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame FramTables 
      Height          =   9465
      Left            =   -45
      TabIndex        =   33
      Top             =   630
      Width           =   4155
      Begin VB.ListBox lstTables 
         Height          =   8640
         Left            =   270
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblTables 
         Height          =   240
         Left            =   270
         TabIndex        =   35
         Top             =   9090
         Width           =   3570
      End
   End
   Begin VB.Frame Frame2 
      Height          =   9465
      Left            =   4140
      TabIndex        =   34
      Top             =   630
      Width           =   10995
      Begin VB.Frame FramOpt 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   180
         TabIndex        =   53
         Top             =   135
         Width           =   3480
         Begin VB.OptionButton optView 
            Caption         =   "Design"
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
            Index           =   0
            Left            =   1485
            TabIndex        =   3
            Top             =   90
            Width           =   1230
         End
         Begin VB.OptionButton optView 
            Caption         =   "Data"
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
            Index           =   1
            Left            =   45
            TabIndex        =   2
            Top             =   90
            Value           =   -1  'True
            Width           =   1230
         End
      End
      Begin VB.Frame FramData 
         Caption         =   "DATA VIEW"
         Height          =   8835
         Left            =   45
         TabIndex        =   36
         Top             =   585
         Width           =   10905
         Begin MSFlexGridLib.MSFlexGrid msflxData 
            Height          =   8190
            Left            =   45
            TabIndex        =   37
            Top             =   270
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   14446
            _Version        =   393216
            AllowUserResizing=   3
         End
         Begin VB.Label lblRecords 
            Height          =   240
            Left            =   135
            TabIndex        =   38
            Top             =   8505
            Width           =   3300
         End
      End
      Begin VB.Frame FramDesign2 
         Caption         =   "DESIGN VIEW"
         Height          =   8835
         Left            =   45
         TabIndex        =   49
         Top             =   585
         Width           =   10905
         Begin VB.CommandButton Command1 
            Caption         =   "Close"
            Height          =   465
            Left            =   9135
            TabIndex        =   32
            Top             =   8235
            Width           =   1635
         End
         Begin VB.CommandButton cmdCancelSql 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   4095
            TabIndex        =   30
            Top             =   360
            Width           =   1320
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   2790
            TabIndex        =   29
            Top             =   360
            Width           =   1320
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   1485
            TabIndex        =   28
            Top             =   360
            Width           =   1320
         End
         Begin VB.CommandButton cmdExecute 
            Caption         =   "Execute"
            Height          =   375
            Left            =   180
            TabIndex        =   27
            Top             =   360
            Width           =   1320
         End
         Begin VB.TextBox txtSQL 
            Appearance      =   0  'Flat
            Height          =   7260
            Left            =   180
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   855
            Width           =   10545
         End
         Begin VB.CommandButton cmdPrintSQL 
            Caption         =   "Print Structure"
            Height          =   465
            Left            =   7515
            TabIndex        =   31
            Top             =   8235
            Width           =   1635
         End
      End
      Begin VB.Frame FramDesign1 
         Caption         =   "DESIGN VIEW"
         Height          =   8835
         Left            =   45
         TabIndex        =   39
         Top             =   585
         Width           =   10905
         Begin VB.Frame FramTAbleFields 
            Caption         =   "TABLE FILEDS"
            Height          =   7845
            Left            =   135
            TabIndex        =   50
            Top             =   360
            Width           =   3885
            Begin VB.CommandButton cmdModify 
               Caption         =   "Modiffy"
               Height          =   375
               Left            =   2520
               TabIndex        =   8
               Top             =   7290
               Width           =   1185
            End
            Begin VB.CommandButton cmdRemove 
               Caption         =   "Remove"
               Height          =   375
               Left            =   1350
               TabIndex        =   7
               Top             =   7290
               Width           =   1185
            End
            Begin VB.ListBox lstFields 
               Height          =   5910
               Left            =   135
               TabIndex        =   5
               Top             =   1260
               Width           =   3615
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "Add "
               Height          =   375
               Left            =   180
               TabIndex        =   6
               Top             =   7290
               Width           =   1185
            End
            Begin VB.TextBox txtTabName 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   135
               TabIndex        =   4
               Top             =   450
               Width           =   3615
            End
            Begin VB.Label Label3 
               Caption         =   "Table Name"
               Height          =   240
               Left            =   135
               TabIndex        =   52
               Top             =   225
               Width           =   1365
            End
            Begin VB.Label Label4 
               Caption         =   "Field List"
               Height          =   240
               Left            =   135
               TabIndex        =   51
               Top             =   990
               Width           =   1365
            End
         End
         Begin VB.CommandButton cmdCloseTBDesign 
            Caption         =   "Close"
            Height          =   465
            Left            =   9135
            TabIndex        =   25
            Top             =   8235
            Width           =   1635
         End
         Begin VB.CommandButton cmdPrintTable 
            Caption         =   "Print Structure"
            Height          =   465
            Left            =   7515
            TabIndex        =   24
            Top             =   8235
            Width           =   1635
         End
         Begin VB.Frame FramFldDetails 
            Caption         =   "FIELD DETAILS"
            Enabled         =   0   'False
            Height          =   7845
            Left            =   4095
            TabIndex        =   40
            Top             =   360
            Width           =   6675
            Begin VB.CommandButton cmdCancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   3330
               TabIndex        =   23
               Top             =   6705
               Width           =   1230
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Height          =   375
               Left            =   2115
               TabIndex        =   22
               Top             =   6705
               Width           =   1230
            End
            Begin VB.TextBox txtDefaultVal 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1890
               TabIndex        =   16
               Top             =   5670
               Width           =   4335
            End
            Begin VB.TextBox txtValidationRule 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1890
               TabIndex        =   15
               Top             =   5085
               Width           =   4335
            End
            Begin VB.TextBox txtValidationTxt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1890
               TabIndex        =   14
               Top             =   4545
               Width           =   4335
            End
            Begin VB.TextBox txtOrdinal 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1890
               TabIndex        =   13
               Top             =   3960
               Width           =   1905
            End
            Begin VB.TextBox txtCollatingOrder 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   330
               Left            =   1890
               TabIndex        =   12
               Top             =   2790
               Width           =   1905
            End
            Begin VB.TextBox txtFldSize 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   330
               Left            =   1890
               TabIndex        =   11
               Top             =   2205
               Width           =   1905
            End
            Begin VB.CheckBox chkReq 
               Caption         =   "  Required"
               Enabled         =   0   'False
               Height          =   240
               Left            =   4185
               TabIndex        =   21
               Top             =   4005
               Width           =   1860
            End
            Begin VB.CheckBox chkZeroLen 
               Caption         =   "  AllowZeroLength"
               Enabled         =   0   'False
               Height          =   240
               Left            =   4185
               TabIndex        =   20
               Top             =   3420
               Width           =   1860
            End
            Begin VB.CheckBox chkAutoInc 
               Caption         =   "  AutoIncrement"
               Enabled         =   0   'False
               Height          =   240
               Left            =   4185
               TabIndex        =   19
               Top             =   2835
               Width           =   1860
            End
            Begin VB.CheckBox chkVarLen 
               Caption         =   "  VariableLength"
               Enabled         =   0   'False
               Height          =   240
               Left            =   4185
               TabIndex        =   18
               Top             =   2250
               Width           =   1860
            End
            Begin VB.CheckBox chkFixedLen 
               Caption         =   "  Fixed Length"
               Enabled         =   0   'False
               Height          =   240
               Left            =   4185
               TabIndex        =   17
               Top             =   1665
               Width           =   1860
            End
            Begin VB.ComboBox cboFldType 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1890
               TabIndex        =   10
               Text            =   "Combo1"
               Top             =   1635
               Width           =   1905
            End
            Begin VB.TextBox txtFldName 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1890
               TabIndex        =   9
               Top             =   1035
               Width           =   4335
            End
            Begin VB.Label Label10 
               Caption         =   "DefaultValue :"
               Height          =   240
               Left            =   450
               TabIndex        =   48
               Top             =   5715
               Width           =   1230
            End
            Begin VB.Label Label9 
               Caption         =   "ValidationRule :"
               Height          =   240
               Left            =   450
               TabIndex        =   47
               Top             =   5130
               Width           =   1230
            End
            Begin VB.Label Label8 
               Caption         =   "ValidationText :"
               Height          =   240
               Left            =   450
               TabIndex        =   46
               Top             =   4590
               Width           =   1230
            End
            Begin VB.Label Label7 
               Caption         =   "OrdinalPosition :"
               Height          =   240
               Left            =   450
               TabIndex        =   45
               Top             =   4005
               Width           =   1230
            End
            Begin VB.Label Label6 
               Caption         =   "CollatingOrder :"
               Height          =   240
               Left            =   450
               TabIndex        =   44
               Top             =   2835
               Width           =   1095
            End
            Begin VB.Label Label5 
               Caption         =   "Size :"
               Height          =   240
               Left            =   450
               TabIndex        =   43
               Top             =   2250
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "Type :"
               Height          =   240
               Left            =   450
               TabIndex        =   42
               Top             =   1665
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Name :"
               Height          =   240
               Left            =   450
               TabIndex        =   41
               Top             =   1080
               Width           =   1095
            End
         End
      End
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "Open Database"
   End
   Begin VB.Menu mnuCompare 
      Caption         =   "Compare Database"
   End
   Begin VB.Menu mnuChangePassword 
      Caption         =   "Change Password"
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Begin VB.Menu subTBStruct 
         Caption         =   "Table Structure"
      End
      Begin VB.Menu SubQueries 
         Caption         =   "Queries"
      End
   End
   Begin VB.Menu SubMenu1 
      Caption         =   "TableSub"
      Visible         =   0   'False
      Begin VB.Menu subOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu subrefresh 
         Caption         =   "Refresh List"
      End
      Begin VB.Menu subDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectedObject As String
Dim DesignNewTable As Boolean
Dim DesignNewQuery As Boolean
Dim EditOldTable As Boolean
Dim EditOldQuery As Boolean
Dim FldTypevalue(11) As Integer
Dim AddFieldFlag, editFieldFlag As Boolean
Dim ATOZ, ZTOA As Boolean


'**************************************************************************
'                       APPLICATION START
'**************************************************************************

Private Sub Form_Load()
    Call ToolBarButtonState(False)
    DesignNewTable = False
    DesignNewQuery = False
    EditOldTable = False
    EditOldQuery = False
    ATOZ = True
    ZTOA = False

    Call LoadFldTypes
End Sub

'**************************************************************************
'                       LOAD DATABASE FILED TYPES
'**************************************************************************

Private Sub LoadFldTypes()
cboFldType.Clear
    cboFldType.AddItem "Boolean"
    cboFldType.ItemData(0) = 1
    FldTypevalue(0) = dbBoolean
    
    cboFldType.AddItem "Byte"
    cboFldType.ItemData(1) = 1
    FldTypevalue(1) = dbByte
    
    cboFldType.AddItem "Integer"
    cboFldType.ItemData(2) = 2
    FldTypevalue(2) = dbInteger
    
    cboFldType.AddItem "Long"
    cboFldType.ItemData(3) = 4
    FldTypevalue(3) = dbLong
    
    cboFldType.AddItem "Currency"
    cboFldType.ItemData(4) = 8
    FldTypevalue(4) = dbCurrency
    
    cboFldType.AddItem "Single"
    cboFldType.ItemData(5) = 4
    FldTypevalue(5) = dbSingle
    
    cboFldType.AddItem "Double"
    cboFldType.ItemData(6) = 8
    FldTypevalue(6) = dbDouble
    
    cboFldType.AddItem "Date/Time"
    cboFldType.ItemData(7) = 8
    FldTypevalue(7) = dbDate
    
    cboFldType.AddItem "Text"
    cboFldType.ItemData(8) = 50
    FldTypevalue(8) = dbText
    
    cboFldType.AddItem "Long Binary"
    cboFldType.ItemData(9) = 0
    FldTypevalue(9) = dbLongBinary
    
    cboFldType.AddItem "Memo"
    cboFldType.ItemData(10) = 0
    FldTypevalue(10) = dbMemo
    
    cboFldType.ListIndex = 8
End Sub

Private Sub Form_Resize()
'    ResizeForm Me
End Sub

Private Sub lstTables_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu SubMenu1
    End If
End Sub
Private Sub subDelete_Click()
Dim str() As String
    str = Split(lstTables.Text, "-> ")
    If str(0) = "T" Then
        If MsgBox("Delete Table ? ", vbYesNo) = vbYes Then
            obDB.TableDefs.Delete (Trim(str(1)))
            MsgBox "Table Deleted.", vbInformation
        End If
        
    ElseIf str(0) = "Q" Then
        If MsgBox("Delete Query ? ", vbYesNo) = vbYes Then
            obDB.QueryDefs.Delete (Trim(str(1)))
            MsgBox "Query Deleted.", vbInformation
        End If
    End If
    GetTablesAndQueries
End Sub
'**************************************************************************
'                       MENU BUTTON EVENTS
'**************************************************************************
Private Sub mnuChangePassword_Click()
    frmChangePassword.Show vbModal
End Sub

Private Sub mnuCompare_Click()
    frmCompareDatabases.Show vbModal
End Sub
Private Sub mnuOpen_Click()
    frmOpenDatabase.Show vbModal
    If ConnectionState Then
        Call GetTablesAndQueries
        Me.Caption = strDBPath
        Call ToolBarButtonState(True)
    End If
End Sub

'**************************************************************************
'                       TOOLBAR BUTTON EVENTS
'**************************************************************************

Public Sub ToolBarButtonState(btStatus As Boolean)
Dim i As Integer
    For i = 2 To 6
        Toolbar1.Buttons(i).Enabled = btStatus
    Next i
End Sub








Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Index
        Case 1:
                frmOpenDatabase.Show vbModal
                If ConnectionState Then
                    Call GetTablesAndQueries
                    Me.Caption = strDBPath
                End If
        Case 2:
                Call GetTablesAndQueries
        Case 3:
                TransferMode = TranferTypes.Import
                frmImportExport.Caption = "IMPORT OBJECTS FROM....."
                frmImportExport.Show vbModal
        Case 4:
                TransferMode = TranferTypes.Export
                frmImportExport.Caption = "EXPORT OBJECTS TO....."
                frmImportExport.Show vbModal
        Case 5:
                frmChangePassword.Show vbModal
        Case 6:
                frmCompareDatabases.Show vbModal
        Case 7:
                frmRetrievePass.Show vbModal
        Case 8:
                DesignNewTable = True
                SelectedObject = "Table"
                optView(0).Value = True
                optView_Click 0
                Call AddTable
        Case 9:
                DesignNewQuery = True
                SelectedObject = "Query"
                txtSQL = ""
                optView(0).Value = True
                optView_Click 0
            
    End Select
End Sub
Private Sub AddTable()
Dim TabName As String
Dim tbDef As DAO.TableDef
     TabName = InputBox("Enter Table Name.", "Add New Table")
     If Trim(TabName) <> "" Then
        txtTabName = Trim(TabName)
        lstFields.Clear
        cmdAdd_Click
     End If
End Sub

'**************************************************************************
'                       GET DATABASE OBJECTS AND DISPLAY
'**************************************************************************

Private Function GetTablesAndQueries()
Dim i As Integer
    Dim tableObj1 As DAO.TableDef
    Dim QueryObj1 As DAO.QueryDef
    i = 0
    DoEvents
    lblTables.Caption = "Wait..........."
    DoEvents
    lstTables.Clear
    msflxData.Clear
     
    For Each tableObj1 In obDB.TableDefs
          lstTables.AddItem "T-> " & tableObj1.Name, i
          i = i + 1
    Next tableObj1
    For Each QueryObj1 In obDB.QueryDefs
          lstTables.AddItem "Q-> " & QueryObj1.Name, i
          i = i + 1
    Next QueryObj1
    If lstTables.ListCount > 0 Then lstTables.ListIndex = 0
    lblTables.Caption = "Tables = " & obDB.TableDefs.Count & ", Queries = " & obDB.QueryDefs.Count
    Set QueryObj1 = Nothing
    Set tableObj1 = Nothing

End Function
'**************************************************************************
'                       FOR SORTING DATA IN GRID
'**************************************************************************

Private Sub msflxData_Click()
    With msflxData
        If .Row = 1 Then
            If ATOZ Then
                .Sort = flexSortGenericAscending
                ATOZ = False
                ZTOA = False
            Else
                .Sort = flexSortGenericDescending
                ZTOA = False
                ATOZ = True
            End If
         End If
    End With
End Sub


'**************************************************************************
'                       SHOW DETAILS OF SELECTED DATABASE OBJECT
'**************************************************************************
Private Sub lstTables_Click()
Dim str() As String
    lstTables.Enabled = False
    str = Split(lstTables.Text, "-> ")
    
        If str(0) = "T" Then
            SelectedObject = "Table"
            Call ShowTableData(str(1))
            Call ShowTableDesign(str(1))
        Else
            SelectedObject = "Query"
            Call ShowQueryData(str(1))
            txtSQL = obDB.QueryDefs(Trim(str(1))).SQL
        End If
        
    Call optView_Click(1)
    lstTables.Enabled = True
    
End Sub
'**************************************************************************
'                       SHOW DATA OF SELECTED TABLE
'**************************************************************************
Public Sub ShowTableData(str As String)
Dim rstemp As DAO.Recordset
On Error Resume Next
    Set rstemp = obDB.TableDefs(str).OpenRecordset
    DoEvents
    lblRecords.Caption = "Wait..........."
    DoEvents
    Call ShowDataInGrid(msflxData, rstemp)
    
    DoEvents
    lblRecords.Caption = rstemp.RecordCount & " records found."
End Sub
'**************************************************************************
'                       SHOW DESIGN OF SELECTED TABLE
'**************************************************************************
Public Sub ShowTableDesign(str As String)
Dim fld As DAO.Field
    txtTabName = Trim(str)
    lstFields.Clear
    For Each fld In obDB.TableDefs(str).Fields
        lstFields.AddItem fld.Name
    Next
    If lstFields.ListCount > 0 Then lstFields.ListIndex = 0
End Sub
'**************************************************************************
'                       SHOW DETAILS OF SELECTED FIELD
'**************************************************************************
Private Sub lstFields_Click()
Dim i, idx As Integer
    Call ClearFldDetails
    With obDB.TableDefs(Trim(txtTabName)).Fields(Trim(lstFields.Text))
        txtFldName = .Name
        txtFldSize = .Size
        txtCollatingOrder = .CollatingOrder
        txtOrdinal = .OrdinalPosition
        txtValidationTxt = .ValidationText
        txtValidationRule = .ValidationRule
        txtDefaultVal = .DefaultValue
        If .AllowZeroLength Then chkZeroLen.Value = vbChecked
        If .Required Then chkReq.Value = vbChecked
        If .Attributes = dbAutoIncrField Then chkAutoInc.Value = vbChecked
        If .Attributes = dbVariableField Then chkVarLen.Value = vbChecked
        If .Attributes = dbFixedField Then chkFixedLen.Value = vbChecked
        idx = -1
        For i = 0 To cboFldType.ListCount - 1
            If Val(FldTypevalue(i)) = .Type Then
                idx = i
            End If
        Next i
        cboFldType.ListIndex = idx
        
    End With
    
End Sub
Private Sub ClearFldDetails()
    txtFldName = ""
    txtFldSize = ""
    txtOrdinal = ""
    txtCollatingOrder = ""
    txtValidationTxt = ""
    txtValidationRule = ""
    txtDefaultVal = ""
    chkFixedLen.Value = vbUnchecked
    chkVarLen.Value = vbUnchecked
    chkAutoInc.Value = vbUnchecked
    chkReq.Value = vbUnchecked
    chkZeroLen.Value = vbUnchecked
End Sub

'**************************************************************************
'                       ADD NEW FIELD IN TABLE
'**************************************************************************
Private Sub cmdAdd_Click()
    AddFieldFlag = True
    editFieldFlag = False
    Call DisableToolbar(True)
    Call ClearFldDetails
    If DesignNewTable = False Then txtOrdinal = obDB.TableDefs(txtTabName).Fields.Count
    txtFldName.SetFocus
    cboFldType_Click
    cboFldType.ListIndex = 8
End Sub
Private Sub DisableToolbar(cmd As Boolean)
    Me.Toolbar1.Enabled = Not cmd
    FramTables.Enabled = Not cmd
    FramTAbleFields.Enabled = Not cmd
    FramFldDetails.Enabled = cmd
    FramOpt.Enabled = Not cmd
End Sub
Private Sub cmdRemove_Click()
    If MsgBox("Remove field ?", vbYesNo, "REMOVE FIELD.") = vbYes Then
        obDB.TableDefs(Trim(txtTabName)).Fields.Delete (Trim(lstFields.Text))
        MsgBox "Field Removed.", vbInformation
        cmdCancel_Click
    End If
End Sub

Private Sub cmdModify_Click()
    AddFieldFlag = False
    editFieldFlag = True
    Call DisableToolbar(True)
    txtFldName.SetFocus
End Sub
Private Sub cboFldType_Click()
    If cboFldType.ListIndex > -1 Then
        Call DisableFields
        Select Case FldTypevalue(cboFldType.ListIndex)
            Case dbBoolean, dbByte, dbInteger, dbCurrency, dbSingle, dbDouble, dbDate, dbBinary:
                chkReq.Enabled = True
            Case dbLong:
                chkReq.Enabled = True
                chkAutoInc.Enabled = True
            Case dbText:
                chkFixedLen.Enabled = True
                chkVarLen.Enabled = True
                chkZeroLen.Enabled = True
                chkReq.Enabled = True
                txtFldSize.Enabled = True
            Case dbMemo:
                chkZeroLen.Enabled = True
                chkReq.Enabled = True
            End Select
            txtFldSize = cboFldType.ItemData(cboFldType.ListIndex)
    End If
End Sub
Private Sub DisableFields()
    chkAutoInc.Enabled = False
    chkFixedLen.Enabled = False
    chkVarLen.Enabled = False
    chkZeroLen.Enabled = False
    chkReq.Enabled = False
    txtFldSize.Enabled = False
End Sub
Private Sub chkFixedLen_Click()
    If chkFixedLen.Value = vbChecked And chkVarLen.Value = vbChecked Then
        chkVarLen.Value = vbUnchecked
    End If
End Sub
Private Sub chkVarLen_Click()
    If chkVarLen.Value = vbChecked And chkFixedLen.Value = vbChecked Then
        chkFixedLen.Value = vbUnchecked
    End If
End Sub
Private Sub cmdUpdate_Click()
On Error GoTo Errmsg
Dim fld As DAO.Field
Dim tbDef As DAO.TableDef
    If ValidField Then
        If DesignNewTable Then
            Set tbDef = obDB.CreateTableDef(Trim(txtTabName))
            Set fld = tbDef.CreateField()
            fld.Name = Trim(txtFldName)
            fld.Size = Trim(txtFldSize)
            fld.Type = FldTypevalue(cboFldType.ListIndex)
            fld.OrdinalPosition = Val(txtOrdinal)
            fld.ValidationText = Trim(txtValidationTxt)
            fld.ValidationRule = Trim(txtValidationRule)
            fld.DefaultValue = Trim(txtDefaultVal)
            If chkFixedLen.Value = vbChecked Then fld.Attributes = dbFixedField
            If chkVarLen.Value = vbChecked Then fld.Attributes = dbVariableField
            If chkAutoInc.Value = vbChecked Then fld.Attributes = dbAutoIncrField
            If chkZeroLen.Value = vbChecked Then fld.AllowZeroLength = True
            If chkReq.Value = vbChecked Then fld.Required = True
            tbDef.Fields.Append fld
            obDB.TableDefs.Append tbDef
        Else
             If editFieldFlag Then
                obDB.TableDefs(Trim(txtTabName)).Fields.Delete (Trim(lstFields.Text))
            End If
            'obDB.TableDefs(Trim(txtTabName)).Fields.Refresh
            Set fld = obDB.TableDefs(Trim(txtTabName)).CreateField()
            fld.Name = Trim(txtFldName)
            fld.Size = Trim(txtFldSize)
            fld.Type = FldTypevalue(cboFldType.ListIndex)
            fld.OrdinalPosition = Val(txtOrdinal)
            fld.ValidationText = Trim(txtValidationTxt)
            fld.ValidationRule = Trim(txtValidationRule)
            fld.DefaultValue = Trim(txtDefaultVal)
            If chkFixedLen.Value = vbChecked Then fld.Attributes = dbFixedField
            If chkVarLen.Value = vbChecked Then fld.Attributes = dbVariableField
            If chkAutoInc.Value = vbChecked Then fld.Attributes = dbAutoIncrField
            If chkZeroLen.Value = vbChecked Then fld.AllowZeroLength = True
            If chkReq.Value = vbChecked Then fld.Required = True
            obDB.TableDefs(Trim(txtTabName)).Fields.Append fld
            MsgBox "Field Added."
        End If
        cmdCancel_Click
    End If
Exit Sub
Errmsg:
    MsgBox "Error No. " & Err.Number & " - " & Err.Description
End Sub
Private Function ValidField() As Boolean
    If Trim(txtFldName) = "" Then
        MsgBox "Please enter a value for Field name."
        txtFldName.SetFocus
        ValidField = False
    Else
        ValidField = True
    End If

End Function
Private Sub cmdCancel_Click()
    Call DisableToolbar(False)
    AddFieldFlag = False
    editFieldFlag = False
    If DesignNewTable Then
        lstTables.AddItem "T-> " & Trim(txtTabName), lstTables.ListCount
        lstTables.Text = "T-> " & Trim(txtTabName)
        
    End If
    DesignNewTable = False
    lstTables_Click
End Sub


'**************************************************************************
'                       SHOW DATA OF SELECTED QUERY
'**************************************************************************
Public Sub ShowQueryData(str As String)
On Error GoTo Errmsg
Dim rstemp As DAO.Recordset
    Set rstemp = obDB.QueryDefs(str).OpenRecordset
    DoEvents
    lblRecords.Caption = "Wait..........."
    DoEvents
    Call ShowDataInGrid(msflxData, rstemp)
    
    DoEvents
    lblRecords.Caption = rstemp.RecordCount & " records found."
Exit Sub
Errmsg:
    MsgBox "Error No: " & Err.Number & " - " & Err.Description
End Sub

'**************************************************************************
'                       SET DATA IN GRID
'**************************************************************************

Sub ShowDataInGrid(msgrid As MSFlexGrid, rstemp As DAO.Recordset)
Dim i, j As Integer
    With msgrid
            .Rows = 1
            .Cols = 1
            .TextMatrix(0, 0) = "Sr. No."
            For i = 1 To rstemp.Fields.Count
                If .Cols <= i Then .Cols = .Cols + 1
                .TextMatrix(0, i) = rstemp.Fields(i - 1).Name
            Next i
            j = 1
            Do While rstemp.EOF = False
                If .Rows <= j Then .Rows = .Rows + 1
                .TextMatrix(j, 0) = j
                 For i = 1 To rstemp.Fields.Count
                    .TextMatrix(j, i) = rstemp.Fields(i - 1).Value & ""
                Next i
                j = j + 1
                rstemp.MoveNext
            Loop
        End With
End Sub
'**************************************************************************
'SELECT OPTION FOR DISPLAYING OBJECT DETAILS IN DATA VIEW OR DESGIN VIEW
'**************************************************************************
Private Sub optView_Click(Index As Integer)
    If optView(1).Value Then
        Call ShowFrame(FramData)
    Else
        If SelectedObject = "Table" Then
            Call ShowFrame(FramDesign1)
        Else
            Call ShowFrame(FramDesign2)
        End If
    End If
End Sub
'**************************************************************************
'                       DISPLAY/HIDE THE FRAME
'**************************************************************************
Public Sub ShowFrame(Fram As Frame)
    FramData.Visible = False
    FramDesign1.Visible = False
    FramDesign2.Visible = False
    Fram.Visible = True
End Sub

'**************************************************************************
'                       EXECUTE THE SQL
'**************************************************************************
Private Sub cmdExecute_Click()
On Error GoTo Errmsg
Dim rstemp As DAO.Recordset
    
    Set rstemp = obDB.OpenRecordset(Trim(txtSQL), dbOpenDynaset, dbReadOnly)
    DoEvents
    lblRecords.Caption = "Wait..........."
    DoEvents
    Call ShowDataInGrid(msflxData, rstemp)
    DoEvents
    lblRecords.Caption = rstemp.RecordCount & " records found."
    optView(1).Value = True
    Call optView_Click(1)
Exit Sub
Errmsg:
    MsgBox Err.Description
End Sub
'**************************************************************************
'                       CLEAR THE SQL EDITOR WINDOW
'**************************************************************************
Private Sub cmdClear_Click()
    txtSQL.Text = ""
End Sub

'**************************************************************************
'                       SAVE THE QUERY AFTER CHANGES
'**************************************************************************
Private Sub cmdSave_Click()
Dim str As String
Dim STR1() As String
Dim tdfNew As DAO.QueryDef
        If DesignNewQuery = False Then
            If MsgBox("Overwrite Existing Query ?", vbYesNo, "Save Query") = vbYes Then
                STR1 = Split(lstTables.Text, "-> ")
                obDB.QueryDefs(Trim(STR1(1))).SQL = Trim(txtSQL)
                Exit Sub
            End If
        End If
        str = InputBox("Enter new query name.", "Save Query")
        Set tdfNew = obDB.CreateQueryDef(str)
        tdfNew.SQL = Trim(txtSQL)
        'obDB.QueryDefs.Append tdfNew
        If MsgBox("Execute Query ? ", vbYesNo, "Run Query") = vbYes Then
            cmdExecute_Click
        End If
        lstTables.AddItem "Q-> " & Trim(str)
        lstTables.Text = "Q-> " & Trim(str)
        DesignNewQuery = False
End Sub


'**************************************************************************
'                      UNDO THE CHANGES IN SQL
'**************************************************************************
Private Sub cmdCancelSql_Click()
    DesignNewQuery = False
    lstTables_Click
End Sub
Private Sub cmdCloseTBDesign_Click()
    optView(1).Value = True
    optView_Click (1)
End Sub

'**************************************************************************
'                      PRINT TABLE STRUCTURE OF A SELECTED TABLE
'**************************************************************************
Private Sub cmdPrintTable_Click()
Dim fs, a
Dim fld As DAO.Field
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(App.Path & "\Tables.htm") Then
        Kill (App.Path & "\Tables.htm")
    End If
    Set a = fs.CreateTextFile(App.Path & "\Tables.htm", True)
    a.WriteLine "<html>"
    a.WriteLine "<Head> <H4><u>" & obDB.Name & "</u></H4></head>"
    a.WriteLine "<body>"
    a.WriteLine "<br><b><u>" & lstTables.Text & "</u></b>"
'Printing the fields of table
    a.WriteLine "<Table>"
    
    a.WriteLine "<tr>"
    a.WriteLine "<td width=100><u> Field Name </u></td>"
    a.WriteLine "<td width=150><u> Type </u></td>"
    a.WriteLine "<td width=100 align=right><u> Size </u></td>"
    a.WriteLine "</tr><br>"
    For Each fld In obDB.TableDefs(Trim(txtTabName)).Fields
        a.WriteLine "<tr>"
        a.WriteLine "<td width=150>" & fld.Name & "</td>"
        a.WriteLine "<td width=100>" & GetFldTypeName(fld.Type) & "</td>"
        a.WriteLine "<td  align=right>" & fld.Size & "</td>"
        a.WriteLine "</tr>"
    Next
    a.WriteLine "</table>"
    a.WriteLine "</body> </HTml>"
    a.Close
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & App.Path & "\Tables.htm", vbMaximizedFocus
End Sub
'**************************************************************************
'                      PRINT SQL STATEMENT
'**************************************************************************
Private Sub cmdPrintSQL_Click()
Dim fs, a
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(App.Path & "\Queries.htm") Then
        Kill (App.Path & "\Queries.htm")
    End If
    Set a = fs.CreateTextFile(App.Path & "\Queries.htm", True)
    a.WriteLine "<html>"
    a.WriteLine "<Head> <h4><u>" & obDB.Name & "</u></h4></head>"
    a.WriteLine "<body>"
    a.WriteLine "<br><b><u>" & lstTables.Text & "</u></b>"
    a.WriteLine "<br><br>" & txtSQL
    a.WriteLine "</body> </HTml>"
    a.Close
    Shell "C:\Program Files\Internet Explorer\iexplore.exe " & App.Path & "\Queries.htm", vbMaximizedFocus
End Sub
'**************************************************************************
'                      PRINT SELECTED TABLES STRUCTURE
'**************************************************************************
Private Sub subTBStruct_Click()
    frmPrintTables.Show vbModal
End Sub

'**************************************************************************
'                      PRINT SELECTED QUERIES
'**************************************************************************
Private Sub SubQueries_Click()
    frmPrintQueries.Show vbModal
End Sub

'**************************************************************************
'                       APPLICATION CLOSE
'**************************************************************************
Private Sub mnuExit_Click()
    End
End Sub
'**************************************************************************
'                       KEYPRESS EVENTS
'**************************************************************************
Private Sub txtSQL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdExecute_Click
    End If
End Sub
