VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDatabaseStructure 
   Caption         =   "DISPLAY DATABASE STUCTURE"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7725
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6030
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   330
         Left            =   4815
         TabIndex        =   1
         Top             =   7245
         Width           =   915
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlxFile1 
         Height          =   6900
         Left            =   225
         TabIndex        =   2
         Top             =   270
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   12171
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
      End
      Begin MSComDlg.CommonDialog cdgFile 
         Left            =   11700
         Top             =   225
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmDatabaseStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SetGrid
    Call Showtables
End Sub
Public Sub SetGrid()
    With MSFlxFile1
    
        .TextMatrix(0, 0) = "Table "
        .TextMatrix(0, 1) = "Field "
        .TextMatrix(0, 2) = "Type"
        .TextMatrix(0, 3) = "Size"
        .ColWidth(0) = .ColWidth(0) * 1.5
        .ColWidth(1) = .ColWidth(1) * 1.5
       
     End With
End Sub
Private Function Showtables() As Boolean
    Dim tableObj1 As DAO.TableDef
    Dim FieldObj1 As DAO.Field
    i = 2
    
    For Each tableObj1 In obDB.TableDefs
        With MSFlxFile1
            .Rows = .Rows + 1
            .TextMatrix(1, flag) = "Total Tables " & obDB.TableDefs.Count
            .Rows = .Rows + 1
            .TextMatrix(i, flag) = tableObj1.Name
                For Each FieldObj1 In tableObj1.Fields
                    .TextMatrix(i, flag + 1) = FieldObj1.Name
                    .TextMatrix(i, flag + 2) = FieldObj1.Type
                    .TextMatrix(i, flag + 3) = FieldObj1.Size
                    .Rows = .Rows + 1
                    i = i + 1
                Next
            i = i + 1
        End With
    Next tableObj1
    Set tableObj1 = Nothing
   
End Function

