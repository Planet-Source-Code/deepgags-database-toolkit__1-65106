VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCompareDatabasesNew 
   Caption         =   "DISPLAY DATABASE STUCTURE"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7725
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   6255
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Height          =   330
         Left            =   4905
         TabIndex        =   8
         Top             =   1035
         Width           =   915
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   330
         Left            =   4815
         TabIndex        =   7
         Top             =   7245
         Width           =   915
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlxFile1 
         Height          =   5595
         Left            =   495
         TabIndex        =   6
         Top             =   1575
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   9869
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
      End
      Begin VB.TextBox txtPass1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "mdecodc"
         Top             =   1020
         Width           =   4230
      End
      Begin VB.TextBox txtFile1 
         Height          =   315
         Left            =   480
         TabIndex        =   2
         Text            =   "C:\Documents and Settings\Administrator\Desktop\CreateTable\vet1.mdb"
         Top             =   390
         Width           =   4905
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "...."
         Height          =   315
         Index           =   0
         Left            =   5445
         TabIndex        =   1
         Top             =   390
         Width           =   375
      End
      Begin MSComDlg.CommonDialog cdgFile 
         Left            =   11700
         Top             =   225
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "PASSWORD"
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
         Left            =   480
         TabIndex        =   5
         Top             =   750
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "DATABASE PATH"
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
         Left            =   480
         TabIndex        =   3
         Top             =   150
         Width           =   2940
      End
   End
End
Attribute VB_Name = "frmCompareDatabasesNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SetGrid
End Sub
Public Sub SetGrid()
    With MSFlxFile1
    
        .TextMatrix(0, 0) = "Table "
        .TextMatrix(0, 1) = "Field "
        .TextMatrix(0, 2) = "Type"
        .TextMatrix(0, 3) = "Size"
        
        .TextMatrix(0, 4) = "Table "
        .TextMatrix(0, 5) = "Field "
        .TextMatrix(0, 6) = "Type"
        .TextMatrix(0, 7) = "Size"
        
        .ColWidth(0) = .ColWidth(0) * 1.5
        .ColWidth(1) = .ColWidth(1) * 1.5
        .ColWidth(4) = .ColWidth(4) * 1.5
        .ColWidth(5) = .ColWidth(5) * 1.5
     End With
End Sub
Private Sub cmdBrowse_Click(index As Integer)
On Error GoTo Errhandler
    
    cdgFile.Filter = "MS Access (*.mdb)|*.mdb|"
    cdgFile.ShowOpen
    txtFile1.Text = cdgFile.FileName
    
    Exit Sub
    
Errhandler:
End Sub

Private Sub cmdShow_Click()
    Dim obDAO1 As DAO.Workspace, obDB1 As DAO.Database
    Dim PassString1 As String
    Dim TableObj As DAO.TableDef
    Set obDAO1 = DAO.DBEngine.Workspaces(0)
    
    PassString1 = ""
    
    Set obDB1 = obDAO1.OpenDatabase(txtFile1.Text, False, False, ";pwd=" & txtPass1 & "")
    
    
    Call Showtables(obDB1, 0)
    obDB1.Close
    obDAO1.Close

    Set obDB1 = Nothing
    Set obDAO1 = Nothing

    
End Sub

Private Function Showtables(DBObject1 As DAO.Database, flag As Integer) As Boolean
    Dim tableObj1 As DAO.TableDef
    Dim FieldObj1 As DAO.Field
    i = 2
    
    For Each tableObj1 In DBObject1.TableDefs
        With MSFlxFile1
            .Rows = .Rows + 1
            .TextMatrix(1, flag) = "Total Tables " & DBObject1.TableDefs.Count
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

