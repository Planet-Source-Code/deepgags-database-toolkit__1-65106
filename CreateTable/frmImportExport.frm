VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportExport 
   Caption         =   "IMPORT / EXPORT OBJECTS FROM....."
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5910
      Begin VB.Frame FramTables 
         Height          =   2085
         Left            =   585
         TabIndex        =   8
         Top             =   585
         Visible         =   0   'False
         Width           =   4605
         Begin VB.CommandButton cmdBack 
            Caption         =   "Back"
            Height          =   330
            Left            =   2160
            TabIndex        =   12
            Top             =   1350
            Width           =   1230
         End
         Begin VB.CommandButton cmdTransfer 
            Caption         =   "Import / Export"
            Height          =   330
            Left            =   945
            TabIndex        =   11
            Top             =   1350
            Width           =   1230
         End
         Begin VB.ComboBox cboObjects 
            Height          =   315
            Left            =   360
            TabIndex        =   9
            Top             =   765
            Width           =   3525
         End
         Begin VB.Label lblmsg 
            Caption         =   "Select Object to Import"
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
            Left            =   360
            TabIndex        =   10
            Top             =   450
            Width           =   2400
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   2970
         TabIndex        =   7
         Top             =   2295
         Width           =   1680
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show Details"
         Height          =   330
         Left            =   990
         TabIndex        =   3
         Top             =   2295
         Width           =   1680
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "...."
         Height          =   315
         Left            =   4740
         TabIndex        =   1
         Top             =   870
         Width           =   375
      End
      Begin VB.TextBox txtFile1 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   870
         Width           =   3735
      End
      Begin VB.TextBox txtPass1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   900
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1725
         Width           =   3735
      End
      Begin MSComDlg.CommonDialog cdgFile 
         Left            =   135
         Top             =   135
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "DATABASE PATH"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   900
         TabIndex        =   6
         Top             =   630
         Width           =   2805
      End
      Begin VB.Label Label5 
         Caption         =   "PASSWORD "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   900
         TabIndex        =   5
         Top             =   1455
         Width           =   2130
      End
   End
End
Attribute VB_Name = "frmImportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obDaoTemp As DAO.Workspace
Dim obDBtemp As DAO.Database

'********************************************************************************
'GET THE DATABASE FILE PATH IN WHICH TO IMPORT / EXPORT
'********************************************************************************
Private Sub cmdBrowse_Click()
On Error GoTo Errmsg
    cdgFile.Filter = "MS Access (*.mdb)|*.mdb|"
    cdgFile.ShowOpen
    txtFile1.Text = cdgFile.FileName
Exit Sub
Errmsg:
    
End Sub

'********************************************************************************
'OPEN THE DATABASE TO WHICH TO IMPORT / EXPORT
'********************************************************************************
Private Sub cmdShow_Click()
    If txtFile1.Text = "" Then
        MsgBox "No File is selected.", vbInformation
        txtFile1.SetFocus
        Exit Sub
    Else
        On Error GoTo Errmsg
        Set obDaoTemp = DAO.DBEngine.Workspaces(0)
        Set obDBtemp = obDaoTemp.OpenDatabase(Trim(txtFile1.Text), False, False, ";pwd=" & Trim(txtPass1) & "")
    End If
    If TransferMode = TranferTypes.Import Then
        lblmsg.Caption = "Select an object to Import"
        cmdTransfer.Caption = "Import"
        Call LoadObjects(obDBtemp)
    ElseIf TransferMode = TranferTypes.Export Then
        lblmsg.Caption = "Select an object to Export"
        cmdTransfer.Caption = "Export"
        Call LoadObjects(obDB)
    End If
 
    FramTables.Visible = True
Exit Sub
Errmsg:
    MsgBox "Invalid Password.Try Again", vbInformation
    txtPass1.SetFocus
End Sub

'********************************************************************************
'LOAD THE OBJECTS OF DATABASE FROM WHICH THE OBJECTS HAS TO BE EXPORTED.
'********************************************************************************

Private Sub LoadObjects(tmpObj As DAO.Database)
Dim tmpTable As DAO.TableDef
Dim tmpQuery As DAO.QueryDef
    cboObjects.Clear
    For Each tmpTable In tmpObj.TableDefs
        cboObjects.AddItem "T->" & tmpTable.Name
    Next
    For Each tmpQuery In tmpObj.QueryDefs
        cboObjects.AddItem "Q->" & tmpQuery.Name
    Next
    If cboObjects.ListCount > 0 Then cboObjects.ListIndex = 0
End Sub
'********************************************************************************
'                      TRANSFER DATABASE OBJECTS
'********************************************************************************
Private Sub cmdTransfer_Click()
Dim str() As String
    str = Split(Trim(cboObjects.Text), ">")
    If str(0) = "T-" Then
        If TransferMode = TranferTypes.Import Then
              Call TransferTable(obDBtemp, str(1), obDB)
        Else
              Call TransferTable(obDB, str(1), obDBtemp)
        End If
    ElseIf str(0) = "Q-" Then
        If TransferMode = TranferTypes.Import Then
              Call TransferQuery(obDBtemp, str(1), obDB)
        Else
              Call TransferQuery(obDB, str(1), obDBtemp)
        End If
            
    End If
    If TransferMode = TranferTypes.Import Then
        MsgBox cboObjects.Text & " Imported Successfully", vbInformation
    Else
        MsgBox cboObjects.Text & " Exported Successfully", vbInformation
    End If

    FramTables.Visible = False
End Sub

'********************************************************************************
'GET BACK TO DATABASE SELECTION
'********************************************************************************
Private Sub cmdBack_Click()
    FramTables.Visible = False
End Sub

'********************************************************************************
'CLOSE THE FORM
'********************************************************************************
Private Sub cmdClose_Click()
    Unload Me
End Sub

