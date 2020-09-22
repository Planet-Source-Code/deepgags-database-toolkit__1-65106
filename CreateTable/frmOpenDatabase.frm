VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOpenDatabase 
   Caption         =   "OPEN DATABASE"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5910
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   2970
         TabIndex        =   7
         Top             =   2295
         Width           =   1680
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
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
         Caption         =   "PASSWORD"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   900
         TabIndex        =   5
         Top             =   1455
         Width           =   2130
      End
   End
End
Attribute VB_Name = "frmOpenDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
On Error GoTo Errmsg
    cdgFile.Filter = "MS Access (*.mdb)|*.mdb|"
    cdgFile.ShowOpen
    txtFile1.Text = cdgFile.FileName
Exit Sub
Errmsg:
End Sub

Private Sub cmdOpen_Click()
On Error GoTo Errhandler
    strDBPath = Trim(txtFile1.Text)
    DBPass = Trim(txtPass1.Text)
    Call OpenConnection
    Call frmMain.ToolBarButtonState(True)
    Unload Me
Exit Sub
Errhandler:
    MsgBox Err.Description
    txtPass1.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
