VERSION 5.00
Begin VB.Form frmChangePassword 
   Caption         =   "CHANGE DATABASE PASSWORD"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2850
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4920
      Begin VB.TextBox txtPass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   585
         Width           =   2745
      End
      Begin VB.TextBox txtPass2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1710
         Width           =   2745
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set"
         Height          =   330
         Left            =   1845
         TabIndex        =   3
         Top             =   2340
         Width           =   1230
      End
      Begin VB.TextBox txtPass1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1125
         Width           =   2745
      End
      Begin VB.Label Label2 
         Caption         =   "Old Password"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   270
         TabIndex        =   7
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Confirm Password"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   270
         TabIndex        =   6
         Top             =   1800
         Width           =   2130
      End
      Begin VB.Label Label5 
         Caption         =   "New Password"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   270
         TabIndex        =   5
         Top             =   1170
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSet_Click()
Dim strOpenPwd As String
    
    
    Call closeConnection
    strOpenPwd = ";pwd=" & Trim(txtPass)

   ' Open database for exclusive access by using current password. To get
   ' exclusive access, you must set the Options argument to True.
   Set obDB = OpenDatabase(Name:=strDBPath, Options:=True, ReadOnly:=False, Connect:=strOpenPwd)
    If Trim(txtPass) = DBPass Then
        If Trim(txtPass1) = Trim(txtPass2) Then
            obDB.NewPassword Trim(txtPass), Trim(txtPass1)
            DBPass = Trim(txtPass1)
            MsgBox "Password changed."
        Else
            MsgBox "Passwords not matched."
        End If
    Else
            MsgBox "Invalid old password."
    End If
    Unload Me
    Call closeConnection
    Call OpenConnection
      
End Sub
