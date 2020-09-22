VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRetrievePass 
   Caption         =   "RETRIEVE PASSWORD"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox pass2000 
      Height          =   330
      Left            =   3060
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox Pass1997 
      Height          =   330
      Left            =   765
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5910
      Begin VB.CommandButton cmdTry 
         Caption         =   "Try Again"
         Height          =   375
         Left            =   405
         TabIndex        =   7
         Top             =   810
         Width           =   915
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   405
         TabIndex        =   6
         Top             =   1215
         Width           =   915
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "...."
         Height          =   315
         Left            =   5400
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   1575
         TabIndex        =   0
         Top             =   735
         Width           =   3735
      End
      Begin VB.TextBox txtPass1 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1350
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
         Left            =   1575
         TabIndex        =   5
         Top             =   495
         Width           =   2805
      End
      Begin VB.Label Label5 
         Caption         =   "PASSWORD "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1575
         TabIndex        =   4
         Top             =   1125
         Width           =   2130
      End
   End
End
Attribute VB_Name = "frmRetrievePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
On Error GoTo Errhandler
    txtPass1.Text = ""
    cdgFile.Filter = "MS Access (*.mdb)|*.mdb|"
    cdgFile.ShowOpen
    txtFileName.Text = cdgFile.FileName
    Call GetPassWord
    
Exit Sub
Errhandler:
End Sub
Private Sub GetPassWord()
    Pass1997.Text = Get1997Pass()
    pass2000.Text = Get2000Pass()
    txtPass1.Text = Trim(Pass1997)
End Sub
Private Sub cmdClose_Click()
    End
End Sub

Private Function Get2000Pass() As String
Dim DBPassword As String
    'On Error GoTo errHandler
    Dim ch(40) As Byte
    Dim x As Integer, sec2, intChar As Integer, blnUse3 As Boolean
    If Trim(txtFileName) = "" Then Exit Function
    'Used integers instead of hex :-)  Easier to read
    sec2 = Array(0, 194, 117, 236, 55, 25, 202, 156, 250, 130, 208, 40, 230, 87, 56, 138, 96, 16, 26, 123, 54, 177, 252, 223, 177, 51, 122, 19, 67, 139, 33, 177, 51, 112, 239, 121, 91, 214, 59, 124, 42)
    'I found that some DB's use this scheme see below for the logic to determine which is which :-)
    sec3 = Array(0, 229, 117, 236, 55, 62, 202, 156, 250, 165, 208, 40, 230, 112, 56, 138, 96, 55, 26, 123, 54, 150, 252, 223, 177, 20, 122, 19, 67, 172, 33, 177, 51, 87, 239, 121, 91, 241, 59, 124, 42)
    
    blnUse3 = False
    
    Open txtFileName.Text For Binary Access Read As #1 Len = 40
    Get #1, &H42, ch
    Close #1
    'Check to see which key by running through first 6 letters of password
    'This is not foolproof by any means.
    For x = 1 To 6
      intChar = ch(x) Xor sec2(x)
         'This is kind of lame but it assumes that most passwords
         'are in this range of keyboard chars :-)
      If ((intChar < 32) Or (intChar > 126)) And (intChar <> 0) Then
         blnUse3 = True 'Set a flag
      End If
    Next x
    For x = 1 To 40
         If blnUse3 = True Then
            intChar = ch(x) Xor sec3(x)
         Else
            intChar = ch(x) Xor sec2(x)
         End If
         DBPassword = DBPassword & Chr(intChar)
    Next x
     Get2000Pass = DBPassword
    Exit Function
Errhandler:
    MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
    Exit Function
End Function

Private Function Get1997Pass() As String
'On Error GoTo errHandler
Dim DBPassword As String
    Dim ch(18) As Byte, x As Integer
    Dim sec
    If Trim(txtFileName) = "" Then Exit Function
    'Used integers instead of hex :-)  Easier to read
    sec = Array(0, 134, 251, 236, 55, 93, 68, 156, 250, 198, 94, 40, 230, 19, 182, 138, 96, 84)
    
    Open txtFileName.Text For Binary Access Read As #1 Len = 18
    Get #1, &H42, ch
    Close #1
    For x = 1 To 17
        DBPassword = DBPassword & Chr(ch(x) Xor sec(x))
    Next x
    Get1997Pass = DBPassword
    Exit Function
Errhandler:
    MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
    Exit Function
    
End Function

Private Sub cmdTry_Click()
    
    If Trim(txtPass1.Text) = CStr(Pass1997) Then
        txtPass1.Text = CStr(pass2000)
    Else
        txtPass1.Text = CStr(Pass1997)
    End If
End Sub
