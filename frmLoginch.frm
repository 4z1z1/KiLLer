VERSION 5.00
Begin VB.Form frmLoginch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2115
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4410
   Icon            =   "frmLoginch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1249.612
   ScaleMode       =   0  'User
   ScaleWidth      =   4140.751
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1016
      Width           =   2282
   End
   Begin VB.TextBox txtPassword1 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   600
      Width           =   2282
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   180
      Width           =   2282
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   960
      TabIndex        =   3
      Top             =   1620
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2400
      TabIndex        =   4
      Top             =   1620
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "@Confirm the password"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Old Password"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1680
   End
   Begin VB.Label lblLabels 
      Caption         =   "&New Password"
      Height          =   270
      Index           =   1
      Left            =   107
      TabIndex        =   2
      Top             =   677
      Width           =   1800
   End
End
Attribute VB_Name = "frmLoginch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'___  ___      ___ _
'|  \/  |     / _ \| |
'| .  . |_ __/ /_\ \ | ___  _ __   ___   ______
'| |\/| | '__|  _  | |/ _ \| '_ \ / _ \ |______|
'| |  | | |  | | | | | (_) | | | |  __/
'\_|  |_/_|  \_| |_/_|\___/|_| |_|\___|
'
'
' _____           _ _
'|_   _|         (_)   (_)
'  | |_   _ _ __  _ ___ _  __ _
'  | | | | | '_ \| / __| |/ _` |
'  | | |_| | | | | \__ \ | (_| |
'  \_/\__,_|_| |_|_|___/_|\__,_|


Dim NewMotDePasse
Dim MotDePasse

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
If txtPassword1 <> txtPassword2 Then
MsgBox "The two passwords do not match", vbInformation, "information"
Exit Sub
End If
    If txtUserName = MotDePasse Then
    NewMotDePasse = Crypt(txtPassword1)
        Set WshShell = CreateObject("Wscript.Shell")
        WshShell.RegWrite "HKEY_CURRENT_USER\Software\killer\pass", NewMotDePasse
MsgBox "Your password has been changed", vbInformation, "Password Changed!"
txtPassword1 = ""
txtPassword2 = ""
txtUserName = ""
Me.Hide
Form1.Show
    Else
        MsgBox "Invalid password, try again!", , "Login"
        txtUserName.SetFocus

    End If
End Sub

Private Sub Form_Load()
On Error GoTo Erreur
Set WshShell = CreateObject("Wscript.Shell")
MotDePasse = WshShell.RegRead("HKEY_CURRENT_USER\Software\killer\pass")
MotDePasse = Crypt(MotDePasse)
Exit Sub
Erreur:
MotDePasse = "0000"
End Sub
