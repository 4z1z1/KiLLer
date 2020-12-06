VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Killer Password"
   ClientHeight    =   1455
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   859.662
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   405
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter the password"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
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

Dim MotDePasse2
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
'Assigns the value False to the global variable
'to indicate connection failure.
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
'Check if the password is correct.
    If Crypt(txtPassword.Text) = MotDePasse2 Then
'Place the code here to report
'to the calling procedure the success of the function.

        txtPassword = ""
        LoginSucceeded = True
        Form1.WindowState = 0
        Form1.lister
        Form1.Show
        Me.Hide
    Else
        MsgBox "Invalid password, try again !", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub


Private Sub Form_Load()
On Error GoTo Erreur
Set WshShell = CreateObject("Wscript.Shell")
MotDePasse2 = WshShell.RegRead("HKEY_CURRENT_USER\Software\killer\pass")


Exit Sub
Erreur:
MotDePasse2 = Crypt("0000")
End Sub
