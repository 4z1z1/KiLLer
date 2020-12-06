VERSION 5.00
Begin VB.Form rapport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "rapport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Effacer"
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "rapport"
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
                                               
                                               


Private Sub Command1_Click()
If MsgBox("Are you sure you want to delete the report?", vbYesNo, "Deletion of report") = vbYes Then
    Open "rpp.dat" For Output As 1
    Close 1
    Text1.Text = ""
End If
End Sub

Private Sub Form_Load()

    Dim count As Integer
    Dim intval As String
    Open "rpp.dat" For Input As 1
    While Not EOF(1)
        Input #1, intval
        Text1 = Text1 & vbCrLf & intval
    Wend

    Close 1

End Sub
