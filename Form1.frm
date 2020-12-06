VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:KiLLer:."
   ClientHeight    =   5775
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Active windows"
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         ToolTipText     =   "Cliquez ici pour raffraichir es fenêtres actives"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ListBox List 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         ItemData        =   "Form1.frx":038A
         Left            =   120
         List            =   "Form1.frx":038C
         TabIndex        =   6
         ToolTipText     =   "Double-cliquez sur une fenêtre pour l'ajouter"
         Top             =   720
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forbidden words"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":038E
         Left            =   120
         List            =   "Form1.frx":0390
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Please express words whose titles should not contain"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2760
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   2880
      Picture         =   "Form1.frx":0392
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu Menu 
      Caption         =   "Files"
      Begin VB.Menu Restor 
         Caption         =   "Restore"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu menuAproposDe 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu option 
      Caption         =   "Option"
      Begin VB.Menu check 
         Caption         =   "Just close Browser windows"
      End
      Begin VB.Menu barre0 
         Caption         =   "-"
      End
      Begin VB.Menu demarrage 
         Caption         =   "Add to startup"
         Checked         =   -1  'True
      End
      Begin VB.Menu u 
         Caption         =   "-"
      End
      Begin VB.Menu actpasse 
         Caption         =   "Activate password"
         Checked         =   -1  'True
      End
      Begin VB.Menu chgpasse 
         Caption         =   "Change password"
      End
      Begin VB.Menu barre 
         Caption         =   "-"
      End
      Begin VB.Menu rappor 
         Caption         =   "Check Report"
      End
   End
End
Attribute VB_Name = "Form1"
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

Dim ActivPasse As Boolean
Private Type IconeTray
    cbSize As Long      'Icon size (in bytes)
    hwnd As Long        'Handle of the window responsible for receiving the messages sent during events on the icon (clicks, double-clicks, etc.)
    uID As Long         'Icon identifier
    uFlags As Long
    uCallbackMessage As Long    'Messages to resend
    hIcon As Long               'Icon handle
    szTip As String * 64        'Text to put in the tooltip
End Type
Dim IconeT As IconeTray
'Constants required
Private Const AJOUT = &H0
Private Const MODIF = &H1
Private Const SUPPRIME = &H2
Private Const MOUSEMOVE = &H200
Private Const MESSAGE = &H1
Private Const Icone = &H2
Private Const TIP = &H4
Private Const DOUBLE_CLICK_GAUCHE = &H203
Private Const BOUTON_GAUCHE_POUSSE = &H201
Private Const BOUTON_GAUCHE_LEVE = &H202
Private Const DOUBLE_CLICK_DROIT = &H206
Private Const BOUTON_DROIT_POUSSE = &H204
Private Const BOUTON_DROIT_LEVE = &H205
'API required
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As IconeTray) As Boolean


Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Dim winWnd
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Dim i







Private Sub actpasse_Click()

If actpasse.Checked = True Then
actpasse.Checked = False
ActivPasse = False

Set WshShell = CreateObject("Wscript.Shell")
WshShell.RegWrite "HKEY_CURRENT_USER\Software\killer\acp", "0"

Else
actpasse.Checked = True
ActivPasse = True
Set WshShell = CreateObject("Wscript.Shell")

WshShell.RegWrite "HKEY_CURRENT_USER\Software\killer\acp", "1"

End If
End Sub

Private Sub Check2_Click()

End Sub

Private Sub check_Click()
If check.Checked = True Then check.Checked = False
If check.Checked = False Then check.Checked = True
End Sub

Private Sub chgpasse_Click()
frmLoginch.Show
Me.Hide
End Sub






Private Sub Command1_Click()
lister
End Sub

Private Sub Command5_Click()
Dim a
Dim b
a = InputBox("What is the word to add", "Word to add")
If a <> "" Then
Combo1.AddItem a
' registration in the registry
del
Set WshShell = CreateObject("Wscript.Shell")
  For b = 0 To Combo1.ListCount - 1
WshShell.RegWrite "HKEY_CURRENT_USER\Software\killer\" & b, Combo1.List(b)
  Next b
Else
MsgBox "The box was empty, please start over", vbCritical, "Information"
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
abcdef = Combo1.ListIndex
Combo1.RemoveItem (abcdef)
' registration in the registry
Dim b
Dim a
del
Set WshShell = CreateObject("Wscript.Shell")
  For b = 0 To Combo1.ListCount - 1
WshShell.RegWrite "HKEY_CURRENT_USER\Software\killer\" & b, Combo1.List(b)
  Next b
End Sub

Private Sub demarrage_Click()
On Error Resume Next
If demarrage.Checked = True Then
'mettre
demarrage.Checked = False
Set WshShell = CreateObject("Wscript.Shell")
WshShell.Regdelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Killer"
Else
'enlever
demarrage.Checked = True
Set WshShell = CreateObject("Wscript.Shell")
WshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Killer", App.Path & "\" & App.EXEName & ".exe"
End If
End Sub

Private Sub Form_Load()
liredem
lirepass
Dim a
Form1.Hide
On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
For a = 0 To 5000
Combo1.AddItem WshShell.RegRead("HKEY_CURRENT_USER\Software\killer\" & a)
Next a

'Preparation of the IconeT variable
IconeT.cbSize = Len(IconeT) 'Icon size in bytes
IconeT.hwnd = Me.hwnd       'Handle of the application (so that it receives the messages sent during a click, double-click ...
IconeT.uID = 1&             'Icon identifier
IconeT.uFlags = Icone Or TIP Or MESSAGE
IconeT.uCallbackMessage = MOUSEMOVE     'Resend messages about mouse action
IconeT.hIcon = Image1.Picture   'Put in icon the image which is in the control "Image1"
IconeT.szTip = "Close unwanted windows" & Chr$(0)    'Tooltip text
'Call the function to put the icon in the system tray
Shell_NotifyIcon AJOUT, IconeT


App.TaskVisible = False     'Remove the application button from task bar
                            
'Menu.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static rec As Boolean, msg As Long

'Occurs when the user acts with the mouse on
'the icon placed in the system tray

msg = X / Screen.TwipsPerPixelX
If rec = False Then
    rec = True
    Select Case msg     '' Different possibilities of action
        Case DOUBLE_CLICK_GAUCHE:   'Put
            Restor_Click     'Here
        Case BOUTON_GAUCHE_POUSSE:  'What
        Case BOUTON_GAUCHE_LEVE:    'You
        Case DOUBLE_CLICK_DROIT:    'Need
        Case BOUTON_DROIT_POUSSE:   'To
        Case BOUTON_DROIT_LEVE:     'Happen
            PopupMenu Menu, , , , menuAproposDe     'brings up the menu
           '"About" will appear in bold
    End Select
    rec = False
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Call the API again to remove the icon from the system tray
'when the program exits, this time using the constant DELETE
'instead of ADD

IconeT.cbSize = Len(IconeT)
IconeT.hwnd = Me.hwnd
IconeT.uID = 1&
'Shell_NotifyIcon hidde, IconeT
Shell_NotifyIcon SUPPRIME, IconeT
End
End Sub






Private Sub Form_Resize()
If Me.WindowState = 1 Then Form1.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)

End
End Sub



Private Sub List_DblClick()
Dim b
Combo1.AddItem List.Text
MsgBox ("Word " & List.Text & " Successfully added")
'enregistrement dans la base de registre
del
Set WshShell = CreateObject("Wscript.Shell")
  For b = 0 To Combo1.ListCount - 1
WshShell.RegWrite "HKEY_CURRENT_USER\Software\killer\" & b, Combo1.List(b)
  Next b
End Sub

Private Sub menuAproposDe_Click()
MsgBox "Codded by: MrAlone -!" & vbCrLf & "If you like it that's great" & vbCrLf & "Copyright 2021", vbInformation, "Hello !"
End Sub

Private Sub rappor_Click()
rapport.Show
End Sub

Private Sub Timer1_Timer()
Dim hwnd As Long
Dim Titre_Fenetre As String * 255
Dim TitreFen As String
Dim j As Long
Dim i
hwnd = GetWindow(GetDesktopWindow(), 5)
Do While (Not IsNull(hwnd)) And (hwnd <> 0) 'Go through each window
    Titre_Fenetre = String(255, 0)  'Format the string intended to host the title of the window
    
    ret = GetWindowText(hwnd, Titre_Fenetre, 255)   'get the window title and the number of characters in this title
    If Titre_Fenetre <> String(255, 0) Then             'If the title is not empty
        If IsWindowVisible(hwnd) = 1 Then                   'To take into account only visible windows (see what happens by removing this condition)
            TitreFen = Titre_Fenetre        'get the window title
            TitreFen = Left(TitreFen, ret)  'without the final additional characters
            j = j + 1
            If Val(j) < 10 Then j = "0" & j
       For i = 0 To Combo1.ListCount - 1
       
            If TitreFen <> "" Then
                If InStr(1, TitreFen, Combo1.List(i), vbTextCompare) Then
winWnd = FindWindow(vbNullString, TitreFen)
If check.Checked = True Then
'TmpRep = SendMessage(CLng(winWnd), CLng("16"), CInt("0"), "0")

TmpRep = SendMessage(CLng(winWnd), CLng("130"), CInt("0"), "0")
TmpRep = SendMessage(CLng(winWnd), CLng("45"), CInt("0"), "0")
TmpRep = SendMessage(CLng(winWnd), CLng("545"), CInt("0"), "0")
TmpRep = SendMessage(CLng(winWnd), CLng("18"), CInt("0"), "0")

ElseIf check.Checked = False Then
TmpRep = SendMessage(CLng(winWnd), CLng("2"), CInt("0"), "0")
TmpRep = SendMessage(CLng(winWnd), CLng("130"), CInt("0"), "0")
TmpRep = SendMessage(CLng(winWnd), CLng("45"), CInt("0"), "0")
TmpRep = SendMessage(CLng(winWnd), CLng("545"), CInt("0"), "0")
TmpRep = SendMessage(CLng(winWnd), CLng("18"), CInt("0"), "0")

If Dir("rpp.dat") = vbNullString Then
Open "rpp.dat" For Output As #1
Print #1, "The day " & Date & " at " & Time & "Window closed : [" & TitreFen & "]" & vbCrLf & "Because it contained :[" & Combo1.List(i) & "]"
Close #1
Else
Open "rpp.dat" For Append As #1
Print #1, "The day " & Date & " at " & Time & vbCrLf & "      Window closed : [" & TitreFen & "]" & vbCrLf & "           Because it contained :[" & Combo1.List(i) & "]"
Close #1
End If
            End If
                End If
End If
        Next i


        End If
    End If
    hwnd = GetWindow(hwnd, 2) 'look for the next window
Loop

End Sub

Private Sub Restor_Click()
If ActivPasse = True Then
frmLogin.Show
Else
Me.WindowState = 0
Me.Show
lister
End If
End Sub

Sub lirepass()
On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
a = WshShell.RegRead("HKEY_CURRENT_USER\Software\killer\acp")
If a = "1" Then
actpasse.Checked = True
ActivPasse = True
Else
actpasse.Checked = False
ActivPasse = False
End If
End Sub

Sub del()
On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
For a = 0 To 5000
 WshShell.Regdelete ("HKEY_CURRENT_USER\Software\killer\" & a)
Next a

End Sub
Sub liredem()
On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
a = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Killer")
If a <> "" Then
demarrage.Checked = True
Else
demarrage.Checked = False
End If
End Sub

Sub lister()
List.Clear
Dim hWnd1 As Long
Dim Titre_Fenetre1 As String * 255
Dim TitreFen1 As String
Dim j1 As Long
List.Clear                       'We empty the listbox
hWnd1 = GetWindow(GetDesktopWindow(), 5)
Do While (Not IsNull(hWnd1)) And (hWnd1 <> 0) 'Go through each window
    Titre_Fenetre1 = String(255, 0)  'Format the string intended to host the title of the window
    
    ret = GetWindowText(hWnd1, Titre_Fenetre1, 255)   'get the window title and the number of characters in this title
    If Titre_Fenetre1 <> String(255, 0) Then             'If the title is not empty
        If IsWindowVisible(hWnd1) = 1 Then                  'To take into account only visible windows (see what happens by removing this condition)
            TitreFen1 = Titre_Fenetre1        'get the window title
            TitreFen1 = Left(TitreFen1, ret)  'without the final additional characters
            j = j + 1
            If Val(j) < 10 Then j = "0" & j
            List.AddItem TitreFen1  'store everything in a listbox
        End If
    End If
    hWnd1 = GetWindow(hWnd1, 2) 'look for the next window
Loop

End Sub
