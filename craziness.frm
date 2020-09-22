VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All code* in_________      _____________     _________      '
'this prog  / _______/|    / ___________/|   / _______/|     '
'is copy-  / /_______|/   / /___________|/  / /______ |/     '
'right by / /_______     / /  _________    / /_______        '
'Bryant  / ________/|   / /  /______  /|  / ________/|       '
'Zadegan/ /________|/  / /   |____ / / / / /________|/       '
'      / /________    / /_________/ / / / /________          '
'     /__________/|  /_____________/ / /__________/|         '
'     |__________|/ |______________|/  |__________|/         '
'                     EGan Electronics                       '
'*all code except the transparency coding. that code belongs to Jan Alexander
' he is at "http://www.jan-alexander.de"

'EGE stands for Egan Electronics. this name cannot be used elsewere! OK!?! OK!!!
'               EG   E
'   This "toy" exploits a big mistake in windows. Not a bug, but a mistake!
'   they probably made it like this to keep cpus from suffering reactor-like
'   temps, but it does leave for some interesting ideas!

'   The "mistake" is at the bottom of the prog.
'   It will be markd by three (3) single quotes.
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Dim X As Integer
'all things above this line are used for transparencies
Dim up As Boolean 'if this is true, then transparency increases. Else, it decreases!

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Text1.Left = Screen.Width - Text1.Width 'aligns text box with the right of screen
X = 0
up = True 'this is used to determine if the transparency is increasing
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'the above line prevents other windows from coming to the top
Form1.Width = Screen.Width
Form1.Height = Screen.Height
'if u add a single quote above and remove one where it sez "me.height=400" then
'this prog will only override the start menu
Form1.Top = Screen.Height - Me.Height
Form1.Left = Screen.Width - Me.Width
MakeTransparent Form1.hWnd, X 'enables transparency
Me.BackColor = vbRed 'ups the fear factor :)
'me.Height = 400
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'if someone presses enter on this text box then this happens:
    If Text1.Text = "BUZZ OFF!" Then End 'this is how the thing ends. rather evil
    'if the victim doesnt know what to type (mwahaha)
End If
End Sub

'Private Sub Form_Click()
'End
'End Sub
'the above sub is good if u want to make this a painless prank :(

Private Sub Timer1_Timer()
'    If Form1.BackColor = vbWhite Then
'        Form1.BackColor = vbBlack
'    Else
'        Form1.BackColor = vbWhite
'    End If
'if u want ur form to flash black and white instead then remove the single quotes above
If X > 150 Then
    up = False
Else
    If X < 1 Then up = True
End If
If up = True Then
    X = X + 5
Else: X = X - 5
End If
MakeTransparent Form1.hWnd, X 'refreshes transparency
Text1.BackColor = Me.BackColor 'makes textbox merge with form
'''Heres the bug! down below!
If X > 0 Or Y > 0 Or cx > 0 Or cy > 0 Then Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'upper line refreshes the "topmost" line. this overrides the windows task manager!
'whats rather unusual is that the normal alt + F4 combo wont kill the prog...
'u think its the top line? btw alt + tab works, but if u want to bring thw windows
'task manager up, the tm will still be overriden.

End Sub
