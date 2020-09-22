VERSION 5.00
Begin VB.Form Popup 
   BorderStyle     =   0  'None
   ClientHeight    =   2265
   ClientLeft      =   1290
   ClientTop       =   1065
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   181
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Sleeper 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1035
      Top             =   1800
   End
   Begin VB.Timer MouseOver 
      Interval        =   50
      Left            =   540
      Top             =   1800
   End
   Begin VB.PictureBox ButtonPictureSource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   2385
      Picture         =   "Popup.frx":0000
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   7
      Top             =   1935
      Width           =   225
   End
   Begin VB.PictureBox ButtonPictureSource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   2115
      Picture         =   "Popup.frx":024A
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   6
      Top             =   1935
      Width           =   225
   End
   Begin VB.PictureBox ButtonPictureSource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1845
      Picture         =   "Popup.frx":0494
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   5
      Top             =   1935
      Width           =   225
   End
   Begin VB.Timer Loader 
      Interval        =   20
      Left            =   45
      Top             =   1800
   End
   Begin VB.PictureBox BackPicture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   0
      Picture         =   "Popup.frx":06DE
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      Begin VB.PictureBox PopupPicture 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   135
         Picture         =   "Popup.frx":FDA0
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   50
         TabIndex        =   4
         Top             =   630
         Width           =   750
      End
      Begin VB.PictureBox ButtonPicture 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2430
         Picture         =   "Popup.frx":11B92
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   1
         Top             =   90
         Width           =   195
      End
      Begin VB.Label PopupInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MsN Popup Clone is coded by Ramci"
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   1080
         TabIndex        =   3
         Top             =   675
         Width           =   1230
         WordWrap        =   -1  'True
      End
      Begin VB.Label PopupCaption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MsN Style Popup"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   2265
      End
   End
End
Attribute VB_Name = "Popup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Const SND_FILENAME = &H20000
Private Const SND_ASYNC = &H1

Private InputArray$()
Private PopupLoaded As Boolean
Private CurrentWindow&, MasterWindow&, CurrentPOINT As POINTAPI
Private RoundRectRgn&
Private MouseOverButton As Boolean, MouseDownButton As Boolean
Private Shell_TrayWnd_hWnd&, Shell_TrayWnd_RECT As RECT, Shell_TrayWnd_Height&
Private Screen_Width&, Screen_Height&

Private Sub ButtonPicture_Click()

    'End if close button clicked
    End

End Sub

Private Sub BackPicture_KeyPress(KeyAscii As Integer)

    'end if ESC is pressed
    If KeyAscii = 27 Then End

End Sub

Private Sub ButtonPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Set the mouse down picture
    MouseDownButton = True
    ButtonPicture.Picture = ButtonPictureSource(2)

End Sub

Private Sub Form_Initialize()

    Height = 10
    'Find the SystemTray
    Shell_TrayWnd_hWnd = FindWindow("Shell_TrayWnd", vbNullString)
    Call GetWindowRect(Shell_TrayWnd_hWnd, Shell_TrayWnd_RECT)
    'Calculate height of SystemTray
    Shell_TrayWnd_Height = Shell_TrayWnd_RECT.Bottom - Shell_TrayWnd_RECT.Top
    'Get Screen Size
    Screen_Width = Screen.Width / Screen.TwipsPerPixelX
    Screen_Height = Screen.Height / Screen.TwipsPerPixelY
    'Move the popup to start up place
    Call MoveWindow(hWnd, Screen_Width - ScaleWidth - 18, Screen_Height - Shell_TrayWnd_Height - ScaleHeight, ScaleWidth, ScaleHeight, True)
    'Round the rect of the PopupPicture picturebox
    RoundRectRgn = CreateRoundRectRgn(0, 0, PopupPicture.ScaleWidth, PopupPicture.ScaleHeight, (PopupPicture.ScaleWidth / 3), (PopupPicture.ScaleHeight / 3))
    Call SetWindowRgn(PopupPicture.hWnd, RoundRectRgn, True)
    Call DeleteObject(RoundRectRgn)

End Sub

Private Sub Form_Load()

    'You may run this popup from your project
    'Put the last MsNPopup sub in your project (at the bottom of the page)
    If Command$ <> "" Then
        InputArray = Split(Command, "|")
        ReDim Preserve InputArray(2)
        If InputArray(0) <> "" Then
            If Dir(InputArray(0)) <> "" Then PopupPicture.Picture = LoadPicture(InputArray(0))
        End If
        If InputArray(1) <> "" Then PopupCaption.Caption = InputArray(1)
        If InputArray(2) <> "" Then PopupInfo.Caption = InputArray(2)
    End If
    'Play the sound of online warning
    Call PlaySound(App.Path + "\online.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC)

End Sub

Private Sub Loader_Timer()

    'Popup Is Loaded? [it is rised to top in this case it is loaded]
    If PopupLoaded = False Then
        'Rise the popup
        If ScaleHeight < BackPicture.ScaleHeight Then
            Call MoveWindow(hWnd, Screen_Width - ScaleWidth - 18, Screen_Height - Shell_TrayWnd_Height - ScaleHeight - 3, ScaleWidth, ScaleHeight + 3, True)
        Else
            Call MoveWindow(hWnd, Screen_Width - ScaleWidth - 18, Screen_Height - Shell_TrayWnd_Height - BackPicture.ScaleHeight, ScaleWidth, BackPicture.ScaleHeight, True)
            Loader.Enabled = False
            PopupLoaded = True
            Sleeper.Enabled = True
        End If
    Else
        'fall the popup
        If ScaleHeight > 3 Then
            Call MoveWindow(hWnd, Screen_Width - ScaleWidth - 18, Screen_Height - Shell_TrayWnd_Height - ScaleHeight + 3, ScaleWidth, ScaleHeight - 3, True)
        Else
            Call MoveWindow(hWnd, Screen_Width - ScaleWidth - 18, Screen_Height - Shell_TrayWnd_Height - 3, ScaleWidth, 3, True)
            End
        End If
    End If

End Sub

Private Sub MouseOver_Timer()

    'Mouse Still Down?
    If MouseDownButton = True And GetAsyncKeyState(vbKeyLButton) = 0 Then MouseDownButton = False
    'Mouse Over Close Button?
    Call GetCursorPos(CurrentPOINT)
    CurrentWindow = WindowFromPoint(CurrentPOINT.X, CurrentPOINT.Y)
    'Mouse came over the close button
    If MouseOverButton = False And CurrentWindow = ButtonPicture.hWnd Then
        MouseOverButton = True
        If MouseDownButton = False Then
            'Set mouse over picture if mouse is over me and not down
            ButtonPicture.Picture = ButtonPictureSource(1)
        Else
            'Set mouse down picture cuz left button of mouse is not left
            ButtonPicture.Picture = ButtonPictureSource(2)
        End If
    End If
    'Mouse left the close button
    If MouseOverButton = True And CurrentWindow <> ButtonPicture.hWnd Then
        MouseOverButton = False
        'Set normal picture to button
        ButtonPicture.Picture = ButtonPictureSource(0)
    End If
    'Get the master handle of the window
    'which mouse is over
    While CurrentWindow > 0
        MasterWindow = CurrentWindow
        CurrentWindow = GetParent(CurrentWindow)
    Wend
    'Mouse Over Popup?
    'Is handle of the master is the handle of my popup ?
    If PopupLoaded = True Then
        If (MasterWindow = hWnd) And (Sleeper.Enabled = True) Then
            Sleeper.Enabled = False
        ElseIf (MasterWindow <> hWnd) And (Sleeper.Enabled = False) Then
            Sleeper.Enabled = True
        End If
    End If
 
End Sub

Private Sub Sleeper_Timer()

    'Wait for 5 secs before fall
    Loader.Enabled = True
    Sleeper.Enabled = False

End Sub

'This is the sub that you can use in your project
Private Sub MsNPopup(PopupPicturePath$, PopupCaption$, PopupInfo$)

    'You must have "msnpopupclone.exe" in the same directory of your project
    If Dir(App.Path + "\msnpopupclone.exe") = "" Then Exit Sub
    Call Shell(App.Path + "\msnpopupclone.exe " + PopupPicturePath + "|" + PopupCaption + "|" + PopupInfo, vbNormalFocus)

End Sub
