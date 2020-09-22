VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Popup"
      Height          =   825
      Left            =   810
      TabIndex        =   0
      Top             =   1035
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Call MsNPopup(App.Path + "\mypic.bmp", "Hey", "This popup is coded my Ramci")

End Sub

'This is the sub that you can use in your project
Private Sub MsNPopup(PopupPicturePath$, PopupCaption$, PopupInfo$)

    'You must have "msnpopupclone.exe" in the same directory of your project
    If Dir(App.Path + "\msnpopupclone.exe") = "" Then Exit Sub
    Call Shell(App.Path + "\msnpopupclone.exe " + PopupPicturePath + "|" + PopupCaption + "|" + PopupInfo, vbNormalFocus)

End Sub
