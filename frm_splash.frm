VERSION 5.00
Begin VB.Form frm_splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4140
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   8220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frm_splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSleep 
      Left            =   7440
      Top             =   2760
   End
   Begin VB.Image imgSplash 
      Appearance      =   0  'Flat
      Height          =   4125
      Left            =   0
      Top             =   0
      Width           =   8220
   End
End
Attribute VB_Name = "frm_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    imgSplash.Picture = LoadPicture(Dir_Icon & "splash_logo.jpg")
    tmrSleep.Interval = Splash_Time
    
End Sub

Private Sub tmrSleep_Timer()

    Unload Me

    If Auth_Require = True Then
        frm_auth_user.Show
    Else
        frm_main_mdi.Show
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frm_splash = Nothing
    
End Sub
