VERSION 5.00
Begin VB.Form frm_view_pic 
   BorderStyle     =   0  'None
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblPic 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2805
   End
   Begin VB.Image imgFull 
      Height          =   3255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frm_view_pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PicWidth, PicHeight, FrmWidth, FrmHeight As Integer

Public Sub ShowPic(ByVal PicPath As String, Optional ByVal Info As String = "")

    With Me
    
        .Hide
    
        lblPic.Caption = " " & Info
        imgFull.Picture = LoadPicture(PicPath)
    
        .Width = 0.9 * frm_main_mdi.ScaleWidth
        .Height = 0.9 * frm_main_mdi.ScaleHeight
    
        PicWidth = imgFull.Picture.Width
        PicHeight = imgFull.Picture.Height
    
        FrmWidth = .Width
        FrmHeight = .Height

        Call Resize(PicWidth, PicHeight, FrmWidth, FrmHeight)
    
        imgFull.Width = PicWidth
        imgFull.Height = PicHeight
    
        .Width = PicWidth
        .Height = PicHeight
    
        .Top = (frm_main_mdi.ScaleHeight - .Height) / 2
        .Left = (frm_main_mdi.ScaleWidth - .Width) / 2
        
        .BorderStyle = 1
        .Show
    
    End With

End Sub

Private Sub imgFull_Click()
    Unload Me
End Sub
