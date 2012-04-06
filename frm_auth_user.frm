VERSION 5.00
Begin VB.Form frm_auth_user 
   BackColor       =   &H00E9E9E9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnLogin 
      BackColor       =   &H00E5E5E5&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "+"
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E9E9E9&
      Caption         =   "Application Authorization"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton btnExit 
         BackColor       =   &H00E5E5E5&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "And the Password :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter your Username :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   4935
   End
End
Attribute VB_Name = "frm_auth_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strikes As Byte

Private Sub Form_Load()

    btnExit.Cancel = True
    btnLogin.Default = True
    
End Sub

Private Sub Form_Paint()
    txtUsername.SetFocus
End Sub

Private Sub btnExit_Click()
    End
End Sub

Private Sub btnLogin_Click()

    Dim user, pass As String
    Dim proceed As Boolean
    
    proceed = False
        
    user = txtUsername.text
    pass = txtPassword.text
    
    If user = "" Or pass = "" Then
        lblResult.Caption = "Please enter the user / pass fields completely."
        Exit Sub
    End If

    If user = Auth_User And String_Encode(pass) = Auth_Pass Then
        proceed = True
    End If

    If proceed = True Then

        lblResult.Caption = "Login Successful. Initializing the application..."

        Unload Me
        frm_main_mdi.Show

    Else
    
        strikes = strikes + 1
        lblResult.Caption = "Login Failed! Attempts Remaining : " & Str(3 - strikes)
        
        If strikes = 3 Then
            End
        End If
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_auth_user = Nothing
End Sub
