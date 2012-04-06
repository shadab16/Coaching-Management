VERSION 5.00
Begin VB.Form frm_test 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Form"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5715
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frm_test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Me.Cls
    For i = 0 To 5
        Print "  i = "; i, "Result = "; Get_Form_by_Tag(i)
    Next i

End Sub

Private Sub Command2_Click()

    Me.Cls
    Dim str1, str2, str3 As String
    
    str1 = "qwerty"
    Print "String 1: "; str1
    
    str2 = String_Encode(str1)
    Print "String 2: "; str2
    
    str3 = String_Decode(str2)
    Print "String 3: "; str3

End Sub

Private Sub Command3_Click()
    frm_test2.Show
End Sub

Private Sub Command4_Click()
    frm_assoc_sc.Show
End Sub

Private Sub Command5_Click()

    Dim nTmp As Long
    Static VarX As Integer
    Set db_RS(10) = New ADODB.RecordSet
    
    VarX = VarX + 1
    
    db_SQL = "insert into test(f1, f2) values('testing','blah')"
    objDB.Execute db_SQL, nTmp, db_RS(10)

End Sub

Private Sub Command7_Click()

    Call frm_view_pic.ShowPic("D:\Pictures 2\Windows Vista.jpg", "Windows Vista Ultimate")

End Sub

Private Sub Command8_Click()

    objDB.Disconnect

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_test = Nothing
End Sub
