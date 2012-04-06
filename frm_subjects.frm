VERSION 5.00
Begin VB.Form frm_subjects 
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Subjects Taught"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   9855
   Begin VB.CommandButton btnFirst 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   625
   End
   Begin VB.CommandButton btnPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   625
   End
   Begin VB.CommandButton btnNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   4395
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   625
   End
   Begin VB.CommandButton btnLast 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   625
   End
   Begin VB.CommandButton btnRec_Add 
      BackColor       =   &H00F2F2F2&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Add a NEW Record"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton btnRec_Edit 
      BackColor       =   &H00F2F2F2&
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "EDIT the currently displayed record"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton btnRec_Save 
      BackColor       =   &H00F2F2F2&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Save all the changes done to the database"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton btnRec_Delete 
      BackColor       =   &H00F2F2F2&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton btnRec_Refresh 
      BackColor       =   &H00F2F2F2&
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Associated Courses"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   4920
      TabIndex        =   23
      Top             =   1680
      Width           =   4695
      Begin VB.ListBox lstCourses 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         ItemData        =   "frm_subjects.frx":0000
         Left            =   240
         List            =   "frm_subjects.frx":0002
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Subject Info."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   4455
      Begin VB.TextBox txtInfo 
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
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
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
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label50 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Subject ID :"
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
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Subject Name :"
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
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F2F2F2&
      Height          =   855
      Left            =   240
      TabIndex        =   18
      Top             =   600
      Width           =   9375
      Begin VB.ComboBox cmbSearch_Field 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_subjects.frx":0004
         Left            =   3600
         List            =   "frm_subjects.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Table Field to Search IN"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtSearch_Input 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         ToolTipText     =   "Text to search FOR"
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cmbSearch_Type 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_subjects.frx":0008
         Left            =   5640
         List            =   "frm_subjects.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Search Types"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btnSearch_Submit 
         BackColor       =   &H00F2F2F2&
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cmbSearch_Start 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frm_subjects.frx":000C
         Left            =   6840
         List            =   "frm_subjects.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Start searching FROM"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Search :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   600
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frm_subjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnEdit, blnAdd As Boolean
Dim sngTemp As Integer, strTemp As String

Private Sub Form_Load()

    btnFirst.Picture = LoadPicture(Dir_Icon & "first.gif")
    btnPrev.Picture = LoadPicture(Dir_Icon & "back.gif")
    btnNext.Picture = LoadPicture(Dir_Icon & "next.gif")
    btnLast.Picture = LoadPicture(Dir_Icon & "last.gif")

    Lock_Controls
    Search_Form_Fill

    Load_Subj_Rec

End Sub

Private Sub Load_Subj_Rec()

    Set db_RS(5) = New ADODB.RecordSet

    db_SQL = "SELECT SUBJECT_ID, SUBJECT_NAME " _
            & "FROM cm_subject ORDER BY subject_id"

    objDB.Query db_SQL, db_RS(5)

    ReLoad_Subj_Rec

End Sub

Private Sub ReLoad_Subj_Rec()

    txtInfo(0).text = db_RS(5).fields(0) & ""
    txtInfo(1).text = db_RS(5).fields(1) & ""
    
    List_Assoc_Courses

End Sub

Private Sub List_Assoc_Courses()

    lstCourses.Clear
    
    Set db_RS(8) = New ADODB.RecordSet
    
    db_SQL = "SELECT " _
            & "C.course_id, C.course_name, C.class, C.type, S.subject_id " _
            & "FROM cm_course C, cm_course_subject_teacher S " _
            & "WHERE C.course_id = S.course_id " _
            & "AND S.subject_id = " & db_RS(5).fields(0) _
            & " ORDER BY C.course_id"

    objDB.Query db_SQL, db_RS(8)
    
    With db_RS(8)
        If .RecordCount > 0 Then
        
            For i = 0 To (.RecordCount - 1)
    
                strTemp = !course_id & " - Class " & !Class & " - " _
                        & !course_name & " (" & !Type & ")"
                lstCourses.AddItem strTemp
                
                objDB.Move "next", db_RS(8)
                
            Next i
        End If
        
        strTemp = ""
        lblInfo.Caption = .RecordCount & " Courses are associated with this subject."
    End With

End Sub

Private Sub Lock_Controls()

    txtInfo(0).Locked = True
    txtInfo(1).Locked = True
    
    btnRec_Save.Enabled = False
    btnRec_Delete.Enabled = False

End Sub

Private Sub UnLock_Controls()

    'txtInfo(0).Locked = False
    txtInfo(1).Locked = False

    btnRec_Save.Enabled = True
    'btnRec_Delete.Enabled = True

End Sub

Private Sub Lock_Navigation()

    btnFirst.Enabled = False
    btnPrev.Enabled = False
    btnNext.Enabled = False
    btnLast.Enabled = False
    
    btnRec_Refresh.Enabled = False
    btnSearch_Submit.Enabled = False

End Sub

Private Sub UnLock_Navigation()

    btnFirst.Enabled = True
    btnPrev.Enabled = True
    btnNext.Enabled = True
    btnLast.Enabled = True
    
    btnRec_Refresh.Enabled = True
    btnSearch_Submit.Enabled = True

End Sub

Private Sub btnRec_Add_Click()

    If blnAdd <> True Then
    
        btnRec_Add.Caption = "CANCEL"
        btnRec_Edit.Enabled = False
        
        UnLock_Controls
        Lock_Navigation
        
        txtInfo(0).text = ""
        txtInfo(1).text = ""
        lstCourses.Clear
    
        blnAdd = True
    
    Else
    
        btnRec_Add.Caption = "ADD"
        btnRec_Edit.Enabled = True
        
        Lock_Controls
        UnLock_Navigation
        btnRec_Refresh_Click
    
        blnAdd = False
    
    End If

End Sub

Private Sub btnRec_Edit_Click()

    If blnEdit <> True Then
    
        btnRec_Edit.Caption = "CANCEL"
        btnRec_Add.Enabled = False
        
        UnLock_Controls
        Lock_Navigation
        
        blnEdit = True
    
    Else
    
        btnRec_Edit.Caption = "EDIT"
        btnRec_Add.Enabled = True
        
        Lock_Controls
        UnLock_Navigation
        btnRec_Refresh_Click
        
        blnEdit = False
    
    End If

End Sub

Private Sub btnRec_Save_Click()

    If Check_Form_Fields() = False Then
        Exit Sub
    End If
    
    If blnEdit = True Then

        Set db_RS(7) = New ADODB.RecordSet

        db_SQL = "UPDATE cm_subject SET " _
                & "SUBJECT_NAME = '" & txtInfo(1).text & "' " _
                & "WHERE SUBJECT_ID = " & txtInfo(0).text

        objDB.Execute db_SQL, sngTemp

        frm_test3.Show
        frm_test3.fill (db_SQL & " . And sngTemp is = " & sngTemp)

        If sngTemp > 0 Then
            lblResult.Caption = "Data edited / updated accordingly."
        Else
            lblResult.Caption = "Data could NOT be edited / updated."
        End If
        
    ElseIf blnAdd = True Then
    
        Set db_RS(7) = New ADODB.RecordSet
        
        db_SQL = "INSERT INTO cm_subject(SUBJECT_ID, SUBJECT_NAME) " _
                & "VALUES(seq_test.nextval, '" & txtInfo(1).text & "')"
    
        objDB.Execute db_SQL, sngTemp
        
        frm_test3.Show
        frm_test3.fill (db_SQL & " . And sngTemp is = " & sngTemp)
        
        If sngTemp > 0 Then
            lblResult.Caption = "New Data added accordingly."
        Else
            lblResult.Caption = "New Data could NOT be added."
        End If
                
        btnRec_Add_Click
        btnLast_Click
        btnRec_Edit_Click

    End If

End Sub

Private Sub btnFirst_Click()
    objDB.Move "first", db_RS(5)
    ReLoad_Subj_Rec
End Sub

Private Sub btnLast_Click()
    objDB.Move "last", db_RS(5)
    ReLoad_Subj_Rec
End Sub

Private Sub btnNext_Click()
    objDB.Move "next", db_RS(5)
    ReLoad_Subj_Rec
End Sub

Private Sub btnPrev_Click()
    objDB.Move "prev", db_RS(5)
    ReLoad_Subj_Rec
End Sub

Private Sub btnRec_Refresh_Click()

    objDB.ReQuery db_RS(5)
    ReLoad_Subj_Rec

End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)

    Dim data As String
    data = txtInfo(Index).text

    Select Case Index
        Case 1:
            If Not Len(data) > 0 Then: Cancel = True
    End Select
    
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Asc(UCase(Chr(KeyAscii)))
        Case Asc("A") To Asc("Z"), Asc("0") To Asc("9"), vbKeyBack, vbKeyDelete:
        Case Asc("/"), Asc("."), Asc(" "), Asc(","), Asc("-"), Asc("_"), Asc("@"):
        Case Else: KeyAscii = 0
    End Select
    
End Sub

Private Function Check_Form_Fields()

    Check_Form_Fields = True

    If txtInfo(1) = "" Then
    
        txtInfo(1).SetFocus
        lblResult.Caption = "Please appropriately fill in the required form fields."
        
        Check_Form_Fields = False
        
    End If

End Function

Private Sub Search_Form_Fill()

    Set db_RS(6) = New ADODB.RecordSet
    
    db_SQL = "SELECT column_name FROM user_tab_cols WHERE table_name = 'CM_SUBJECT'"
    objDB.Query db_SQL, db_RS(6)
    
    For i = 0 To (db_RS(6).RecordCount - 1)
    
        cmbSearch_Field.AddItem LCase(db_RS(6).fields(0))
        objDB.Move "next", db_RS(6)
        
    Next i
    
    cmbSearch_Type.AddItem "Exact", 0
    cmbSearch_Type.AddItem "Like", 1
    
    cmbSearch_Start.AddItem "Current Record", 0
    cmbSearch_Start.AddItem "Beginning", 1
    
    cmbSearch_Field.ListIndex = 1
    cmbSearch_Type.ListIndex = 1
    cmbSearch_Start.ListIndex = 1

End Sub

Private Sub btnSearch_Submit_Click()

    Search_Recordset Me, db_RS(5)
    ReLoad_Subj_Rec

End Sub

Private Sub txtSearch_Input_GotFocus()
    btnSearch_Submit.Default = True
End Sub

Private Sub txtSearch_Input_LostFocus()
    btnSearch_Submit.Default = False
End Sub

Private Sub lstCourses_DblClick()
    frm_courses.View Get_Identifier(lstCourses.List(lstCourses.ListIndex))
End Sub

Public Sub Add()

    If blnAdd <> True And blnEdit <> True Then
    
        Me.Show
        btnRec_Add_Click
        Me.SetFocus
        
    End If

End Sub

Public Sub View(ByVal ID As Integer)

    If blnAdd <> True And blnEdit <> True Then
    
        Me.Show
        
        txtSearch_Input.text = Str(ID)
        
        Select_List_Item cmbSearch_Field, "subject_id"
        
        cmbSearch_Type.ListIndex = 0
        cmbSearch_Start.ListIndex = 1
        
        btnSearch_Submit_Click
        
        Me.SetFocus
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_subjects = Nothing
End Sub
