VERSION 5.00
Begin VB.Form frm_courses 
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Courses Available"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   17
      Top             =   6480
      Width           =   1095
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
      TabIndex        =   16
      Top             =   6480
      Width           =   1095
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
      TabIndex        =   15
      ToolTipText     =   "Save all the changes done to the database"
      Top             =   6480
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
      TabIndex        =   14
      ToolTipText     =   "EDIT the currently displayed record"
      Top             =   6480
      Width           =   1215
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
      TabIndex        =   13
      ToolTipText     =   "Add a NEW Record"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton btnLast 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   625
   End
   Begin VB.CommandButton btnNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   4395
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   625
   End
   Begin VB.CommandButton btnPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   625
   End
   Begin VB.CommandButton btnFirst 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   625
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Enrolled Students"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5160
      TabIndex        =   34
      Top             =   3360
      Width           =   4455
      Begin VB.ListBox lstStudents 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         ItemData        =   "frm_courses.frx":0000
         Left            =   240
         List            =   "frm_courses.frx":0002
         TabIndex        =   8
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Associated Subjects"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5160
      TabIndex        =   33
      Top             =   1560
      Width           =   4455
      Begin VB.ListBox lstSubjects 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         ItemData        =   "frm_courses.frx":0004
         Left            =   240
         List            =   "frm_courses.frx":0006
         TabIndex        =   7
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Course Info."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   26
      Top             =   1560
      Width           =   4695
      Begin VB.CheckBox chkVisible 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Visible in the Records or Not ? (Check = Yes)"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3720
         Width           =   3735
      End
      Begin VB.ComboBox cmbInfo 
         Height          =   315
         Index           =   2
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Top             =   2520
         Width           =   2175
      End
      Begin VB.ComboBox cmbInfo 
         Height          =   315
         Index           =   1
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox cmbInfo 
         Height          =   315
         Index           =   0
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
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
         Left            =   2040
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
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
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00CCCCCC&
         X1              =   0
         X2              =   4680
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00CCCCCC&
         X1              =   0
         X2              =   4680
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Course Type :"
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
         TabIndex        =   32
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fees (Rs) :"
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
         TabIndex        =   31
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Duration :"
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
         TabIndex        =   30
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Class / Grade :"
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
         TabIndex        =   29
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F2F2F2&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Name :"
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
         TabIndex        =   28
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label50 
         BackColor       =   &H00F2F2F2&
         BackStyle       =   0  'Transparent
         Caption         =   "Course ID :"
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
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F2F2F2&
      Height          =   855
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   9375
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
         ItemData        =   "frm_courses.frx":0008
         Left            =   6840
         List            =   "frm_courses.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Start searching FROM"
         Top             =   360
         Width           =   1455
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
         TabIndex        =   22
         Top             =   360
         Width           =   735
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
         ItemData        =   "frm_courses.frx":000C
         Left            =   5640
         List            =   "frm_courses.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         ToolTipText     =   "Search Types"
         Top             =   360
         Width           =   1095
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
         TabIndex        =   18
         ToolTipText     =   "Text to search FOR"
         Top             =   360
         Width           =   2415
      End
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
         ItemData        =   "frm_courses.frx":0010
         Left            =   3600
         List            =   "frm_courses.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         ToolTipText     =   "Table Field to Search IN"
         Top             =   360
         Width           =   1935
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
         TabIndex        =   25
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   6000
      Width           =   9375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Available Courses Info."
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
      TabIndex        =   23
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frm_courses"
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
    Fill_Combo_Lists
    
    Load_Course_Rec

End Sub

Private Sub Load_Course_Rec()

    Set db_RS(9) = New ADODB.RecordSet

    db_SQL = "SELECT " _
            & "course_id, course_name, fees, " _
            & "class, duration, type, enabled " _
            & "FROM cm_course ORDER BY course_id"

    objDB.Query db_SQL, db_RS(9)

    ReLoad_Course_Rec

End Sub

Private Sub ReLoad_Course_Rec()

    For i = 0 To 2
        txtInfo(i).text = db_RS(9).fields(i) & ""
        Select_List_Item cmbInfo(i), db_RS(9).fields(i + 3)
    Next i

    chkVisible.Value = db_RS(9).fields(6)
    
    List_Assoc_Subjects
    List_Assoc_Students

End Sub

Private Sub Fill_Combo_Lists()

    With cmbInfo(0)
        .AddItem "9", 0
        .AddItem "10", 1
        .AddItem "11", 2
        .AddItem "12", 3
    End With

    With cmbInfo(1)
        .AddItem "1 Year", 0
        .AddItem "2 Months", 1
        .AddItem "1 Month", 2
    End With

    With cmbInfo(2)
        .AddItem "Fresher", 0
        .AddItem "Dropper", 1
        .AddItem "Crash", 2
    End With

End Sub

Private Sub List_Assoc_Subjects()

    lstSubjects.Clear
    
    Set db_RS(12) = New ADODB.RecordSet
    
    db_SQL = "SELECT T.subject_id, S.subject_name " _
            & "FROM cm_course C, cm_course_subject_teacher T, cm_subject S " _
            & "WHERE C.course_id = T.course_id AND S.subject_id = T.subject_id " _
            & "AND C.course_id = " & db_RS(9).fields(0) _
            & " ORDER BY C.course_id, T.subject_id"

    objDB.Query db_SQL, db_RS(12)
    
    With db_RS(12)
        If .RecordCount > 0 Then
            For i = 0 To (.RecordCount - 1)
    
                lstSubjects.AddItem !subject_id & " - " & !subject_name
                objDB.Move "next", db_RS(12)
                
            Next i
        End If
    End With

End Sub

Private Sub List_Assoc_Students()

    lstStudents.Clear
    
    Set db_RS(13) = New ADODB.RecordSet
    
    db_SQL = "SELECT C.student_id ""ID"", S.first_name ""FName"", S.last_name ""LName"", " _
            & "TO_CHAR(C.date_started, 'DD-Mon-YY') ""Date"" " _
            & "FROM cm_student_course C, cm_student S " _
            & "WHERE C.student_id = S.student_id " _
            & "AND C.course_id = " & db_RS(9).fields(0) _
            & " ORDER BY S.student_id"

    objDB.Query db_SQL, db_RS(13)

    With db_RS(13)
        If .RecordCount > 0 Then
            For i = 0 To (.RecordCount - 1)

                lstStudents.AddItem !ID & " - " & !fname & " " & !lname & " (" & !Date & ")"
                objDB.Move "next", db_RS(13)

            Next i
        End If
    End With

End Sub

Private Sub Lock_Controls()
    
    For i = 0 To 2
        txtInfo(i).Locked = True
        cmbInfo(i).Locked = True
    Next i
    
    chkVisible.Enabled = False
    btnRec_Save.Enabled = False
    btnRec_Delete.Enabled = False

End Sub

Private Sub UnLock_Controls()

    For i = 0 To 2
        txtInfo(i).Locked = False
        cmbInfo(i).Locked = False
    Next i
    
    txtInfo(0).Locked = True
    
    chkVisible.Enabled = True
    btnRec_Save.Enabled = True

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

Private Sub btnFirst_Click()
    objDB.Move "first", db_RS(9)
    ReLoad_Course_Rec
End Sub

Private Sub btnLast_Click()
    objDB.Move "last", db_RS(9)
    ReLoad_Course_Rec
End Sub

Private Sub btnNext_Click()
    objDB.Move "next", db_RS(9)
    ReLoad_Course_Rec
End Sub

Private Sub btnPrev_Click()
    objDB.Move "prev", db_RS(9)
    ReLoad_Course_Rec
End Sub

Private Sub btnRec_Refresh_Click()

    objDB.ReQuery db_RS(9)
    ReLoad_Course_Rec

End Sub

Private Sub btnRec_Add_Click()

    If blnAdd <> True Then
    
        btnRec_Add.Caption = "CANCEL"
        btnRec_Edit.Enabled = False
        
        UnLock_Controls
        Lock_Navigation
        
        For i = 0 To 2
            txtInfo(i).text = ""
            cmbInfo(i).ListIndex = 0
        Next i
        
        lstSubjects.Clear
        lstStudents.Clear
    
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
    
        blnEdit = False
    
    End If

End Sub

Private Sub btnRec_Save_Click()

    If blnEdit = True Then
    
        Set db_RS(11) = New ADODB.RecordSet

        db_SQL = "UPDATE cm_course SET " _
                & "   COURSE_NAME   = '" & txtInfo(1).text _
                & "', CLASS         =  " & cmbInfo(0).List(cmbInfo(0).ListIndex) _
                & " , DURATION      = '" & cmbInfo(1).List(cmbInfo(1).ListIndex) _
                & "', FEES          =  " & txtInfo(2).text _
                & " , TYPE          = '" & cmbInfo(2).List(cmbInfo(2).ListIndex) _
                & "', ENABLED       = '" & chkVisible.Value _
                & "' WHERE COURSE_ID = " & txtInfo(0).text
               
        objDB.Execute db_SQL, sngTemp
        
        frm_test3.Show
        frm_test3.fill (db_SQL & " . And sngTemp is = " & sngTemp)
        
        If sngTemp > 0 Then
            lblResult.Caption = "Data edited / updated accordingly."
        Else
            lblResult.Caption = "Data could NOT be edited / updated."
        End If
    
    ElseIf blnAdd = True Then

        Set db_RS(11) = New ADODB.RecordSet

        db_SQL = "INSERT INTO cm_course( COURSE_ID, TYPE, COURSE_NAME, " _
                & "DURATION, ENABLED, FEES, CLASS) VALUES( " _
                    & "seq_courseid.nextval , " _
                    & "'" & cmbInfo(2).List(cmbInfo(2).ListIndex) & "', " _
                    & "'" & txtInfo(1).text & "' , " _
                    & "'" & cmbInfo(1).List(cmbInfo(1).ListIndex) & "' , " _
                    & "'" & chkVisible.Value & "', " _
                    & txtInfo(2).text & " , " _
                    & cmbInfo(0).List(cmbInfo(0).ListIndex) _
                & " )"

        objDB.Execute db_SQL, sngTemp

        frm_test3.Show
        frm_test3.fill (db_SQL & " . And sngTemp is = " & sngTemp)

        lblResult.Caption = "New Data added accordingly"

        btnRec_Add_Click
        btnLast_Click
        btnRec_Edit_Click
    
    End If

End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Asc(UCase(Chr(KeyAscii)))
        Case Asc("A") To Asc("Z"), Asc("0") To Asc("9"), vbKeyBack, vbKeyDelete:
        Case Asc("/"), Asc("."), Asc(" "), Asc(","), Asc("-"), Asc("_"), Asc("@"):
        Case Else: KeyAscii = 0
    End Select
    
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)

    Dim data As String
    data = txtInfo(Index).text

    Select Case Index
        Case 1, 2:
            If Not Len(data) > 0 Then: Cancel = True
    End Select

End Sub

Private Sub Search_Form_Fill()

    Set db_RS(10) = New ADODB.RecordSet
    
    db_SQL = "SELECT column_name FROM user_tab_cols WHERE table_name = 'CM_COURSE'"
    objDB.Query db_SQL, db_RS(10)
    
    For i = 0 To (db_RS(10).RecordCount - 1)
    
        cmbSearch_Field.AddItem LCase(db_RS(10).fields(0))
        objDB.Move "next", db_RS(10)
        
    Next i
    
    cmbSearch_Type.AddItem "Exact", 0
    cmbSearch_Type.AddItem "Like", 1
    
    cmbSearch_Start.AddItem "Current Record", 0
    cmbSearch_Start.AddItem "Beginning", 1
    
    cmbSearch_Field.ListIndex = 2
    cmbSearch_Type.ListIndex = 1
    cmbSearch_Start.ListIndex = 1

End Sub

Private Sub btnSearch_Submit_Click()

    Search_Recordset Me, db_RS(9)
    ReLoad_Course_Rec
    
End Sub

Private Sub txtSearch_Input_GotFocus()
    btnSearch_Submit.Default = True
End Sub

Private Sub txtSearch_Input_LostFocus()
    btnSearch_Submit.Default = False
End Sub

Private Sub lstSubjects_DblClick()
    frm_subjects.View Get_Identifier(lstSubjects.List(lstSubjects.ListIndex))
End Sub

Private Sub lstStudents_DblClick()
    frm_students.View Get_Identifier(lstStudents.List(lstStudents.ListIndex))
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
        
        Select_List_Item cmbSearch_Field, "course_id"
        
        cmbSearch_Type.ListIndex = 0
        cmbSearch_Start.ListIndex = 1
        
        btnSearch_Submit_Click
        
        Me.SetFocus
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_courses = Nothing
End Sub
