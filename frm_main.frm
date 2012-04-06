VERSION 5.00
Begin VB.Form frm_report_quick 
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports » Quick Statistics"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleMode       =   0  'User
   ScaleWidth      =   8942.084
   Begin VB.Frame frameStats 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Quick Statistics"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   3750
         X2              =   3750
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Label lblStat_Attendance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         TabIndex        =   16
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lblStat_Popular 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Highest Attendance :"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Most Popular Course :"
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
         Left            =   3960
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblStat_Profit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6600
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblStat_Expense 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6600
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblStat_Staff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblStat_Teachers 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblStat_Courses 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblStat_Students 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00EDEDED&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Profit (Quarter, INR) :"
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
         Left            =   3960
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00EDEDED&
         BackStyle       =   0  'Transparent
         Caption         =   "Expenses (Month, INR) :"
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
         Left            =   3960
         TabIndex        =   5
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00EDEDED&
         BackStyle       =   0  'Transparent
         Caption         =   "Staff (Incl teachers) :"
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
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00EDEDED&
         BackStyle       =   0  'Transparent
         Caption         =   "Teachers on Duty :"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F7F7F7&
         BackStyle       =   0  'Transparent
         Caption         =   "Courses Available :"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Students : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frm_report_quick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call Build_Stats
End Sub

Public Sub Build_Stats()

    Set db_RS(1) = New ADODB.RecordSet

    db_SQL = "SELECT " _
            & "(select count(*) from cm_student where enabled = '1') students, " _
            & "(select count(*) from cm_staff where enabled = '1') staff, " _
            & "(select count(*) from cm_staff where job_type = 'Teacher' AND enabled = '1') teachers, " _
            & "(select count(*) from cm_course where enabled = '1') courses " _
            & "FROM dual"

    objDB.Query db_SQL, db_RS(1)

    With db_RS(1)
        lblStat_Students.Caption = !students
        lblStat_Staff.Caption = !staff
        lblStat_Teachers.Caption = !teachers
        lblStat_Courses.Caption = !courses
    End With

    Dim strTemp As String

    objDB.Exec_Proc_Out "procPopularCourse", strTemp
    lblStat_Popular.Caption = strTemp

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_report_quick = Nothing
End Sub
