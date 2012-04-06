VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frm_main_mdi 
   Appearance      =   0  'Flat
   BackColor       =   &H00020002&
   Caption         =   "Coaching Management"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11985
   Icon            =   "frm_main_mdi.frx":0000
   LinkTopic       =   "main_sdi"
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picToggle 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   180
      ScaleWidth      =   11985
      TabIndex        =   1
      ToolTipText     =   "Toggle Pane : One-Click Access"
      Top             =   6675
      Width           =   11985
   End
   Begin VB.PictureBox picPane 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F5F5F5&
      Height          =   1400
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   11925
      TabIndex        =   0
      Top             =   6855
      Width           =   11985
      Begin VB.CommandButton btnStaff 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnSubject 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Subjects"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnCourse 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Courses"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4200
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnStudent 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Students"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnFinance 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Finance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton btnReport 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6840
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog Common_Dialog 
      Left            =   11520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Orientation     =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_New 
         Caption         =   "&New"
         Begin VB.Menu mnuFile_New_Student 
            Caption         =   "Student"
         End
         Begin VB.Menu mnuFile_New_Staff 
            Caption         =   "Staff"
         End
         Begin VB.Menu mnuFile_New_Course 
            Caption         =   "Course"
         End
         Begin VB.Menu mnuFile_New_Subject 
            Caption         =   "Subject"
         End
      End
      Begin VB.Menu mnuFile_Close 
         Caption         =   "Cl&ose"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Begin VB.Menu mnuEdit_Cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit_Paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit_Delete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_SelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuManage 
      Caption         =   "&Manage"
      Begin VB.Menu mnuManage_Students 
         Caption         =   "&Students"
      End
      Begin VB.Menu mnuManage_Staff 
         Caption         =   "&Teachers / Staff"
      End
      Begin VB.Menu mnuManage_Courses 
         Caption         =   "&Courses"
      End
      Begin VB.Menu mnuManage_Subjects 
         Caption         =   "Su&bjects"
      End
      Begin VB.Menu mnuManage_Assoc 
         Caption         =   "Associations"
         Begin VB.Menu mnuManage_AssocSC 
            Caption         =   "Student - Course"
         End
         Begin VB.Menu mnuManage_AssocCST 
            Caption         =   "Course - Subject - Staff"
         End
      End
   End
   Begin VB.Menu mnuFinance 
      Caption         =   "&Finance"
      Enabled         =   0   'False
      Begin VB.Menu mnuFinance_Fees 
         Caption         =   "&Fees"
      End
      Begin VB.Menu mnuFinance_Salary 
         Caption         =   "Pay / &Salary"
      End
      Begin VB.Menu mnuFinance_Misc 
         Caption         =   "Misc &Expenses"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReport_Quick 
         Caption         =   "&Quick Stats"
      End
      Begin VB.Menu mnuReport_Finance 
         Caption         =   "Finance &Statements"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReport_Test 
         Caption         =   "&Test Reports"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReport_Student 
         Caption         =   "Student &Overview"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReport_Custom 
         Caption         =   "Custom &Reports"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindow_TileH 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWindow_TileV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindow_Cascade 
         Caption         =   "&Cascade Windows"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_Help 
         Caption         =   "Help &Topics"
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frm_main_mdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnPaneOpen As Boolean

Private Sub MDIForm_Load()
  
    Me.Picture = LoadPicture(Dir_Icon & "abstract_burst.jpg")
    Me.Icon = LoadPicture(Dir_Icon & "app_icon.ico")
    
    picToggle.Picture = LoadPicture(Dir_Icon & "toggle.gif")
    
    btnStudent.Picture = LoadPicture(Dir_Icon & "student.gif")
    btnStaff.Picture = LoadPicture(Dir_Icon & "staff.gif")
    btnSubject.Picture = LoadPicture(Dir_Icon & "subjects.gif")
    btnCourse.Picture = LoadPicture(Dir_Icon & "courses.gif")
    btnReport.Picture = LoadPicture(Dir_Icon & "reports.gif")
    btnFinance.Picture = LoadPicture(Dir_Icon & "finance.gif")
    
    blnPaneOpen = True

    'frm_test.Show
    'frm_report_quick.Show
    'frm_report_grid.Show

End Sub

' MENU -> FILE

Private Sub mnuFile_Exit_Click()

    If MsgBox("Are you sure you want to exit?", vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If

    objDB.Disconnect
    End

End Sub

Private Sub mnuFile_New_Student_Click()
    frm_students.Add
End Sub

Private Sub mnuFile_New_Staff_Click()
    frm_staff.Add
End Sub

Private Sub mnuFile_New_Subject_Click()
    frm_subjects.Add
End Sub

Private Sub mnuFile_New_Course_Click()
    frm_courses.Add
End Sub


' MENU -> MANAGE

Private Sub mnuManage_Courses_Click()
    Load_Course
End Sub

Private Sub mnuManage_Staff_Click()
    Load_Staff
End Sub

Private Sub mnuManage_Students_Click()
    Load_Student
End Sub

Private Sub mnuManage_Subjects_Click()
    Load_Subject
End Sub

Private Sub mnuManage_AssocCST_Click()
    Load_Assoc_CST
End Sub

Private Sub mnuManage_AssocSC_Click()
    Load_Assoc_SC
End Sub

' MENU -> REPORTS

Private Sub mnuReport_Quick_Click()
    frm_report_quick.Show
End Sub

' MENU -> WINDOWS

Private Sub mnuWindow_Cascade_Click()
    frm_main_mdi.Arrange vbCascade
End Sub

Private Sub mnuWindow_TileH_Click()
    frm_main_mdi.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindow_TileV_Click()
    frm_main_mdi.Arrange vbTileVertical
End Sub

' MENU -> HELP

Private Sub mnuHelp_Help_Click()

    Dim msg As String
    msg = "Why do you need help to run such a simple application ?" & vbNewLine _
        & "Almost everything is self explanatory." & vbNewLine & vbNewLine _
        & "Nevertheless, if you need help, contact Shadab. LOL."
        
    MsgBox msg, vbQuestion, "Sorry, But..."
    
End Sub

Private Sub mnuHelp_About_Click()
    frm_about.Show vbModal
End Sub

' PANE :: BUTTONS

Private Sub btnStudent_Click()
    Load_Student
End Sub

Private Sub btnSubject_Click()
    Load_Subject
End Sub

Private Sub btnCourse_Click()
    Load_Course
End Sub

Private Sub btnStaff_Click()
    Load_Staff
End Sub

' PANE :: TOGGLE

Private Sub picToggle_Click()
    
    If blnPaneOpen = True Then
    
        picPane.Height = 0
        blnPaneOpen = False
        
    ElseIf blnPaneOpen = False Then
    
        picPane.Height = 1400
        blnPaneOpen = True
        
    End If

End Sub
