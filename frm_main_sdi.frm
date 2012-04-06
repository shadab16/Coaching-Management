VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm frm_main_sdi 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Coaching Management"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11190
   LinkTopic       =   "main_sdi"
   Picture         =   "frm_main_sdi.frx":0000
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   635
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "col1"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_New 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFile_Close 
         Caption         =   "Cl&ose"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
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
   End
   Begin VB.Menu mnuFinance 
      Caption         =   "&Finance"
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
      Begin VB.Menu mnuReport_Finance 
         Caption         =   "Finance &Statements"
      End
      Begin VB.Menu mnuReport_Test 
         Caption         =   "&Test Reports"
      End
      Begin VB.Menu mnuReport_Student 
         Caption         =   "Student &Overview"
      End
      Begin VB.Menu mnuReport_Custom 
         Caption         =   "Custom &Reports"
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
Attribute VB_Name = "frm_main_sdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

    Call Init
    
    frm_main.Show
    frm_test.Show
    
    StatusBar.Panels.Item(1).Text = "Startup:: Total Forms - " & Forms.Count _
                                    & " : Child Forms - " & Get_Total_Forms()

    Debug.Print mnuWindow.WindowList

End Sub

' MENU -> WINDOWS

Private Sub mnuWindow_Cascade_Click()
    frm_main_sdi.Arrange vbCascade
End Sub

Private Sub mnuWindow_TileH_Click()
    frm_main_sdi.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindow_TileV_Click()
    frm_main_sdi.Arrange vbTileVertical
End Sub

' MENU -> HELP

Private Sub mnuHelp_Help_Click()

    Dim msg As String
    msg = "Why do you need help to run such a simple application ?" & vbNewLine _
        & "Almost everything is self explanatory." & vbNewLine & vbNewLine _
        & "Nevertheless, if you need help, contact Shadab. LOL."
        
    MsgBox msg, vbQuestion + vbOKOnly, "Sorry, But..."
    
End Sub

Private Sub mnuHelp_About_Click()
    frm_about.Show vbModal
End Sub
