VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_report_grid 
   BackColor       =   &H00F2F2F2&
   Caption         =   "Reports » Grid"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlexGrid 
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11245
      _Version        =   393216
      ScrollTrack     =   -1  'True
      GridLines       =   2
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm_report_grid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Set db_RS(30) = New ADODB.RecordSet
    db_SQL = "SELECT * FROM cm_student ORDER BY student_id"

    objDB.Query db_SQL, db_RS(30)
    Set FlexGrid.RecordSet = db_RS(30)

End Sub

Private Sub Form_Resize()

    If Me.WindowState <> 1 Then
        FlexGrid.Width = Me.Width - 750
        FlexGrid.Height = Me.Height - 2500
    End If
    
End Sub

Private Sub FlexGrid_Click()
    MsgBox FlexGrid.TextMatrix(FlexGrid.RowSel, FlexGrid.ColSel)
End Sub
