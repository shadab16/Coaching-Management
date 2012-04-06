VERSION 5.00
Begin VB.Form frm_students 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Student Information"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   657
   Begin VB.Frame Frame6 
      BackColor       =   &H00F2F2F2&
      Height          =   735
      Left            =   240
      TabIndex        =   55
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
         ItemData        =   "frm_students.frx":0000
         Left            =   6840
         List            =   "frm_students.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         ToolTipText     =   "Start searching FROM"
         Top             =   240
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
         TabIndex        =   34
         Top             =   240
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
         ItemData        =   "frm_students.frx":0004
         Left            =   5640
         List            =   "frm_students.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         ToolTipText     =   "Search Types"
         Top             =   240
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
         TabIndex        =   30
         ToolTipText     =   "Text to search FOR"
         Top             =   240
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
         ItemData        =   "frm_students.frx":0008
         Left            =   3600
         List            =   "frm_students.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         ToolTipText     =   "Table Field to Search IN"
         Top             =   240
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
         TabIndex        =   56
         Top             =   270
         Width           =   855
      End
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
      TabIndex        =   25
      Top             =   7230
      Width           =   1095
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
      Left            =   5160
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
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
      TabIndex        =   24
      Top             =   7230
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
      TabIndex        =   23
      ToolTipText     =   "Save all the changes done to the database"
      Top             =   7230
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
      TabIndex        =   22
      ToolTipText     =   "EDIT the currently displayed record"
      Top             =   7230
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
      TabIndex        =   21
      ToolTipText     =   "Add a NEW Record"
      Top             =   7230
      Width           =   1215
   End
   Begin VB.CommandButton btnLast 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7230
      Width           =   625
   End
   Begin VB.CommandButton btnNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   4395
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7230
      Width           =   625
   End
   Begin VB.CommandButton btnPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7230
      Width           =   625
   End
   Begin VB.CommandButton btnFirst 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7230
      Width           =   625
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Additional Notes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7080
      TabIndex        =   38
      Top             =   4440
      Width           =   2535
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
         Height          =   1935
         Index           =   15
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Photograph"
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
      Left            =   7080
      TabIndex        =   37
      Top             =   1560
      Width           =   2535
      Begin VB.CommandButton btnPic_Remove 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton btnPic_Add 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Add Picture"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Image picStudent 
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   120
         MousePointer    =   2  'Cross
         Stretch         =   -1  'True
         ToolTipText     =   "Click on the thumbnail to see an enlarged specimen."
         Top             =   240
         Width           =   2295
      End
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
      Index           =   14
      Left            =   1800
      TabIndex        =   14
      Top             =   6240
      Width           =   1575
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
      Index           =   13
      Left            =   5160
      TabIndex        =   13
      Top             =   5760
      Width           =   1575
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
      Index           =   12
      Left            =   1800
      TabIndex        =   12
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Miscellaneous"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   36
      Top             =   5400
      Width           =   6615
      Begin VB.CheckBox chkVisible 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Visible in the Records or Not ?"
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
         Left            =   3480
         TabIndex        =   16
         Top             =   840
         Width           =   3015
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00CCCCCC&
         X1              =   3320
         X2              =   3320
         Y1              =   120
         Y2              =   1320
      End
      Begin VB.Label Label16 
         BackColor       =   &H00F2F2F2&
         Caption         =   "EMail ID"
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
         TabIndex        =   53
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Leave Date"
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
         Left            =   3480
         TabIndex        =   52
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Admission Date"
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
         TabIndex        =   51
         Top             =   360
         Width           =   1335
      End
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
      Index           =   5
      Left            =   5160
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
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
      Index           =   6
      Left            =   1800
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
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
      Index           =   4
      Left            =   1800
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
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
      Index           =   11
      Left            =   5160
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
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
      Index           =   10
      Left            =   1800
      TabIndex        =   10
      Top             =   4800
      Width           =   1575
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
      Index           =   9
      Left            =   5160
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
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
      Index           =   8
      Left            =   1800
      TabIndex        =   8
      Top             =   4320
      Width           =   1575
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
      Index           =   7
      Left            =   5160
      TabIndex        =   7
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   35
      Top             =   3480
      Width           =   6615
      Begin VB.Line Line2 
         BorderColor     =   &H00CCCCCC&
         X1              =   3320
         X2              =   3320
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Label Label13 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Cell Number"
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
         Left            =   3480
         TabIndex        =   50
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Home Phone"
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
         TabIndex        =   49
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Pin Code"
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
         Left            =   3480
         TabIndex        =   48
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00F2F2F2&
         Caption         =   "City / District"
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
         TabIndex        =   47
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Address Line 2"
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
         Left            =   3480
         TabIndex        =   46
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Address Line 1"
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
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
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
      Index           =   3
      Left            =   5160
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
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
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
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
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Personal Info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   29
      Top             =   1560
      Width           =   6615
      Begin VB.Line Line1 
         BorderColor     =   &H00CCCCCC&
         X1              =   3320
         X2              =   3320
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Label Label7 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Mother's Name"
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
         Left            =   3480
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Last Name"
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
         Left            =   3480
         TabIndex        =   43
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Date of Birth"
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
         Left            =   3480
         TabIndex        =   42
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Father's Name"
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
         TabIndex        =   41
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F2F2F2&
         Caption         =   "First Name"
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
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label50 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Student ID"
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
         TabIndex        =   39
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FBFBFB&
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
      Height          =   300
      Left            =   240
      TabIndex        =   54
      Top             =   6825
      Width           =   9375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student Information"
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
      TabIndex        =   28
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frm_students"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blnEdit, blnAdd, blnPicture As Boolean
Dim strPicPath, strPicName, strPicDefault As String
Dim sngTemp As Integer

Private Sub Form_Load()

    btnFirst.Picture = LoadPicture(Dir_Icon & "first.gif")
    btnPrev.Picture = LoadPicture(Dir_Icon & "back.gif")
    btnNext.Picture = LoadPicture(Dir_Icon & "next.gif")
    btnLast.Picture = LoadPicture(Dir_Icon & "last.gif")

    Lock_Controls
    Search_Form_Fill

    Load_Stud_Rec

End Sub

Private Sub Load_Stud_Rec()

    Set db_RS(2) = New ADODB.RecordSet

    db_SQL = "SELECT " _
            & "STUDENT_ID, BIRTH_DATE, FIRST_NAME, LAST_NAME, " _
            & "FATHER_NAME, MOTHER_NAME, ADDRESS_LINE_1, ADDRESS_LINE_2, " _
            & "CITY, PIN_CODE, PHONE_NUMBER, CELL_NUMBER, " _
            & "DATE_ADMISSION, DATE_LEAVE, EMAIL, MISC_INFO, Enabled " _
            & "FROM cm_student ORDER BY student_id"

    objDB.Query db_SQL, db_RS(2)

    ReLoad_Stud_Rec

End Sub

Private Sub ReLoad_Stud_Rec()

    For i = 0 To 15
        txtInfo(i).text = db_RS(2).fields(i) & ""
    Next i
    
    chkVisible.Value = db_RS(2).fields(16)
    
    Load_Stud_Pic

End Sub

Private Sub Lock_Controls()

    For i = 0 To (txtInfo.Count - 1)
        txtInfo(i).Locked = True
    Next i
    
    chkVisible.Enabled = False

    btnPic_Add.Enabled = False
    btnPic_Remove.Enabled = False
    
    btnRec_Save.Enabled = False
    btnRec_Delete.Enabled = False

End Sub

Private Sub UnLock_Controls()

    For i = 1 To (txtInfo.Count - 1)
        txtInfo(i).Locked = False
    Next i
    
    chkVisible.Enabled = True
    
    btnPic_Add.Enabled = True
    
    If blnPicture = True Then
        btnPic_Remove.Enabled = True
    End If
    
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

Private Sub Load_Stud_Pic()
    
    strPicName = "stud_" & db_RS(2).fields(0) & ".jpg"
    strPicPath = Dir_Pics & strPicName
    strPicDefault = Dir_Pics & "stud_0.jpg"

    blnPicture = objFSO.FileExists(strPicPath)

    If blnPicture = True Then
        picStudent.Picture = LoadPicture(strPicPath)
    Else
        picStudent.Picture = LoadPicture(strPicDefault)
    End If

End Sub

Private Sub picStudent_Click()

    If blnPicture = True Then
        Call frm_view_pic.ShowPic(strPicPath, txtInfo(2) & " " & txtInfo(3))
    End If

End Sub

Private Sub btnFirst_Click()
    objDB.Move "first", db_RS(2)
    ReLoad_Stud_Rec
End Sub

Private Sub btnLast_Click()
    objDB.Move "last", db_RS(2)
    ReLoad_Stud_Rec
End Sub

Private Sub btnNext_Click()
    objDB.Move "next", db_RS(2)
    ReLoad_Stud_Rec
End Sub

Private Sub btnPrev_Click()
    objDB.Move "prev", db_RS(2)
    ReLoad_Stud_Rec
End Sub

Private Sub btnRec_Refresh_Click()

    objDB.ReQuery db_RS(2)
    ReLoad_Stud_Rec

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

Private Sub btnRec_Add_Click()

    If blnAdd <> True Then
    
        btnRec_Add.Caption = "CANCEL"
        btnRec_Edit.Enabled = False
        
        UnLock_Controls
        Lock_Navigation
        
        btnPic_Remove.Enabled = False
        btnPic_Add.Enabled = False
        
        For i = 0 To (txtInfo.Count - 1)
            txtInfo(i).text = ""
        Next i
        
        picStudent.Picture = LoadPicture(strPicDefault)
    
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

Private Sub btnRec_Save_Click()

    If Check_Form_Fields() = False Then
        Exit Sub
    End If
    
    If blnEdit = True Then
        
        Set db_RS(3) = New ADODB.RecordSet
        
        db_SQL = "UPDATE cm_student SET " _
                & "  BIRTH_DATE         = TO_DATE('" & txtInfo(1).text & "', 'MM/DD/YYYY')" _
                & ", FIRST_NAME         = '" & txtInfo(2).text _
                & "', LAST_NAME         = '" & txtInfo(3).text _
                & "', FATHER_NAME       = '" & txtInfo(4).text _
                & "', MOTHER_NAME       = '" & txtInfo(5).text _
                & "', ADDRESS_LINE_1    = '" & txtInfo(6).text _
                & "', ADDRESS_LINE_2    = '" & txtInfo(7).text _
                & "', CITY              = '" & txtInfo(8).text _
                & "', PIN_CODE          = TO_NUMBER('" & txtInfo(9).text & "')" _
                & ", PHONE_NUMBER       = '" & txtInfo(10).text _
                & "', CELL_NUMBER       = '" & txtInfo(11).text _
                & "', DATE_ADMISSION    = TO_DATE('" & txtInfo(12).text & "', 'MM/DD/YYYY')" _
                & ", DATE_LEAVE         = TO_DATE('" & txtInfo(13).text & "', 'MM/DD/YYYY')" _
                & ", EMAIL              = '" & txtInfo(14).text _
                & "', MISC_INFO         = '" & MultiLine_To_Single(txtInfo(15).text) _
                & "', ENABLED           = '" & chkVisible.Value _
                & "' WHERE STUDENT_ID   = " & txtInfo(0).text
               
        objDB.Execute db_SQL, sngTemp
        
        frm_test3.Show
        frm_test3.fill (db_SQL & " . And sngTemp is = " & sngTemp)
        
        If sngTemp > 0 Then
            lblResult.Caption = "Data edited / updated accordingly."
        Else
            lblResult.Caption = "Data could NOT be edited / updated."
        End If
        
    ElseIf blnAdd = True Then
    
        Set db_RS(3) = New ADODB.RecordSet
        
        db_SQL = "INSERT INTO cm_student(STUDENT_ID, BIRTH_DATE, " _
                & "FIRST_NAME, LAST_NAME, FATHER_NAME, MOTHER_NAME, " _
                & "ADDRESS_LINE_1, ADDRESS_LINE_2, CITY, PIN_CODE, " _
                & "PHONE_NUMBER, CELL_NUMBER, DATE_ADMISSION, DATE_LEAVE, " _
                & "EMAIL, MISC_INFO, ENABLED) VALUES(" _
                    & "seq_studentid.nextval, " _
                    & "TO_DATE('" & txtInfo(1).text & "') , '" _
                    & txtInfo(2).text & "' , '" _
                    & txtInfo(3).text & "' , '" _
                    & txtInfo(4).text & "' , '" _
                    & txtInfo(5).text & "' , '" _
                    & txtInfo(6).text & "' , '" _
                    & txtInfo(7).text & "' , '" _
                    & txtInfo(8).text & "' , " _
                    & "TO_NUMBER('" & txtInfo(9).text & "') , '" _
                    & txtInfo(10).text & "' , '" _
                    & txtInfo(11).text & "' , " _
                    & "TO_DATE('" & txtInfo(12).text & "') , " _
                    & "TO_DATE('" & txtInfo(13).text & "') , '" _
                    & txtInfo(14).text & "' , '" _
                    & MultiLine_To_Single(txtInfo(15).text) & "' , '" _
                    & chkVisible.Value & "' " _
                & ")"
        
        objDB.Execute db_SQL, sngTemp

        frm_test3.Show
        frm_test3.fill (db_SQL & " . And sngTemp is = " & sngTemp)
        
        lblResult.Caption = "New Data added accordingly"
                
        btnRec_Add_Click
        btnLast_Click
        btnRec_Edit_Click

    End If

End Sub

Private Sub btnRec_Delete_Click()
    'delete the damn record
End Sub

Private Sub btnPic_Add_Click()
    
    With objDialog
    
        .Filter = "JPEG Image (*.jpg) | *.jpg"
        .InitDir = App.Path
        .ShowOpen
        
        If .FileName <> "" And .FileName <> strPicPath Then
                   
            If objFSO.FileExists(strPicPath) Then
            
                objFSO.CopyFile strPicPath, Dir_Backup, True
                objFSO.DeleteFile strPicPath
                
            End If
           
            objFSO.CopyFile .FileName, strPicPath, True
            Load_Stud_Pic
            
            blnPicture = True
            btnPic_Remove.Enabled = True
    
        End If
    
    End With
    
End Sub

Private Sub btnPic_Remove_Click()

    objFSO.CopyFile strPicPath, Dir_Backup, True
    objFSO.DeleteFile strPicPath

    Load_Stud_Pic

    blnPicture = False
    btnPic_Remove.Enabled = False
    
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
        Case 2, 3, 6, 7, 8: 'Not null
        
            If Not Len(data) > 0 Then
                Cancel = True
            End If
            
        Case 9: 'Pin Code (Should be numeric)
        
            If Not IsNumeric(data) Then
                Cancel = True
            End If
            
        Case 10: 'Home Phone (Format : 0755-4276675)
            
            Dim tmp1, tmp2, tmp3 As String
            
            tmp1 = Left$(data, 4)
            tmp2 = Mid(data, 5, 1)
            tmp3 = Right$(data, 7)
            
            If (Len(tmp1) <> 4 Or Not IsNumeric(tmp1) Or tmp2 <> "-" _
            Or Len(tmp3) <> 7 Or Not IsNumeric(tmp3)) And data <> "" Then
                Cancel = True
            End If
    
        Case 11: 'Cell Number (Should be numeric too)
        
            If Not IsNumeric(data) And data <> "" Then
                Cancel = True
            End If

        Case 1, 12: 'Not Null + Date format
        
            If Not IsDate(data) Or Not Len(data) > 0 Then
                Cancel = True
            End If
            
        Case 13: 'Date format
            
            If Not IsDate(data) And data <> "" Then
                Cancel = True
            End If

    End Select
    
End Sub

Private Function Check_Form_Fields()

    Dim fields(), i As Variant
    fields() = Array(1, 2, 3, 6, 7, 8, 9, 12)

    Check_Form_Fields = True
    
    For Each i In fields()
        If txtInfo(i) = "" Then
        
            txtInfo(i).SetFocus
            lblResult.Caption = "Please fill in the required fields"
            
            Check_Form_Fields = False
            Exit For

        End If
    Next

End Function

Private Sub Search_Form_Fill()

    Set db_RS(4) = New ADODB.RecordSet
    
    db_SQL = "SELECT column_name FROM user_tab_cols WHERE table_name = 'CM_STUDENT'"
    objDB.Query db_SQL, db_RS(4)
    
    For i = 0 To (db_RS(4).RecordCount - 1)
    
        cmbSearch_Field.AddItem LCase(db_RS(4).fields(0))
        objDB.Move "next", db_RS(4)
        
    Next i
    
    cmbSearch_Type.AddItem "Exact", 0
    cmbSearch_Type.AddItem "Like", 1
    
    cmbSearch_Start.AddItem "Current Record", 0
    cmbSearch_Start.AddItem "Beginning", 1
    
    cmbSearch_Field.ListIndex = 10
    cmbSearch_Type.ListIndex = 1
    cmbSearch_Start.ListIndex = 1

End Sub

Private Sub btnSearch_Submit_Click()

    Search_Recordset Me, db_RS(2)
    ReLoad_Stud_Rec

End Sub

Private Sub txtSearch_Input_GotFocus()
    btnSearch_Submit.Default = True
End Sub

Private Sub txtSearch_Input_LostFocus()
    btnSearch_Submit.Default = False
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
        
        Select_List_Item cmbSearch_Field, "student_id"
        
        cmbSearch_Type.ListIndex = 0
        cmbSearch_Start.ListIndex = 1
        
        btnSearch_Submit_Click
        
        Me.SetFocus
        
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_students = Nothing
End Sub
