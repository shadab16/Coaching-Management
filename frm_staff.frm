VERSION 5.00
Begin VB.Form frm_staff 
   BackColor       =   &H00F2F2F2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Staff Information"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9795
   Begin VB.Frame Frame7 
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
      Height          =   1815
      Left            =   240
      TabIndex        =   54
      Top             =   4800
      Width           =   5895
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
         Height          =   1425
         ItemData        =   "frm_staff.frx":0000
         Left            =   240
         List            =   "frm_staff.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.CommandButton btnFirst 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7200
      Width           =   625
   End
   Begin VB.CommandButton btnPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7200
      Width           =   625
   End
   Begin VB.CommandButton btnNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   4395
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7200
      Width           =   625
   End
   Begin VB.CommandButton btnLast 
      Appearance      =   0  'Flat
      BackColor       =   &H00F2F2F2&
      Height          =   495
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7200
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
      TabIndex        =   21
      ToolTipText     =   "Add a NEW Record"
      Top             =   7200
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
      Top             =   7200
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
      TabIndex        =   23
      ToolTipText     =   "Save all the changes done to the database"
      Top             =   7200
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
      TabIndex        =   24
      Top             =   7200
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
      TabIndex        =   25
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Frame Frame6 
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
      Height          =   1815
      Left            =   6240
      TabIndex        =   51
      Top             =   4800
      Width           =   3375
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
         Height          =   1455
         Index           =   13
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00F2F2F2&
      Caption         =   "Job Info"
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
      Left            =   6840
      TabIndex        =   46
      Top             =   1440
      Width           =   2775
      Begin VB.ComboBox cmbInfo 
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
         Left            =   1320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1275
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
         Left            =   1320
         TabIndex        =   12
         Top             =   1800
         Width           =   1275
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
         Index           =   10
         Left            =   1320
         TabIndex        =   11
         Top             =   1320
         Width           =   1275
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
         Index           =   9
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label18 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Salary (INR)"
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
         TabIndex        =   50
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Job Type"
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
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Date Leave"
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
         TabIndex        =   48
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Date of Hire"
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
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   855
      Left            =   6840
      TabIndex        =   45
      Top             =   3840
      Width           =   2775
      Begin VB.CheckBox chkVisible 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Visible in the Records ?"
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
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
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
      TabIndex        =   38
      Top             =   2880
      Width           =   6495
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
         Left            =   4800
         TabIndex        =   4
         Top             =   360
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
         Index           =   5
         Left            =   1560
         TabIndex        =   5
         Top             =   840
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
         Left            =   4800
         TabIndex        =   6
         Top             =   840
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
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
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
         Left            =   4800
         TabIndex        =   8
         Top             =   1320
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
         Index           =   3
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   1575
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
         TabIndex        =   44
         Top             =   360
         Width           =   1335
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
         TabIndex        =   43
         Top             =   360
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
         TabIndex        =   42
         Top             =   840
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
         TabIndex        =   41
         Top             =   840
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
         TabIndex        =   40
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   39
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00CCCCCC&
         X1              =   3320
         X2              =   3320
         Y1              =   120
         Y2              =   1800
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1335
      Left            =   240
      TabIndex        =   34
      Top             =   1440
      Width           =   6495
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
         Left            =   4800
         TabIndex        =   15
         Top             =   360
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
         Left            =   1560
         TabIndex        =   0
         Top             =   360
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
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   840
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
         Left            =   4800
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Email"
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
         TabIndex        =   53
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label50 
         BackColor       =   &H00F2F2F2&
         Caption         =   "Staff ID"
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
         TabIndex        =   37
         Top             =   360
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
         TabIndex        =   36
         Top             =   840
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
         TabIndex        =   35
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00CCCCCC&
         X1              =   3315
         X2              =   3315
         Y1              =   120
         Y2              =   1320
      End
   End
   Begin VB.Frame Frame_Search 
      BackColor       =   &H00F2F2F2&
      Height          =   735
      Left            =   240
      TabIndex        =   32
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
         ItemData        =   "frm_staff.frx":0004
         Left            =   3600
         List            =   "frm_staff.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Table Field to Search IN"
         Top             =   240
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
         TabIndex        =   26
         ToolTipText     =   "Text to search FOR"
         Top             =   240
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
         ItemData        =   "frm_staff.frx":0008
         Left            =   5640
         List            =   "frm_staff.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         ToolTipText     =   "Search Types"
         Top             =   240
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
         TabIndex        =   30
         Top             =   240
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
         ItemData        =   "frm_staff.frx":000C
         Left            =   6840
         List            =   "frm_staff.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         ToolTipText     =   "Start searching FROM"
         Top             =   240
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
         TabIndex        =   33
         Top             =   270
         Width           =   855
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
      TabIndex        =   52
      Top             =   6720
      Width           =   9375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Information"
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
      TabIndex        =   31
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frm_staff"
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

    Load_Staff_Rec

End Sub

Private Sub Load_Staff_Rec()

    Set db_RS(14) = New ADODB.RecordSet
    
    db_SQL = "SELECT " _
            & "STAFF_ID, FIRST_NAME, LAST_NAME, ADDRESS_LINE_1, " _
            & "ADDRESS_LINE_2, CITY, PIN_CODE, PHONE_NUMBER, " _
            & "CELL_NUMBER, DATE_HIRE, DATE_LEAVE, SALARY, " _
            & "EMAIL, MISC_INFO, ENABLED, JOB_TYPE " _
            & "FROM cm_staff ORDER BY staff_id"

    objDB.Query db_SQL, db_RS(14)
    
    ReLoad_Staff_Rec
    
End Sub

Private Sub ReLoad_Staff_Rec()

    For i = 0 To 13
        txtInfo(i).text = db_RS(14).fields(i) & ""
    Next i
    
    chkVisible.Value = db_RS(14).fields(14)
    Select_List_Item cmbInfo, db_RS(14).fields(15)
    
    List_Assoc_Courses

End Sub

Private Sub Fill_Combo_Lists()

    With cmbInfo
        .AddItem "Teacher", 0
        .AddItem "Receptionist", 1
        .AddItem "Clerk", 2
        .AddItem "Cashier", 3
        .AddItem "Librarian", 4
    End With

End Sub

Private Sub List_Assoc_Courses()

    lstCourses.Clear
    
    Set db_RS(17) = New ADODB.RecordSet
    
    db_SQL = "SELECT T.course_id, T.subject_id, S.subject_name, " _
            & "C.course_name, C.type, C.class " _
            & "FROM cm_course_subject_teacher T, cm_course C, cm_subject S " _
            & "WHERE T.course_id = C.course_id AND T.subject_id = S.subject_id " _
            & "AND staff_id = " & db_RS(14).fields(0)
    
    objDB.Query db_SQL, db_RS(17)
    
    With db_RS(17)
        If .RecordCount > 0 Then
        
            For i = 0 To (.RecordCount - 1)
    
                strTemp = !course_id & " - " & !course_name & " (" & !Type _
                        & ") - " & !Class & "th - " & !subject_name
                        
                lstCourses.AddItem strTemp
                
                objDB.Move "next", db_RS(17)
                
            Next i

        End If
        strTemp = ""
    End With
    
End Sub

Private Sub btnFirst_Click()
    objDB.Move "first", db_RS(14)
    ReLoad_Staff_Rec
End Sub

Private Sub btnLast_Click()
    objDB.Move "last", db_RS(14)
    ReLoad_Staff_Rec
End Sub

Private Sub btnNext_Click()
    objDB.Move "next", db_RS(14)
    ReLoad_Staff_Rec
End Sub

Private Sub btnPrev_Click()
    objDB.Move "prev", db_RS(14)
    ReLoad_Staff_Rec
End Sub

Private Sub btnRec_Refresh_Click()

    objDB.ReQuery db_RS(14)
    ReLoad_Staff_Rec

End Sub

Private Sub Lock_Controls()

    For i = 0 To 13
        txtInfo(i).Locked = True
    Next i
    
    cmbInfo.Locked = True
    chkVisible.Enabled = False
    
    btnRec_Save.Enabled = False
    btnRec_Delete.Enabled = False

End Sub

Private Sub UnLock_Controls()

    For i = 1 To 13
        txtInfo(i).Locked = False
    Next i
    
    cmbInfo.Locked = False
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

Private Sub btnRec_Add_Click()

    If blnAdd <> True Then
    
        btnRec_Add.Caption = "CANCEL"
        btnRec_Edit.Enabled = False
        
        Lock_Navigation
        UnLock_Controls
        
        For i = 0 To (txtInfo.Count - 1)
            txtInfo(i).text = ""
        Next i
        
        cmbInfo.ListIndex = 0
        chkVisible.Value = 1
        
        lstCourses.Clear
    
        blnAdd = True
        
    Else
    
        btnRec_Add.Caption = "ADD"
        btnRec_Edit.Enabled = True
        
        UnLock_Navigation
        Lock_Controls
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
    
        Set db_RS(16) = New ADODB.RecordSet
        
        db_SQL = "UPDATE cm_staff SET " _
                & "FIRST_NAME       = '" & txtInfo(1) & "' , " _
                & "LAST_NAME        = '" & txtInfo(2) & "' , " _
                & "ADDRESS_LINE_1   = '" & txtInfo(3) & "' , " _
                & "ADDRESS_LINE_2   = '" & txtInfo(4) & "' , " _
                & "CITY             = '" & txtInfo(5) & "' , " _
                & "PIN_CODE         = TO_NUMBER(" & txtInfo(6) & ") , " _
                & "PHONE_NUMBER     = '" & txtInfo(7) & "' , " _
                & "CELL_NUMBER      = '" & txtInfo(8) & "' , " _
                & "DATE_HIRE        = TO_DATE('" & txtInfo(9) & "') , " _
                & "DATE_LEAVE       = TO_DATE('" & txtInfo(10) & "') , " _
                & "JOB_TYPE         = '" & cmbInfo.List(cmbInfo.ListIndex) & "' , " _
                & "SALARY           = TO_NUMBER(" & txtInfo(11) & ") , " _
                & "EMAIL            = '" & txtInfo(12) & "' , " _
                & "MISC_INFO        = '" & MultiLine_To_Single(txtInfo(13)) & "' , " _
                & "ENABLED          = '" & chkVisible.Value & "' " _
                & "WHERE STAFF_ID   = " & txtInfo(0)
                
        objDB.Execute db_SQL, sngTemp

        frm_test3.Show
        frm_test3.fill (db_SQL & " . And sngTemp is = " & sngTemp)
        
        If sngTemp > 0 Then
            lblResult.Caption = "Data edited / updated accordingly."
        Else
            lblResult.Caption = "Data could NOT be edited / updated."
        End If
    
    ElseIf blnAdd = True Then
    
        Set db_RS(16) = New ADODB.RecordSet
        
        db_SQL = "INSERT INTO cm_staff( STAFF_ID, FIRST_NAME, LAST_NAME, " _
                & "ADDRESS_LINE_1, ADDRESS_LINE_2, CITY, PIN_CODE, " _
                & "PHONE_NUMBER, CELL_NUMBER, DATE_HIRE, DATE_LEAVE, " _
                & "JOB_TYPE, SALARY, EMAIL, MISC_INFO, ENABLED ) VALUES ( " _
                    & "seq_staffid.nextval , " _
                    & "'" & txtInfo(1) & "' , " _
                    & "'" & txtInfo(2) & "' , " _
                    & "'" & txtInfo(3) & "' , " _
                    & "'" & txtInfo(4) & "' , " _
                    & "'" & txtInfo(5) & "' , " _
                    & "TO_NUMBER(" & txtInfo(6) & ") , " _
                    & "'" & txtInfo(7) & "' , " _
                    & "'" & txtInfo(8) & "' , " _
                    & "TO_DATE('" & txtInfo(9) & "') , " _
                    & "TO_DATE('" & txtInfo(10) & "') , " _
                    & "'" & cmbInfo.List(cmbInfo.ListIndex) & "' , " _
                    & "TO_NUMBER(" & txtInfo(11) & ") , " _
                    & "'" & txtInfo(12) & "' , " _
                    & "'" & MultiLine_To_Single(txtInfo(13)) & "' , " _
                    & "'" & chkVisible.Value & "' " _
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
        Case 1, 2, 3, 4, 5: 'Not null
        
            If Not Len(data) > 0 Then
                Cancel = True
            End If
            
        Case 6, 11: 'Pin Code & Salary (Should be numeric)
        
            If Not IsNumeric(data) Then
                Cancel = True
            End If
            
        Case 7: 'Home Phone (Format : 0755-4276675)
            
            Dim tmp1, tmp2, tmp3 As String
            
            tmp1 = Left$(data, 4)
            tmp2 = Mid(data, 5, 1)
            tmp3 = Right$(data, 7)
            
            If (Len(tmp1) <> 4 Or Not IsNumeric(tmp1) Or tmp2 <> "-" _
            Or Len(tmp3) <> 7 Or Not IsNumeric(tmp3)) And data <> "" Then
                Cancel = True
            End If
    
        Case 8: 'Cell Number (Should be numeric too)
        
            If Not IsNumeric(data) And data <> "" Then
                Cancel = True
            End If

        Case 9: 'Not Null + Date format
        
            If Not IsDate(data) Or Not Len(data) > 0 Then
                Cancel = True
            End If
            
        Case 10: 'Date format
            
            If Not IsDate(data) And data <> "" Then
                Cancel = True
            End If
    
    End Select

End Sub

Private Function Check_Form_Fields()

    Dim fields(), i As Variant
    fields() = Array(1, 2, 3, 4, 5, 6, 9, 11)

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

    Set db_RS(15) = New ADODB.RecordSet
    
    db_SQL = "SELECT column_name FROM user_tab_cols WHERE table_name = 'CM_STAFF'"
    objDB.Query db_SQL, db_RS(15)
    
    For i = 0 To (db_RS(15).RecordCount - 1)
    
        cmbSearch_Field.AddItem LCase(db_RS(15).fields(0))
        objDB.Move "next", db_RS(15)
        
    Next i
    
    cmbSearch_Type.AddItem "Exact", 0
    cmbSearch_Type.AddItem "Like", 1
    
    cmbSearch_Start.AddItem "Current Record", 0
    cmbSearch_Start.AddItem "Beginning", 1
    
    cmbSearch_Field.ListIndex = 8
    cmbSearch_Type.ListIndex = 1
    cmbSearch_Start.ListIndex = 1

End Sub

Private Sub btnSearch_Submit_Click()

    Search_Recordset Me, db_RS(14)
    ReLoad_Staff_Rec

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
        
        Select_List_Item cmbSearch_Field, "staff_id"
        
        cmbSearch_Type.ListIndex = 0
        cmbSearch_Start.ListIndex = 1

        btnSearch_Submit_Click
        
        Me.SetFocus
        
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frm_staff = Nothing
End Sub
