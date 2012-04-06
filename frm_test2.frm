VERSION 5.00
Begin VB.Form frm_test2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5145
   Begin VB.CommandButton btnWrite 
      Caption         =   "WRITE to FILE"
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
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label lblResult 
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
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
End
Attribute VB_Name = "frm_test2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnWrite_Click()

    Dim objFile As TextStream
    Dim strName, strPath, strText As String
    Dim i As Integer
    
    strName = Replace(Format$(Date, "YYYY-MM-DD"), "/", "-") & "-" & Int(Timer()) & ".txt"
    strPath = Dir_Report & strName
       
    Set objFile = objFSO.OpenTextFile(strPath, ForWriting, True, TristateTrue)

    Dim RS As New ADODB.RecordSet
    objDB.Query "SELECT RPAD(subject_id, 10) AS id, " _
            & "RPAD(subject_name, 25) AS name FROM cm_subject " _
            & "ORDER BY subject_id", RS
    
    With objFile
    
        If RS.RecordCount > 0 Then
        
            .WriteLine ("Subject ID" & vbTab & "Subject Name")
            .WriteLine (String(10, "-") & vbTab & String(25, "-"))
        
            For i = 1 To RS.RecordCount
            
                strText = RS!ID & vbTab & RS!Name
                .WriteLine (strText)
            
                objDB.Move "next", RS
                
            Next i
    
        End If
        
        RS.Close
        objDB.Query "SELECT RPAD(student_id, 10) AS id, RPAD(first_name, 15) AS fname," _
                    & " RPAD(last_name, 15) AS lname FROM cm_student" _
                    & " ORDER BY student_id", RS
        
        If RS.RecordCount > 0 Then
        
            .WriteBlankLines (3)
            .WriteLine ("Student ID" & vbTab & "First Name" & vbTab & vbTab & "Last Name")
            .WriteLine (String(10, "-") & vbTab & String(15, "-") & vbTab & String(15, "-"))
        
            For i = 1 To RS.RecordCount
            
                strText = RS!ID & vbTab & RS!fname & vbTab & RS!lname
                .WriteLine (strText)
            
                objDB.Move "next", RS
                
            Next i
    
        End If
        
        .Close
    
    End With

    lblResult.Caption = strName & " written successfully."
    
End Sub
