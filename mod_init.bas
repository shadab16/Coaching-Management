Attribute VB_Name = "mod_init"
'############## GLOBAL / LOCAL VARIABLES, CONSTANTS & OBJECTS ############

' Database Constants
Private Const DB_Type As String = "oracle"
Private Const DB_Version As String = "10g"
Private Const DB_DSN As String = "OracleXE"
Private Const DB_User As String = "test"
Private Const DB_Pass As String = "qwerty"

' User Authorization Constants
Public Const Auth_Require As Boolean = False ' True / False
Public Const Auth_User As String = "Test"
Public Const Auth_Pass As String = "142136154141139134"

' Splash Page Options
Public Const Splash_Display As Boolean = False ' True / False
Public Const Splash_Time As Integer = 2000

' Image directories
Public Dir_Icon, Dir_Pics, Dir_Backup, Dir_Report As String

' Declaring the application critical objects.
Public objDB As cls_database
Public objFSO As FileSystemObject
Public objDialog As Object

Public db_RS(1 To 100) As ADODB.RecordSet
Public db_SQL As String

'##################### STARTUP / USER AUTHORIZATION  #####################

Public Sub Main()

    Call Init

    If Splash_Display = True Then
        frm_splash.Show
    ElseIf Auth_Require = True Then
        frm_auth_user.Show
    Else
        frm_main_mdi.Show
    End If

End Sub

'################## APPLICATION INITIALIZING PROCEDURE  ##################

Public Sub Init()

    Set objDB = New cls_database
    Set objFSO = New FileSystemObject
    
    Set objDialog = frm_main_mdi.Common_Dialog

    Dir_Icon = App.Path & "\icons\"
    Dir_Pics = App.Path & "\pictures\"
    Dir_Backup = App.Path & "\pictures\backup\"
    Dir_Report = App.Path & "\reports\"

    With objDB
        .DbType = DB_Type
        .DbVersion = DB_Version
        .DataSource = DB_DSN
        .UserName = DB_User
        .Password = DB_Pass
    
        .Connect
    End With

End Sub
