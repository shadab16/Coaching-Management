VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###################### VARIABLE DECLARATION #############################

' Variables holding the connectivity information
Private dbConfig_DbType As String
Private dbConfig_DbVersion As String
Private dbConfig_DbName As String 'Optional for Oracle
Private dbConfig_DataSource As String
Private dbConfig_Host As String 'Optional for Oracle

' The database username and associated password
Private dbConfig_UserName As String
Private dbConfig_Password As String

' Objects used for actual interaction with the DB (but not always)
Private dbAdo_Conn As ADODB.Connection
Private dbAdo_Cmd As ADODB.Command

' Miscellaneous variables
Private db_SQL As String
Private db_Connected As Boolean

'###################### CONSTRUCTOR & DESTRUCTOR #########################

' On object initialization, set the default db type.
Private Sub Class_Initialize()
    Me.DbType = "oracle"
End Sub

'######################### CLASS PROPERTIES ##############################

' Get / Set property : Database Type
Public Property Let DbType(ByVal data As String)

    dbConfig_DbType = LCase(Trim(data))
    
    ' Set some of the defaults
    If dbConfig_DbType = "oracle" Then
        dbConfig_DbVersion = "10g"
        dbConfig_DataSource = "OracleXE"
    End If
    
End Property

Public Property Get DbType() As String
    DbType = dbConfig_DbType
End Property

' Get / Set property : Database Version
Public Property Let DbVersion(ByVal data As String)

    dbConfig_DbVersion = LCase(Trim(data))
    
    ' Set the default property
    If dbConfig_DbVersion = "10g" Then
        dbConfig_DataSource = "OracleXE"
    End If
    
End Property

Public Property Get DbVersion() As String
    DbVersion = dbConfig_DbVersion
End Property

' Get / Set property : Data Source
Public Property Let DataSource(ByVal data As String)
    dbConfig_DataSource = Trim(data)
End Property

Public Property Get DataSource() As String
    DataSource = dbConfig_DataSource
End Property

' Get / Set property : Host
Public Property Let Host(ByVal data As String)
    dbConfig_Host = data
End Property

Public Property Get Host() As String
    Host = dbConfig_Host
End Property

' Get / Set property : Username
Public Property Let UserName(ByVal data As String)
    dbConfig_UserName = Trim(data)
End Property

Public Property Get UserName() As String
    UserName = dbConfig_UserName
End Property

' Get / Set property : Password
Public Property Let Password(ByVal data As String)
    dbConfig_Password = data
End Property

Public Property Get Password() As String
    Password = dbConfig_Password
End Property
    
'############################# METHODS ###################################

' Establish the DB connection
Public Sub Connect()

    If db_Connected = True Then
        Exit Sub
    End If

    Set dbAdo_Conn = New ADODB.Connection
    Set dbAdo_Cmd = New ADODB.Command
    
    If dbConfig_DbType = "oracle" Then
    
        If dbConfig_DbVersion = "10g" Then
            
            dbAdo_Conn.ConnectionString = "Data Source=" & dbConfig_DataSource _
                                & ";User Id=" & dbConfig_UserName _
                                & ";Password=" & dbConfig_Password & ";"
        
        ElseIf dbConfig_DbVersion = "9i" Then
        
            dbAdo_Conn.ConnectionString = "Provider=" & dbConfig_DataSource _
                                & ";Password=" & dbConfig_Password _
                                & ";User ID=" & dbConfig_UserName _
                                & ";Persist Security Info=True"
                                
        Else
        
            ' ------- Throw error as we don't don't support this DB version -------
        
        End If

    ElseIf dbConfig_DbType = "mysql" Then

        dbAdo_Conn.ConnectionString = "DRIVER={" & dbConfig_DataSource _
                    & "};SERVER=" & dbConfig_Host _
                    & ";DATABASE=" & dbConfig_DbName _
                    & ";UID=" & dbConfig_UserName _
                    & ";PWD=" & dbConfig_Password
    Else
    
        ' ------------ Throw error because the DB type is unsupported --------------
        
    End If
    
    dbAdo_Conn.Open
    db_Connected = True
    
    dbAdo_Cmd.ActiveConnection = dbAdo_Conn

End Sub

' Disconnect from the database
Public Sub Disconnect()

    If db_Connected = False Then
        Exit Sub
    End If

    For Each Frm In Forms
        If Frm.Name <> "frm_main_mdi" Then
            Unload Frm
        End If
    Next Frm

    For i = 1 To 100
        Set db_RS(i) = Nothing
    Next i
    
    dbAdo_Conn.Close

    db_Connected = False

End Sub

Public Sub Query(ByVal QueryString As String, _
                ByRef RecordSet As ADODB.RecordSet)

    db_SQL = Trim(QueryString)
    RecordSet.Open db_SQL, dbAdo_Conn, adOpenStatic, adLockPessimistic

End Sub

Public Sub ReQuery(ByRef RecordSet As ADODB.RecordSet)

    Dim RS_Bookmark
    RS_Bookmark = RecordSet.Bookmark

    objDB.Commit
    RecordSet.ReQuery
    
    RecordSet.Bookmark = RS_Bookmark
    RS_Bookmark = Null

End Sub

Public Sub Execute(ByVal QueryString As String, Optional ByRef Count, _
                    Optional ByRef RecordSet As ADODB.RecordSet)

    db_SQL = Trim(QueryString)
    
    If RecordSet Is Nothing Then
    
        If IsMissing(Count) Then
            dbAdo_Conn.Execute db_SQL
        Else
            dbAdo_Conn.Execute db_SQL, Count
        End If
    
    Else
    
        If IsMissing(Count) Then
            Set RecordSet = dbAdo_Conn.Execute(db_SQL)
        Else
            Set RecordSet = dbAdo_Conn.Execute(db_SQL, Count)
        End If

    End If

    Me.Commit

End Sub

Public Sub Exec_Proc_Out(ByVal ProcName As String, ByRef ParamOUT As String, _
        Optional ByRef ParamOUT1 As String, Optional ByRef ParamOUT2 As String)

    Dim objParam As New ADODB.Parameter, strTemp As String

    dbAdo_Cmd.CommandType = adCmdStoredProc
    dbAdo_Cmd.CommandText = ProcName

    Set objParam = dbAdo_Cmd.CreateParameter("Param1", adVarChar, adParamOutput, 32, strTemp)
    dbAdo_Cmd.Parameters.Append objParam

    If ParamOUT1 <> "" Then
    
        Set objParam = dbAdo_Cmd.CreateParameter("Param2", adVarChar, adParamOutput, 32, strTemp)
        dbAdo_Cmd.Parameters.Append objParam

    End If

    If ParamOUT2 <> "" Then
    
        Set objParam = dbAdo_Cmd.CreateParameter("Param3", adVarChar, adParamOutput, 32, strTemp)
        dbAdo_Cmd.Parameters.Append objParam

    End If

    dbAdo_Cmd.Execute

    ParamOUT = dbAdo_Cmd.Parameters("Param1").Value
    
    If ParamOUT1 <> "" Then: ParamOUT1 = dbAdo_Cmd.Parameters("Param2").Value
    If ParamOUT2 <> "" Then: ParamOUT2 = dbAdo_Cmd.Parameters("Param3").Value

End Sub

Public Sub Commit()
    
    With dbAdo_Cmd
        .CommandText = "Commit"
        .CommandType = adCmdText
        .Execute
        .CommandText = ""
    End With

End Sub

Public Sub Move(ByVal MoveType As String, ByRef RS As ADODB.RecordSet)

    Select Case MoveType
    
        Case "first":
            RS.MoveFirst
            
        Case "last":
            RS.MoveLast

        Case "next":
            RS.MoveNext
            If RS.EOF = True Then
                RS.MoveFirst
            End If

        Case "prev":
            RS.MovePrevious
            If RS.BOF = True Then
                RS.MoveLast
            End If
            
    End Select

End Sub
