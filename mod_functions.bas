Attribute VB_Name = "mod_functions"
Public Function Get_Form_by_Tag(ByVal TagName As String)
   
    For i = 0 To (Forms.Count - 1)
        If Forms(i).Tag = TagName Then
            Get_Form_by_Tag = i
            Exit Function
        End If
    Next i

    Get_Form_by_Tag = False

End Function

Public Function Get_Total_Forms() As Byte

    Dim i As Byte, Frm As Form
    i = 0
    
    For Each Frm In Forms
        If Frm.Name <> "frm_main_mdi" Then
            If Frm.MDIChild = True Then
                i = i + 1
            End If
        End If
    Next Frm

    Get_Total_Forms = i
    
End Function

Public Function String_Encode(ByVal text As String) As String

    Dim Encoded As String
    
    For i = 1 To Len(text)
        Encoded = Encoded & Format$((255 - Asc(Mid(text, i, 1))), "000")
    Next i
    
    String_Encode = Encoded

End Function

Public Function String_Decode(ByVal text As String) As String

    Dim Decoded As String
    
    For i = 1 To Len(text) Step 3
        Decoded = Decoded & Chr(255 - Mid(text, i, 3))
    Next i
    
    String_Decode = Decoded

End Function

Public Function File_Exists(ByVal FileName As String, _
                Optional ByVal Search As String = "rel") As Boolean

    Select Case Search
        Case "rel": FileName = App.Path & "\" & Trim(FileName)
        Case "abs": FileName = Trim(FileName)
        Case Else: Exit Function
    End Select
        
    If Dir(FileName) <> "" Then
        File_Exists = True
    Else
        File_Exists = False
    End If

End Function

Public Function MultiLine_To_Single(ByVal Multi As String) As String

    Replace Multi, Chr(13), " "
    MultiLine_To_Single = Multi

End Function

Public Function Get_Identifier(ByVal text As String) As Integer

    Dim words() As String
    words() = Split(text, " ")
    
    Get_Identifier = 0
    
    If IsNumeric(words(0)) Then
        Get_Identifier = Int(words(0))
    End If

End Function

Public Sub Resize(ByRef x1, ByRef y1, ByVal x2 As Double, ByVal y2 As Double)

    Dim Ratio
    Ratio = x1 / y1

    If (x1 > x2) Or (y1 > y2) Then
    
        Diff_X = x1 / x2
        Diff_Y = y1 / y2
        
        If Diff_X >= Diff_Y Then

            x1 = x2
            y1 = x1 / Ratio
        
        ElseIf Diff_X < Diff_Y Then
        
            y1 = y2
            x1 = y1 * Ratio
        
        End If
    
    End If

End Sub

Public Sub Select_List_Item(Ctrl As Control, text As String, _
                            Optional ByID As Boolean = False)

    If Not (TypeOf Ctrl Is ListBox Or TypeOf Ctrl Is ComboBox) Then
        Exit Sub
    End If
    
    With Ctrl
    
        If Not .ListCount > 0 Then: Exit Sub
        
        For i = 0 To (.ListCount - 1)
    
            If ByID = False Then
                
                If .List(i) = text Then
                    .ListIndex = i: Exit Sub
                End If
                
            ElseIf ByID = True Then
            
                If Get_Identifier(.List(i)) = text Then
                    .ListIndex = i: Exit Sub
                End If
    
            End If
    
        Next i
    
    End With

End Sub

Public Sub Search_Recordset(ByRef Frm As Form, ByRef RS As ADODB.RecordSet)

    On Error GoTo errSearch

    Dim strField, strType, strText, strStart, Criteria As String
    Dim RS_Bookmark, Skip

    With Frm
        
        strField = .cmbSearch_Field.text
        strType = .cmbSearch_Type.text
        strText = Trim(.txtSearch_Input.text)
        strStart = .cmbSearch_Start.text
        
    End With
        
    If strText = "" Then
    
        MsgBox ("Fill in the DATA to search FOR")
        Exit Sub
        
    ElseIf strType = "" Then
    
        MsgBox ("Select the search TYPE to perform, from the dropdown menu")
        Exit Sub
        
    ElseIf strField = "" Then
    
        MsgBox ("Select the form FIELD to search IN")
        Exit Sub
        
    End If
    
    If strType = "Exact" Then
        Criteria = strField & " = '" & strText & "'"
    ElseIf strType = "Like" Then
        Criteria = strField & " LIKE '%" & strText & "%'"
    End If
    
    If strStart = "Current Record" Then
    
        strStart = adBookmarkCurrent
        Skip = 1
        
    Else
    
        strStart = adBookmarkFirst
        Skip = 0
        
    End If
    
    RS_Bookmark = RS.Bookmark

    RS.Find Criteria, Skip, adSearchForward, strStart

    If RS.EOF = True Or RS.BOF = True Then
        RS.Bookmark = RS_Bookmark
    End If

    RS_Bookmark = Null
    
    Exit Sub
    
errSearch:
    Frm.lblResult.Caption = "Search could not be completed!"
    
End Sub

Public Sub Clear_Immediate_Window()

    For i = 1 To 200
        Debug.Print vbCrLf
    Next i

End Sub
