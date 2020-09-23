Attribute VB_Name = "DEModule"
'Test to find out if we have a valid dataenvironment and command
Public Function IsRecordset(ByVal Name As String) As Boolean
Dim i As Long
Dim DataEnvName As String
Dim CommandName As String
Dim s() As String
Dim IsDataEnvironment As Boolean
Dim NrDE As Long
Dim NrCom As Long

'Ex. Name = "DataEnvironment1.rsCommand1"
s = Split(Name, ".")
'S(0) = "DataEnvironment1"
'S(1) = "rsCommand1"


If UBound(s) <> 1 Then
    IsRecordset = False
    Exit Function
End If


If Len(s(1)) <= 2 Then
    IsRecordset = False
    Exit Function
End If

If Left$(Trim$(s(1)), 2) = "rs" Then
    CommandName = Right$(s(1), Len(s(1)) - 2)

    For i = 1 To VBInstance.ActiveVBProject.VBComponents.Count
        If VBInstance.ActiveVBProject.VBComponents(i).Name = s(0) Then
            NrDE = i
            IsDataEnvironment = True
            
            Exit For
        End If
    Next i
    
    If IsDataEnvironment Then
        If Right$(VBInstance.ActiveVBProject.VBComponents(NrDE).DesignerID, 16) = ".DataEnvironment" Then

            For j = 1 To VBInstance.ActiveVBProject.VBComponents(NrDE).Designer.DECommands.Count
                If VBInstance.ActiveVBProject.VBComponents(NrDE).Designer.DECommands(j).Name = CommandName Then
                    NrCom = j
                    IsRecordset = True
                    Exit Function
                End If
            Next j
        End If
    Else
        IsRecordset = False
        Exit Function
    End If
Else
    IsRecordset = False
    Exit Function
End If

End Function

'Extract Fields List from the dataenvironment command
Public Function FieldsList(ByVal Name As String) As String
Dim i As Long
Dim DataEnvName As String
Dim CommandName As String
Dim s() As String
Dim IsDataEnvironment As Boolean
Dim ReturnList As String
Dim NrDE As Long
Dim NrCom As Long

'Ex. Name = "DataEnvironment1.rsCommand1"
s = Split(Name, ".")
'S(0) = "DataEnvironment1"
'S(1) = "rsCommand1"

ReturnList = ""

If UBound(s) <> 1 Then
    FieldsList = ""
    Exit Function
End If



If Len(s(1)) <= 2 Then
    FieldsList = ""
    Exit Function
End If


If Left$(s(1), 2) = "rs" Then

    CommandName = Right$(s(1), Len(s(1)) - 2)
    'Ex. CommandName = "Command1"
    
    For i = 1 To VBInstance.ActiveVBProject.VBComponents.Count
        If VBInstance.ActiveVBProject.VBComponents(i).Name = s(0) Then
            NrDE = i
            IsDataEnvironment = True
            Exit For
        End If
    Next i
    
    If IsDataEnvironment Then
        If Right$(VBInstance.ActiveVBProject.VBComponents(NrDE).DesignerID, 16) = ".DataEnvironment" Then
            For j = 1 To VBInstance.ActiveVBProject.VBComponents(NrDE).Designer.DECommands.Count
                If VBInstance.ActiveVBProject.VBComponents(NrDE).Designer.DECommands(j).Name = CommandName Then
                   For k = 1 To VBInstance.ActiveVBProject.VBComponents(NrDE).Designer.DECommands(j).defields.Count
                        ReturnList = ReturnList & VBInstance.ActiveVBProject.VBComponents(NrDE).Designer.DECommands(j).defields(k).Name & "|"
                   Next k
                   ReturnList = Left$(ReturnList, Len(ReturnList) - 1)
                   FieldsList = ReturnList
                   Exit Function
                End If
            Next j
        End If
    Else
        FieldsList = ""
        Exit Function
    End If
Else
    FieldsList = ""
    Exit Function
End If

End Function

Public Function GetEnumValuesForCodePane(cp As CodePane) As String
    Dim cm As CodeModule
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim sSwitchVar As String
    Dim sSwitchVarType As String
    Dim i As Long
    
   On Error GoTo GetEnumValuesForCodePane_Error

    cp.GetSelection y1, x1, y2, x2

    If x1 <> x2 Then Exit Function
    If y1 <> y2 Then Exit Function

    If x1 = 0 Then Exit Function
    
    Set cm = cp.CodeModule
    
    If ((Right$(Trim$(cm.Lines(y1, 1)), 8)) = (".Fields(")) Or ((Right$(Trim$(cm.Lines(y1, 1)), 9)) = (".Fields (")) Or ((Right$(Trim$(cm.Lines(y1, 1)), 10)) = (".Fields (" & Chr(34))) Or ((Right$(Trim$(cm.Lines(y1, 1)), 9)) = (".Fields(" & Chr(34))) Or (UCase$(Right$(Trim$(cm.Lines(y1, 1)), 1)) = "!") Then
        If UCase$(Right$(Trim$(cm.Lines(y1, 1)), 1)) = "!" Then
            i = Len(cm.Lines(y1, 1))
            While Mid$(cm.Lines(y1, 1), i, 1) <> " " And Mid$(cm.Lines(y1, 1), i, 1) <> "=" And i >= 0
                i = i - 1
            Wend
            If IsRecordset(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 1))) Then
                GetEnumValuesForCodePane = FieldsList(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 1)))
            End If
        ElseIf ((Right$(Trim$(cm.Lines(y1, 1)), 9)) = (".Fields(" & Chr(34))) Then
            i = Len(cm.Lines(y1, 1))
            While Mid$(cm.Lines(y1, 1), i, 1) <> " " And i >= 0
                i = i - 1
            Wend
            If IsRecordset(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 9))) Then
                GetEnumValuesForCodePane = FieldsList(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 9)))
            End If
        ElseIf ((Right$(Trim$(cm.Lines(y1, 1)), 10)) = (".Fields (" & Chr(34))) Then
            i = Len(cm.Lines(y1, 1)) - 3
            While Mid$(cm.Lines(y1, 1), i, 1) <> " " And i >= 0
                i = i - 1
            Wend
            If IsRecordset(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 10))) Then
                GetEnumValuesForCodePane = FieldsList(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 10)))
            End If
        ElseIf ((Right$(Trim$(cm.Lines(y1, 1)), 9)) = (".Fields (")) Then
            i = Len(cm.Lines(y1, 1)) - 2
            While Mid$(cm.Lines(y1, 1), i, 1) <> " " And i >= 0
                i = i - 1
            Wend
            If IsRecordset(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 9))) Then
                GetEnumValuesForCodePane = FieldsList(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 9)))
            End If
        Else
            i = Len(cm.Lines(y1, 1)) - 1
            While Mid$(cm.Lines(y1, 1), i, 1) <> " " And i >= 0
                i = i - 1
            Wend
            If IsRecordset(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 8))) Then
                GetEnumValuesForCodePane = FieldsList(Trim$(Left$(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i), Len(Right$(cm.Lines(y1, 1), Len(cm.Lines(y1, 1)) - i)) - 8)))
            End If
        End If
    Else
        Exit Function
    End If
    

   On Error GoTo 0
   Exit Function

GetEnumValuesForCodePane_Error:

End Function
