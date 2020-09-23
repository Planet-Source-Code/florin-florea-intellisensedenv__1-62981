Attribute VB_Name = "Module1"
Public Type POINTAPI
    X As Long
    Y As Long
End Type

' Windows Hook Types
Private Const WH_CBT As Long = 5
Private Const WH_KEYBOARD As Long = 2
' Generic hook code
Private Const HC_ACTION As Long = 0
' CBT action code constants
Private Const HCBT_SETFOCUS As Long = 9
' SetWindowPos() constants
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOSIZE As Long = &H1
' RedrawWindow() constants
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_UPDATENOW As Long = &H100
' Function declares

' Function declares
Public Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32.dll" () As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCaretPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long

Public VBInstance             As VBIDE.VBE


Public pt As POINTAPI
Public m_bRunning As Boolean

' Our hook handles
Private hHookCBT As Long, hHookKyBd As Long
' Handle to the active window, if it is a code pane window (0 if it's not)
Private hWndActiveCodePane As Long
' Row and Col positions when the listbox was shown
Private PaneCol As Long, PaneRow As Long

Public Property Let Running(ByVal bRunning As Boolean)
    If bRunning = m_bRunning Then Exit Property
    If bRunning Then
        ' Set the hooks, fail if they can't be set
        hHookCBT = SetWindowsHookEx(WH_CBT, AddressOf CBTProc, 0, App.ThreadID)
        If hHookCBT = 0 Then GoTo FAIL_AT_0
        
        hHookKyBd = SetWindowsHookEx(WH_KEYBOARD, AddressOf KyBdProc, 0, App.ThreadID)
        If hHookKyBd = 0 Then GoTo FAIL_AT_1
                
        m_bRunning = True
        Exit Property
    End If
    
    ' Unhook
    UnhookWindowsHookEx hHookKyBd
    hHookKyBd = 0
FAIL_AT_1:
    UnhookWindowsHookEx hHookCBT
    hHookCBT = 0
FAIL_AT_0:
    m_bRunning = False
End Property

Public Property Get Running() As Boolean
    Running = m_bRunning
End Property

Private Function CBTProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim sWindowClass As String, nLen As Long
    Select Case nCode
    Case HCBT_SETFOCUS
        ' wParam : hWnd of the window gaining the keyboard focus
        ' lParam : hWnd of the window losing the keyboard focus
        
        If wParam = Form1.hwnd Then
            SetFocus hWndActiveCodePane
            CBTProc = 1
            Exit Function
            ' Done differently, with a timer [try it and see]
            GoTo END_OF_FUNCTION
        End If
        If hWndActiveCodePane = lParam And hWndActiveCodePane <> 0 Then
            ' The code pane is losing focus
            ' We need to hide our popup window if it's visible
            Form1.Visible = False
        End If
        ' Now check if the window gaining the focus is a code window
        ' First, we assume it isn't
        hWndActiveCodePane = 0
        ' Now get the class name of the window
        sWindowClass = Space$(64)
        nLen = GetClassName(wParam, sWindowClass, 64)
        If nLen > 0 Then
            ' Function succeeded, let's adjust the string
            sWindowClass = Left$(sWindowClass, nLen)
            ' Now test against the class for code windows
            If sWindowClass = "VbaWindow" Then
                ' It is - remember the window
                hWndActiveCodePane = wParam
            End If
        End If
    End Select
    ' Keep the chain going
END_OF_FUNCTION:
    CBTProc = CallNextHookEx(hHookCBT, nCode, wParam, lParam)
End Function

Private Function KyBdProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If hWndActiveCodePane = 0 Then GoTo DO_NOT_PROCESS

    Dim cp As CodePane, cm As CodeModule
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim i As Long, j As Long
    Dim sLine As String, sWord As String
    Dim arrStrings() As String
    Dim sEnumValues As String
    Dim pt As POINTAPI
    Dim SS() As String
    
    If Form1.Visible = False Then
        'frmAddIn.Caption = wParam
        If wParam = 49 Or wParam = 222 Or wParam = 57 Then
            
            Set cp = VBInstance.ActiveCodePane
            cp.GetSelection y1, x1, y2, x2
            
            sEnumValues = GetEnumValuesForCodePane(cp)
            If sEnumValues = "" Then GoTo DO_NOT_PROCESS
            
            Form1.List1.Clear
            SS = Split(sEnumValues, "|")
            For i = 1 To UBound(SS)
                Form1.List1.AddItem SS(i)
            Next i
            
            PaneCol = x1
            PaneRow = y1
            
            GetCaretPos pt
            ClientToScreen hWndActiveCodePane, pt
            
            SetWindowPos Form1.hwnd, HWND_TOPMOST, pt.X - 18, pt.Y + 12, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOSIZE
            'Form1.Visible = True
            RedrawWindow hWndActiveCodePane, ByVal 0&, 0, RDW_INVALIDATE Or RDW_UPDATENOW

            
        End If
    Else
        Select Case wParam
        Case vbKeyUp
            If lParam > 0 Then
                Form1.HandleKeyUp
            End If
            KyBdProc = 1
            Exit Function
        Case vbKeyDown
            If lParam > 0 Then
                Form1.HandleKeyDown
            End If
            KyBdProc = 1
            Exit Function
        Case vbKeyEscape
            Form1.Visible = False
            KyBdProc = 1
            Exit Function
        Case vbKeySpace, vbKeyReturn, 190, 187
            ' If the key is being pressed then insert the text
            If lParam > 0 Then
                ReplaceCurrentWord Form1.GetSelectedText
                Form1.Visible = False
            End If
            ' If the user pressed control at the same time,
            ' hide the keystroke from VB
            If GetAsyncKeyState(vbKeyControl) < 0 Then
                KyBdProc = 1
                Exit Function
            End If
            If wParam <> 190 And wParam <> 187 Then
                wParam = 0
                lParam = 0
                KyBdProc = 1
                Exit Function
            End If
        Case Else
'            ' If the key is being pressed...
''            If wParam <> 222 And wParam <> 16 Then '222=" 16=RShift
''                MsgBox wParam 'Show the pressed key code
''            End If
'            If lParam < 0 Then
'                ' Get the currently typed word
'                Set cp = VBInstance.ActiveCodePane
'                If cp Is Nothing Then GoTo DO_NOT_PROCESS
'                cp.GetSelection y1, x1, y2, x2
'                If x1 > 1 Then
'                    Set cm = cp.CodeModule
'                    sLine = cm.Lines(y1, 1)
'                    For i = x1 - 1 To 1 Step -1
'                        Select Case CLng(AscW(Mid$(sLine, i, 1)))
'                        Case 32&, 40&, 41&, 44&, 43&, 45&, 42&, 47&, 92&
'                            sWord = Mid$(sLine, i + 1, x1 - i - 1)
'                            Exit For
'                        End Select
'                    Next
'                    If sWord = "" Then sWord = Left$(sLine, x1 - 1)
'                    ' Tell the listbox to search for that word
'                    Form1.SetSearchWord sWord
'                End If
'            End If
        End Select
    End If
    
DO_NOT_PROCESS:
KyBdProc = CallNextHookEx(hHookKyBd, nCode, wParam, lParam)

End Function

Public Function CheckListBoxNeeded()
    ' This function checks if it's time to hide the list box
    ' it is called through a timer on the list box form
    If Form1.Visible = False Then Exit Function
    
    If GetFocus = 0 Then
        ' This means another process has the keyboard focus
        ' We'll hide the listbox
        Form1.Visible = False
    End If
    
    Dim cp As CodePane, cm As CodeModule, x1 As Long, y1 As Long, x2 As Long, y2 As Long, i As Long, sLine As String
    ' Retrieve the active code pane
    Set cp = VBInstance.ActiveCodePane
    ' No active code pane => no list box [this should be handled in the CBTProc anyway, but you can't be too careful]
    If cp Is Nothing Then Form1.Visible = False: Exit Function
    ' Get the associated code module
    Set cm = cp.CodeModule
    ' Retrieve the selection
    cp.GetSelection y1, x1, y2, x2
    ' If we're not on the same line then we must hide the list box
    If y1 <> PaneRow Then Form1.Visible = False: Exit Function
    ' If we're on the same column then keep it
    If x1 = PaneCol Then Exit Function
    ' Retrieve the line
    sLine = cm.Lines(y1, 1)
    ' Empty line => no list box
    If sLine = "" Then Form1.Visible = False: Exit Function
    

    For i = PaneCol To x1 Step Sgn(x1 - PaneCol)
        If i <= Len(sLine) Then
            Select Case CLng(AscW(Mid$(sLine, i, 1)))
            Case 32&, 40&, 41&, 44&, 43&, 45&, 42&, 47&, 92&
                Form1.Visible = False
                Exit Function
            End Select
        End If
    Next
End Function

Public Sub SetFocusToCodeWindow()
    ' Stop the listbox from retaining the focus
    If hWndActiveCodePane <> 0 Then
        If Form1.Visible Then
            SetFocus hWndActiveCodePane
        End If
    End If
End Sub

Public Sub ReplaceCurrentWord(sWord As String)
    Dim cp As CodePane, cm As CodeModule, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim sLine As String, i As Long
    Set cp = VBInstance.ActiveCodePane
    Set cm = cp.CodeModule
    cp.GetSelection y1, x1, y2, x2
    sLine = cm.Lines(y1, 1)
    i = Len(sLine)
    Select Case Right$(sLine, 1)
    Case "!"
        cm.ReplaceLine y1, sLine & sWord
        cp.SetSelection y1, i + Len(sWord) + 1, y1, i + Len(sWord) + 1
    Case Chr(34)
        cm.ReplaceLine y1, sLine & sWord & Chr(34) & ")"
        cp.SetSelection y1, i + Len(sWord) + 4, y1, i + Len(sWord) + 4
    Case "("
        cm.ReplaceLine y1, sLine & Chr(34) & sWord & Chr(34) & ")"
        cp.SetSelection y1, i + Len(sWord) + 5, y1, i + Len(sWord) + 5
    End Select
End Sub


Public Function CutSpaces(s As String) As String

  Dim i As Long
  Dim contor As Long

    If (s <> "") Then
        i = Len(s)
        While ((Getc1(s, i) = Asc(" ")) Or (Getc1(s, i) = Asc(vbTab)) Or (Getc1(s, i) = Asc(vbCr)) Or (Getc1(s, i) = Asc(vbCrLf)) Or (Getc1(s, i) = Asc(vbLf)) Or (Getc1(s, i) = 10)) And i >= 2
            contor = contor + 1
            i = i - 1
        Wend
        i = i + 1
        s = Left$(s, Len(s) - contor)

        i = 1
        contor = 0
        While ((Getc1(s, i) = Asc(" ")) Or (Getc1(s, i) = Asc(vbTab)) Or (Getc1(s, i) = Asc(vbCr)) Or (Getc1(s, i) = Asc(vbCrLf)) Or (Getc1(s, i) = Asc(vbLf)) Or (Getc1(s, i) = 10)) And i <= Len(s)
            contor = contor + 1
            If i < Len(s) Then
                i = i + 1
              Else
                GoTo AAA
            End If
        Wend
AAA:
        CutSpaces = Right$(s, Len(s) - contor)
      Else
        CutSpaces = ""
    End If

End Function

Public Function CutSpacesRight(s As String) As String

  Dim i As Long
  Dim contor As Long

    i = Len(s)
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) = " " Then
            contor = contor + 1
          Else
            Exit For
        End If
    Next i

    CutSpacesRight = Left$(s, Len(s) - contor)

End Function

Public Function Getc(s As String, i As Long) As String

    Getc = Right$(Left$(s, i), 1)

End Function

Public Function Getc1(s As String, i As Long) As String

    Getc1 = Asc(Mid$(s, i))

End Function
