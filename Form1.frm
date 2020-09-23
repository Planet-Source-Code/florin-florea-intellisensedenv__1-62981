VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1200
      ItemData        =   "Form1.frx":0000
      Left            =   105
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   45
      Width           =   4500
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   855
      Top             =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_GotFocus()
SetFocusToCodeWindow
End Sub

Private Sub Form_Resize()
List1.Top = 0
List1.Left = 0
List1.Width = Me.Width
List1.Height = Me.Height


End Sub

Private Sub Timer1_Timer()
    CheckListBoxNeeded
    If Visible Then SetFocusToCodeWindow
End Sub




Public Function GetSelectedText() As String
If List1.ListIndex > -1 Then
    GetSelectedText = List1.Text
End If
End Function

Public Sub HandleKeyUp()
If List1.ListIndex > 0 Then
    List1.ListIndex = List1.ListIndex - 1
End If
End Sub

Public Sub HandleKeyDown()
If List1.ListIndex < List1.ListCount - 1 Then
    List1.ListIndex = List1.ListIndex + 1
End If
End Sub

Public Sub SetSearchWord(sWord As String)
    Dim i As Long, j As Long, n As Long
    Dim sPattern As String
    sPattern = UCase$(sWord) & "*"
    n = Len(sWord)
    ' This is a very slow way of checking
    ' This could definately be improved
    For i = 0 To List1.ListCount - 1
        If UCase$(List1.List(i)) Like sPattern Then
            List1.ListIndex = i
            Exit Sub
        End If
    Next
    List1.ListIndex = -1
End Sub

