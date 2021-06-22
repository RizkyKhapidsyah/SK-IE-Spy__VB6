Attribute VB_Name = "iesp"
Option Explicit
Global ftr As String
' Local pointer to the main form list box
Private mListBox        As Control

Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Const VK_F10 = &H79

Public Const VK_F9 = &H78

Public Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" _
        (ByVal hwnd As Long, _
         ByVal lpString As String, _
         ByVal cch As Long) As Long
Public Const SW_HIDE = 0

Public Const SW_NORMAL = 1

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetActiveWindow Lib "user32" () As Long


Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Const WM_GETTEXT = &HD
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXTLENGTH = &HE
Public Const VK_LBUTTON = &H1
 Public Const VK_RETURN = &HD
Private Declare Function EnumChildWindows Lib "user32" _
        (ByVal hWndParent As Long, _
         ByVal lpEnumFunc As Long, _
         ByVal lParam As Long) As Boolean

Public Sub ListChildWindows _
            (ctlListBox As Control, _
             hwnd As Long)

Dim bResult         As Boolean

' Grab the pointer to the main form list box
Set mListBox = ctlListBox

' Clear the referenced list box
mListBox.Clear

' MAke the call to start the callback series
bResult = EnumChildWindows(hwnd, _
            AddressOf ChildCallback, 0&)

End Sub

Public Function ChildCallback _
            (ByVal hWndChild As Long, _
             lRaram As Long) As Boolean

Dim sTempStr            As String
Dim sListText           As String
Dim lResult             As Long

' Get the window text for the child window
sTempStr = String(255, " ")
lResult = GetWindowText(hWndChild, _
            ByVal sTempStr, 254&)

' Build a string containing the window text
If InStr(1, sTempStr, vbNullChar) > 0 Then
    sTempStr = Left$(sTempStr, Len(sTempStr) - 1)
End If

' Concatenate the window handle and text
sListText = hWndChild & _
        " - " & sTempStr

' Add the item to the list box
mListBox.AddItem sListText

' Set the return value to keep the callback going
ChildCallback = True

End Function
