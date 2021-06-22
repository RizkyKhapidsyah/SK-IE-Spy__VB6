VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "IE Spy"
   ClientHeight    =   4950
   ClientLeft      =   1740
   ClientTop       =   2070
   ClientWidth     =   6645
   Icon            =   "iespy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstWindows3 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   6612
   End
   Begin VB.ListBox lstWindows2 
      Height          =   2205
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3132
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   4080
      Top             =   1560
   End
   Begin VB.ListBox lstWindows 
      Height          =   2205
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Label Label3 
      Caption         =   "Press F10 to unhide me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3840
      TabIndex        =   5
      Top             =   3240
      Width           =   2892
   End
   Begin VB.Label Label2 
      Caption         =   "Press F9 to hide me "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   2892
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Visited Url's"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   2652
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim bet As Long
 Dim met As Long



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
w = Format(Date, "ddmmyy")
y = Format(Time, "hhmmss")
x = "IESPY" & w & y & ".Log"
b = lstWindows3.ListCount
Open App.Path & "\" & x For Output As #1
For mt = 0 To b - 1
lstWindows3.ListIndex = mt
Write #1, lstWindows3.Text
Next mt
Close #1
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If GetAsyncKeyState(VK_F9) Then
pp = ShowWindow(Me.hwnd, SW_HIDE)
End If
If GetAsyncKeyState(VK_F10) Then
pp = ShowWindow(Me.hwnd, SW_NORMAL)
End If
Dim hDesktop        As Long
Dim mtr As Long
Dim stext As String * 100
        Dim b As Long
        Dim x As Long
        Dim t As Long
        Dim w As Long
        Dim der As Long
 
         hDesktop = GetDesktopWindow()
        Call ListChildWindows(lstWindows, hDesktop)
        b = lstWindows.ListCount
        For t = 0 To b - 1
        lstWindows.ListIndex = t
        x = InStr(lstWindows.Text, "Microsoft Internet Explorer")
  If x <> 0 Then
  w = InStr(lstWindows.Text, "-")
  der = Left(lstWindows.Text, w - 1)
     
     Call ListChildWindows(lstWindows2, der)
   mp = lstWindows2.ListCount
   For er = 1 To mp
   lstWindows2.ListIndex = er
w = InStr(lstWindows2.Text, "-")
  der = Left(lstWindows2.Text, w - 1)
  dear = GetClassName(der, stext, 100)
  cutdear = Left(stext, dear)
  xx = InStr(cutdear, "Edit")
  If xx <> 0 Then
  Dim wintext As String
Dim slength As Long
Dim retval As Long
Dim tryval As Long
  Dim max As Boolean
  slength = SendMessage(der, WM_GETTEXTLENGTH, ByVal CLng(0), ByVal CLng(0)) + 1
wintext = Space(slength)
Dim ert As Long
ert = slength
retval = SendMessage(der, WM_GETTEXT, ByVal slength, ByVal wintext)
wer = Len(wintext)
wintext = Left(wintext, wer - 1)
max = True
If met <> 0 Then
bet = met
lstWindows.ListIndex = bet

For k = 0 To bet
lstWindows3.ListIndex = k
If lstWindows3.Text = wintext Then
max = False
GoTo 25
End If
25:
Next k
If max = True Then
If GetAsyncKeyState(VK_RETURN) Then
qw = InStr(wintext, "http:")
If qw <> 0 Then lstWindows3.AddItem wintext
End If
If GetAsyncKeyState(VK_LBUTTON) Then
qw = InStr(wintext, "http:")
If qw <> 0 Then lstWindows3.AddItem wintext
End If
met = met + 1
 lstWindows3.ListIndex = met
 
 End If
 End If
If met = 0 Then
lstWindows3.ListIndex = 0
qw = InStr(wintext, "http:")
qs = InStr(wintext, "ftp:")

If (qw <> 0) Or (qs <> 0) Then lstWindows3.AddItem wintext
met = met + 1
End If
End If
 
  
  Next er
    End If
    Next t
    
End Sub
