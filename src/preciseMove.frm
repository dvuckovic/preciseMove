VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "preciseMove"
   ClientHeight    =   900
   ClientLeft      =   2700
   ClientTop       =   2745
   ClientWidth     =   3825
   Icon            =   "preciseMove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin preciseMove.Hover CloseBtn 
      Height          =   330
      Left            =   15
      TabIndex        =   2
      Top             =   555
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   582
      Caption         =   "&Close"
   End
   Begin VB.Timer tmrKeys 
      Interval        =   20
      Left            =   3240
      Top             =   1560
   End
   Begin VB.TextBox DeltaBox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3390
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   585
      Width           =   390
   End
   Begin VB.Label Label1 
      Caption         =   $"preciseMove.frx":2372
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   3930
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2
Private Type MOUSEINPUT
  dx As Long
  dy As Long
  mouseData As Long
  dwFlags As Long
  dwtime As Long
  dwExtraInfo As Long
End Type
Private Type INPUT_TYPE
  dwType As Long
  xi(0 To 23) As Byte
End Type
Dim inputEvents(0 To 1) As INPUT_TYPE
Dim mouseEvent As MOUSEINPUT
Dim KeyLoop As Byte
Dim KeyResult As Long
Dim x, y As Integer
Private Function InitSendInput(x, y) As Integer
    mouseEvent.dx = x
    mouseEvent.dy = y
    mouseEvent.mouseData = 0
    mouseEvent.dwFlags = MOUSEEVENTF_MOVE 'MOUSEEVENTF_LEFTDOWN Or
    mouseEvent.dwtime = 0
    mouseEvent.dwExtraInfo = 0
    inputEvents(0).dwType = INPUT_MOUSE
    CopyMemory inputEvents(0).xi(0), mouseEvent, Len(mouseEvent)
    'mouseEvent.dx = 0
    'mouseEvent.dy = 0
    'mouseEvent.mouseData = 0
    'mouseEvent.dwFlags = MOUSEEVENTF_LEFTUP
    'mouseEvent.dwtime = 0
    'mouseEvent.dwExtraInfo = 0
    'inputEvents(1).dwType = INPUT_MOUSE
    'CopyMemory inputEvents(1).xi(0), mouseEvent, Len(mouseEvent)
    InitSendInput = SendInput(1, inputEvents(0), Len(inputEvents(0)))
End Function
Private Sub CloseBtn_Click()
    tmrKeys.Enabled = False
    End
End Sub
Private Sub Form_Load()
    DeltaBox.Text = 1.5
End Sub
Private Sub Form_Unload(Cancel As Integer)
    tmrKeys.Enabled = False
    End
End Sub
Private Sub tmrKeys_Timer()
   On Error Resume Next
   For KeyLoop = 1 To 255
       KeyResult = GetAsyncKeyState(KeyLoop)
       Moving = False
       If KeyResult = -32767 Then
            Select Case KeyLoop
                Case vbKeyNumpad2:  x = 0: y = CDec(DeltaBox.Text): Moving = True
                Case vbKeyNumpad4:  x = -CDec(DeltaBox.Text): y = 0: Moving = True
                Case vbKeyNumpad6:  x = CDec(DeltaBox.Text): y = 0: Moving = True
                Case vbKeyNumpad8:  x = 0: y = -CDec(DeltaBox.Text): Moving = True
            End Select
            If Moving Then
                InitSendInput x, y
            End If
        End If
   Next
End Sub
