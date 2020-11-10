VERSION 5.00
Begin VB.UserControl Hover 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   ClipControls    =   0   'False
   DefaultCancel   =   -1  'True
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   435
   ScaleWidth      =   600
   ToolboxBitmap   =   "Hover.ctx":0000
   Begin VB.Label LabelRight 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   0
      TabIndex        =   2
      Top             =   0
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label LabelMnemonic 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   0
      TabIndex        =   1
      Top             =   0
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label LabelLeft 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   30
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Visible         =   0   'False
      Begin VB.Menu mnuOption1 
         Caption         =   ""
      End
      Begin VB.Menu mnuOption2 
         Caption         =   ""
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClick 
         Caption         =   "&Click"
      End
   End
End
Attribute VB_Name = "Hover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEL) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
 
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SIZEL
    cx As Long
    cy As Long
End Type

Enum EdgeType
    [None] = 0
    [Thick Raised] = &H5
    [Thick Sunken] = (&H2 Or &H8)
End Enum

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Dim rRect As RECT
Dim fRect As RECT
Dim bHovering As Boolean
Dim bPressed As Boolean
Dim bFocused As Boolean
Dim nTextLeft As Long
Dim nTextTop As Long
Dim nControlHeight As Long
Dim nControlWidth As Long

Const m_def_BackColourNormal = vbButtonFace
Const m_def_BackColourHover = vbButtonFace
Const m_def_BackColourClick = vbButtonFace
Const m_def_ColourClick = vbHighlight
Const m_def_MouseHand = False
Const m_def_ColourHover = vbHighlight
Const m_def_ColourText = vbButtonText
Const m_def_BorderHover = &H5
Const m_def_BorderNormal = &H0
Const m_def_BorderClick = (&H2 Or &H8)
Const m_def_Enabled = True

Dim m_BorderHover As EdgeType
Dim m_BorderNormal As EdgeType
Dim m_BorderClick As EdgeType
Dim m_MouseHand As Boolean
Dim m_Enabled As Boolean
Dim m_Caption As String
Dim m_CaptionWOM As String
Dim m_LinkToURL As Boolean
Dim m_URL As String
Dim m_Image As Picture
Dim m_Image2 As Picture
Dim m_Context As Boolean
Dim m_ContextOption1 As String
Dim m_ContextOption2 As String

Event Click()
Event ContextOption1()
Event ContextOption2()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Private Sub LabelLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub LabelLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove Button, Shift, x, y
End Sub
Private Sub LabelLeft_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub
Private Sub LabelMnemonic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub LabelMnemonic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove Button, Shift, x, y
End Sub
Private Sub LabelMnemonic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub
Private Sub LabelRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub LabelRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove Button, Shift, x, y
End Sub
Private Sub LabelRight_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub
Private Sub mnuClick_Click()
    RaiseEvent Click
    DoEvents
End Sub
Private Sub mnuOption1_Click()
    RaiseEvent ContextOption1
End Sub
Private Sub mnuOption2_Click()
    RaiseEvent ContextOption2
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub
Private Sub UserControl_GotFocus()
    bFocused = True
    UserControl_Paint
End Sub
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_Caption = Ambient.DisplayName
    m_CaptionWOM = m_Caption
    m_BorderHover = m_def_BorderHover
    m_BorderNormal = m_def_BorderNormal
    m_BorderClick = m_def_BorderClick
    m_Context = False
    m_ContextOption1 = "&Edit"
    m_ContextOption2 = "&Delete"
    Set m_Image = Nothing
    Set m_Image2 = Nothing
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then bPressed = True: UserControl_Paint
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then bPressed = False: UserControl_Paint: RaiseEvent Click
    If KeyCode = 93 Then UserControl_MouseDown vbRightButton, 0, 0, 0
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub UserControl_LostFocus()
    bFocused = False
    UserControl_Paint
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Temp As Boolean
    If (x < 0) Or (y < 0) Or (x > nControlWidth) Or (y > nControlHeight) Then
        Temp = False
        Call ReleaseCapture
    Else
        Temp = True
        Call SetCapture(UserControl.hwnd)
    End If
    If bHovering <> Temp Then
        If Button <> 0 Then bPressed = True
        bHovering = Temp
        If bHovering = False Then bPressed = False
        UserControl_Paint
    End If
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Button
        Case vbRightButton
            If m_Context Then UserControl.PopupMenu mnuContext, , , , mnuClick
        Case vbLeftButton
            bPressed = True
            UserControl_Paint
    End Select
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim bTemp As Boolean
    bTemp = bPressed
    bPressed = False
    bHovering = False
    UserControl_Paint
    If bTemp = True And m_Enabled = True Then
        RaiseEvent Click
        DoEvents
    End If
End Sub
Private Sub UserControl_Paint()
    Dim rct As RECT
    UserControl.Cls
    If m_Enabled = True Then
        UserControl.Enabled = True
        LabelLeft.Enabled = True
        LabelMnemonic.Enabled = True
        LabelRight.Enabled = True
        If m_Image Is Nothing Then
            If bFocused Then
                UserControl.ForeColor = vbBlack
                DrawFocusRect UserControl.hdc, fRect
            End If
            If bPressed Then
                DrawEdge UserControl.hdc, rRect, m_BorderClick, &H100F
                PlaceLabel True
            Else
                If bHovering Then
                    DrawEdge UserControl.hdc, rRect, m_BorderHover, &H100F
                Else
                    DrawEdge UserControl.hdc, rRect, m_BorderNormal, &H100F
                End If
                PlaceLabel False
            End If
        Else
            If Len(m_Caption) > 0 Then
                If bFocused Then
                    UserControl.ForeColor = vbBlack
                    DrawFocusRect UserControl.hdc, fRect
                End If
                If bPressed Then
                    PlaceLabel True
                    FigureTextSize
                    PaintPicture m_Image, nTextLeft * Screen.TwipsPerPixelX - 240, nTextTop * Screen.TwipsPerPixelY - 15
                    DrawEdge UserControl.hdc, rRect, m_BorderClick, &H100F
                Else
                    PlaceLabel False
                    FigureTextSize
                    PaintPicture m_Image, nTextLeft * Screen.TwipsPerPixelX - 250, nTextTop * Screen.TwipsPerPixelY - 25
                    If bHovering Then
                        DrawEdge UserControl.hdc, rRect, m_BorderHover, &H100F
                    Else
                        DrawEdge UserControl.hdc, rRect, m_BorderNormal, &H100F
                    End If
                End If
            Else
                If bFocused Then
                    UserControl.ForeColor = vbBlack
                    DrawFocusRect UserControl.hdc, fRect
                End If
                If bPressed Then
                    FigureTextSize
                    PaintPicture m_Image, UserControl.Width / 2 - 110, UserControl.Height / 2 - 110
                    DrawEdge UserControl.hdc, rRect, m_BorderClick, &H100F
                Else
                    FigureTextSize
                    PaintPicture m_Image, UserControl.Width / 2 - 120, UserControl.Height / 2 - 120
                    If bHovering Then
                        DrawEdge UserControl.hdc, rRect, m_BorderHover, &H100F
                    Else
                        DrawEdge UserControl.hdc, rRect, m_BorderNormal, &H100F
                    End If
                End If
            End If
        End If
    Else
        UserControl.Enabled = False
        UserControl.Cls
        'LabelLeft.Visible = False
        'LabelMnemonic.Visible = False
        'LabelRight.Visible = False
        LabelLeft.Enabled = False
        LabelMnemonic.Enabled = False
        LabelRight.Enabled = False
        PlaceLabel False
        FigureTextSize
        If Not m_Image2 Is Nothing Then
            PaintPicture m_Image2, nTextLeft * Screen.TwipsPerPixelX - 250, nTextTop * Screen.TwipsPerPixelY - 25, , , , , , , vbMergePaint
        End If
        'UserControl.ForeColor = vb3DHighlight
        'Call TextOut(UserControl.hdc, nTextLeft + 2, nTextTop + 1, m_CaptionWOM, Len(m_CaptionWOM))
        'UserControl.ForeColor = vb3DShadow
        'Call TextOut(UserControl.hdc, nTextLeft + 1, nTextTop, m_CaptionWOM, Len(m_CaptionWOM))
    End If
End Sub
Private Sub UserControl_Resize()
    rRect.Left = 0
    rRect.Top = 0
    rRect.Bottom = UserControl.Height \ Screen.TwipsPerPixelY
    rRect.Right = UserControl.Width \ Screen.TwipsPerPixelX
    fRect.Left = 3
    fRect.Top = 3
    fRect.Bottom = UserControl.Height \ Screen.TwipsPerPixelY - 3
    fRect.Right = UserControl.Width \ Screen.TwipsPerPixelX - 3
    nControlHeight = UserControl.Height
    nControlWidth = UserControl.Width
    FigureTextSize
    UserControl_Paint
End Sub
Private Function CaptionLeft$()
    Dim Mnemonic As Integer
    Mnemonic = InStr(m_Caption, "&")
    If Mnemonic <> 0 Then
        CaptionLeft = Left(m_Caption, Mnemonic - 1)
    Else
        CaptionLeft = m_Caption
    End If
End Function
Private Function CaptionMnemonic$()
    Dim Mnemonic As Integer
    Mnemonic = InStr(m_Caption, "&")
    If Mnemonic <> 0 Then
        UserControl.AccessKeys = Mid(m_Caption, Mnemonic + 1, 1)
        CaptionMnemonic = UserControl.AccessKeys
    End If
End Function
Private Function CaptionRight$()
    Dim Mnemonic As Integer
    Mnemonic = InStr(m_Caption, "&")
    If Mnemonic <> 0 Then
        CaptionRight = Mid(m_Caption, Mnemonic + 2)
    End If
End Function
Private Sub FigureTextSize()
    Dim slTemp As SIZEL
    Call GetTextExtentPoint32(hdc, m_CaptionWOM, Len(m_CaptionWOM), slTemp)
    nTextLeft = (rRect.Right - slTemp.cx) \ 2
    nTextTop = (rRect.Bottom - slTemp.cy) \ 2
End Sub
Private Sub PlaceLabel(Moved As Boolean)
    Dim WidthLeft, WidthMnemonic, WidthRight As Integer
    If Len(CaptionLeft) > 0 Then
        LabelLeft.Visible = True
        LabelLeft.Caption = CaptionLeft
        WidthLeft = LabelLeft.Width
    Else
        LabelLeft.Visible = False
        LabelLeft.Caption = vbNullString
        WidthLeft = 0
    End If
    If Len(CaptionMnemonic) > 0 Then
        LabelMnemonic.Visible = True
        LabelMnemonic.Caption = CaptionMnemonic
        WidthMnemonic = LabelMnemonic.Width
    Else
        LabelMnemonic.Visible = False
        LabelMnemonic.Caption = vbNullString
        WidthMnemonic = 0
    End If
    If Len(CaptionRight) > 0 Then
        LabelRight.Visible = True
        LabelRight.Caption = CaptionRight
        WidthRight = LabelRight.Width
    Else
        LabelRight.Visible = False
        LabelRight.Caption = vbNullString
        WidthRight = 0
    End If
    If Moved Then
        LabelLeft.Left = (UserControl.Width - (WidthLeft + WidthMnemonic + WidthRight)) / 2 + 10
        LabelLeft.Top = (UserControl.Height - LabelLeft.Height) / 2 + 10
        LabelMnemonic.Left = LabelLeft.Left + WidthLeft
        LabelMnemonic.Top = LabelLeft.Top
        LabelRight.Left = LabelMnemonic.Left + WidthMnemonic
        LabelRight.Top = LabelLeft.Top
    Else
        LabelLeft.Left = (UserControl.Width - (WidthLeft + WidthMnemonic + WidthRight)) / 2 - 5
        LabelLeft.Top = (UserControl.Height - LabelLeft.Height) / 2 - 5
        LabelMnemonic.Left = LabelLeft.Left + WidthLeft
        LabelMnemonic.Top = LabelLeft.Top
        LabelRight.Left = LabelMnemonic.Left + WidthMnemonic
        LabelRight.Top = LabelLeft.Top
    End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Caption = PropBag.ReadProperty("Caption", UserControl.Ambient.DisplayName)
    m_BorderHover = PropBag.ReadProperty("BorderHover", m_def_BorderHover)
    m_BorderNormal = PropBag.ReadProperty("BorderNormal", m_def_BorderNormal)
    m_BorderClick = PropBag.ReadProperty("BorderClick", m_def_BorderClick)
    m_Context = PropBag.ReadProperty("Context", False)
    m_ContextOption1 = PropBag.ReadProperty("ContextOption1", "&Edit")
    m_ContextOption2 = PropBag.ReadProperty("ContextOption2", "&Delete")
    Set m_Image = PropBag.ReadProperty("ImageIcon", Nothing)
    Set m_Image2 = PropBag.ReadProperty("ImageBitmap", Nothing)
    m_CaptionWOM = CaptionLeft & CaptionMnemonic & CaptionRight
    mnuOption1.Caption = m_ContextOption1
    mnuOption2.Caption = m_ContextOption2
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Ambient.DisplayName)
    Call PropBag.WriteProperty("BorderHover", m_BorderHover, m_def_BorderHover)
    Call PropBag.WriteProperty("BorderNormal", m_BorderNormal, m_def_BorderNormal)
    Call PropBag.WriteProperty("BorderClick", m_BorderClick, m_def_BorderClick)
    Call PropBag.WriteProperty("ImageIcon", m_Image, Nothing)
    Call PropBag.WriteProperty("ImageBitmap", m_Image2, Nothing)
    Call PropBag.WriteProperty("Context", m_Context, False)
    Call PropBag.WriteProperty("ContextOption1", m_ContextOption1, "&Edit")
    Call PropBag.WriteProperty("ContextOption2", m_ContextOption2, "&Delete")
End Sub
Public Property Get BorderHover() As EdgeType
    BorderHover = m_BorderHover
End Property
Public Property Let BorderHover(ByVal New_BorderHover As EdgeType)
    m_BorderHover = New_BorderHover
    PropertyChanged "BorderHover"
End Property
Public Property Get BorderNormal() As EdgeType
    BorderNormal = m_BorderNormal
End Property
Public Property Let BorderNormal(ByVal New_BorderNormal As EdgeType)
    m_BorderNormal = New_BorderNormal
    PropertyChanged "BorderNormal"
    UserControl_Paint
End Property
Public Property Get BorderClick() As EdgeType
    BorderClick = m_BorderClick
End Property
Public Property Let BorderClick(ByVal New_BorderClick As EdgeType)
    m_BorderClick = New_BorderClick
    PropertyChanged "BorderClick"
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    m_CaptionWOM = CaptionLeft & CaptionMnemonic & CaptionRight
    UserControl_Paint
End Property
Public Property Get ImageIcon() As Picture
    Set ImageIcon = m_Image
End Property
Public Property Set ImageIcon(ByVal New_Image As Picture)
    Set m_Image = New_Image
    PropertyChanged "ImageIcon"
    Refresh
End Property
Public Property Get ImageBitmap() As Picture
    Set ImageBitmap = m_Image2
End Property
Public Property Set ImageBitmap(ByVal New_Image As Picture)
    Set m_Image2 = New_Image
    PropertyChanged "ImageBitmap"
    Refresh
End Property
Public Property Get Context() As Boolean
    Context = m_Context
End Property
Public Property Let Context(ByVal New_Context As Boolean)
    m_Context = New_Context
    PropertyChanged "Context"
End Property
Public Property Get ContextOption1() As String
    ContextOption1 = m_ContextOption1
End Property
Public Property Let ContextOption1(ByVal New_ContextOption1 As String)
    m_ContextOption1 = New_ContextOption1
    PropertyChanged "ContextOption1"
    mnuOption1.Caption = m_ContextOption1
End Property
Public Property Get ContextOption2() As String
    ContextOption2 = m_ContextOption2
End Property
Public Property Let ContextOption2(ByVal New_ContextOption2 As String)
    m_ContextOption2 = New_ContextOption2
    PropertyChanged "ContextOption2"
    mnuOption2.Caption = m_ContextOption2
End Property

