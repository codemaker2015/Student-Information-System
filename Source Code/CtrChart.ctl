VERSION 5.00
Begin VB.UserControl CtrChart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Label crsLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   525
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   1140
      TabIndex        =   0
      Top             =   1950
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      DrawMode        =   4  'Mask Not Pen
      Visible         =   0   'False
      X1              =   47
      X2              =   283
      Y1              =   100
      Y2              =   100
   End
End
Attribute VB_Name = "CtrChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'  mChart.Ctl
'
'  01/08/2014 by Vishnu Sivan
'  codemaker2014@gmail.com
'


Option Explicit
Private Type GRADIENT_TRIANGLE
   Vertex1 As Long
   Vertex2 As Long
   Vertex3 As Long
End Type
Private Type TRIVERTEX
   x     As Long
   y     As Long
   Red   As Integer
   Green As Integer
   Blue  As Integer
   Alpha As Integer
End Type
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2&
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, _
                                                                                  pVertex As TRIVERTEX, _
                                                                                  ByVal dwNumVertex As Long, _
                                                                                  pMesh As GRADIENT_TRIANGLE, _
                                                                                  ByVal dwNumMesh As Long, _
                                                                                  ByVal dwMode As Long) As Long
Private Declare Function ColorAdjustLuma Lib "shlwapi.dll" (ByVal clrRGB As Long, _
                                                            ByVal n As Long, _
                                                            ByVal fScale As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, _
                                                                     ByVal W As Long, _
                                                                     ByVal E As Long, _
                                                                     ByVal O As Long, _
                                                                     ByVal W As Long, _
                                                                     ByVal I As Long, _
                                                                     ByVal u As Long, _
                                                                     ByVal S As Long, _
                                                                     ByVal C As Long, _
                                                                     ByVal OP As Long, _
                                                                     ByVal CP As Long, _
                                                                     ByVal Q As Long, _
                                                                     ByVal PAF As Long, _
                                                                     ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private mValue() As Double
Private mLabel() As String
Private mBarColor() As Long
Private mRedraw As Boolean
Private mBgLineVisible As Boolean
Private BarEvidenziata As Single
Private crY As Single
Private mMax As Double
Private mHighLight As Long
Private mMainBarColor As Long
Private mSpace As Integer

Public Property Get Space() As Integer
   Space = mSpace
End Property

Public Property Let Space(ByVal NewVal As Integer)
 mSpace = NewVal
 Draw
End Property

Public Property Get BackColor() As OLE_COLOR

   BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal nColor As OLE_COLOR)

   UserControl.BackColor = nColor
   Draw

End Property

Public Property Get BarColor(ByVal index As Integer) As Long

   On Error Resume Next
   If mBarColor(index) = -1 Then
      BarColor = MainBarColor
     Else 'NOT MBARCOLOR(INDEX)...
      BarColor = mBarColor(index)
   End If

End Property

Public Property Let BarColor(ByVal index As Integer, _
                             ByVal NewVal As Long)

   On Error Resume Next
   mBarColor(index) = NewVal
   Draw
   On Error GoTo 0

End Property

Private Sub Draw()

   If mRedraw Then
      GradientBg
      Drawlabel
      DrawBar
      UserControl.Refresh
   End If

End Sub

Private Sub DrawBar()

   Dim x      As Long
   Dim y      As Long
   Dim width  As Double
   Dim height As Long
   Dim a      As Integer
   Dim stp    As Double

   On Error GoTo noBar
   width = (UserControl.ScaleWidth - 125) / ValueCount
   y = UserControl.ScaleHeight - 50
   stp = (UserControl.ScaleHeight - 100) / GetMax
   UserControl.ForeColor = 0
   For a = 1 To ValueCount
      x = 60 + ((a - 1) * width)
      If UserControl.Ambient.UserMode Then
         height = stp * mValue(a)
        Else 'USERCONTROL.AMBIENT.USERMODE = FALSE/0
         height = stp * Rnd(Timer) * GetMax
      End If
      DrawSignleBar x, y, width - mSpace, height, IIf(BarEvidenziata = a Or (a = 5 And Not UserControl.Ambient.UserMode), mHighLight, BarColor(a)), mLabel(a)
   Next a
noBar:

End Sub

Private Sub drawBgLine(ByVal x As Long, _
                       ByVal y As Long, _
                       ByVal lngHeight As Long, _
                       ByVal lngWidth As Long, _
                       ByVal lvl As Integer)

   Dim TriVert(6) As TRIVERTEX
   Dim gTRi(4)    As GRADIENT_TRIANGLE
   Dim FrC        As Long

   If lvl Then
      FrC = GetColorLvl(UserControl.BackColor, 4.8)
     Else 'LVL = FALSE/0
      FrC = GetColorLvl(UserControl.BackColor, 5.8)
   End If
   TriVert(0).x = x
   TriVert(0).y = y
   GradientFillColor TriVert(0), GetColorLvl(FrC, 7.5)
   TriVert(1).x = x
   TriVert(1).y = y - lngHeight
   GradientFillColor TriVert(1), GetColorLvl(FrC, 8)
   TriVert(2).x = x + 30
   TriVert(2).y = y - lngHeight - 30
   GradientFillColor TriVert(2), GetColorLvl(FrC, 5.2)
   TriVert(3).x = x + 30
   TriVert(3).y = y - 30
   GradientFillColor TriVert(3), GetColorLvl(FrC, 4.8)
   TriVert(4).x = x + lngWidth
   TriVert(4).y = y - lngHeight - 30
   GradientFillColor TriVert(4), GetColorLvl(FrC, 5.5)
   TriVert(5).x = x + lngWidth
   TriVert(5).y = y - 30
   GradientFillColor TriVert(5), GetColorLvl(FrC, 6)
   With gTRi(0)
      .Vertex1 = 0
      .Vertex2 = 1
      .Vertex3 = 2
   End With 'gTRi(0)
   With gTRi(1)
      .Vertex1 = 0
      .Vertex2 = 2
      .Vertex3 = 3
   End With 'gTRi(1)
   With gTRi(2)
      .Vertex1 = 2
      .Vertex2 = 3
      .Vertex3 = 4
   End With 'gTRi(2)
   With gTRi(3)
      .Vertex1 = 3
      .Vertex2 = 4
      .Vertex3 = 5
   End With 'gTRi(3)
   GradientFillTriangle UserControl.hdc, TriVert(0), 6, gTRi(0), 4, GRADIENT_FILL_TRIANGLE

End Sub

Private Sub DrawFood()

   Dim TriVert(6) As TRIVERTEX
   Dim gTRi(4)    As GRADIENT_TRIANGLE

   TriVert(0).x = 70
   TriVert(0).y = UserControl.ScaleHeight - 70
   GradientFillColor TriVert(0), &HE4F6FF
   TriVert(1).x = 5
   TriVert(1).y = UserControl.ScaleHeight - 5
   GradientFillColor TriVert(1), &HFFFFFF
   TriVert(2).x = UserControl.ScaleWidth - 5
   TriVert(2).y = UserControl.ScaleHeight - 70
   GradientFillColor TriVert(2), &HFFFFFF
   TriVert(3).x = UserControl.ScaleWidth - 70
   TriVert(3).y = UserControl.ScaleHeight - 5
   GradientFillColor TriVert(3), &HE4F6FF
   With gTRi(0)
      .Vertex1 = 0
      .Vertex2 = 1
      .Vertex3 = 2
   End With 'gTRi(0)
   With gTRi(1)
      .Vertex1 = 1
      .Vertex2 = 2
      .Vertex3 = 3
   End With 'gTRi(1)
   GradientFillTriangle UserControl.hdc, TriVert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE

End Sub

Private Sub Drawlabel()

   If mBgLineVisible Then
      DrawYlabelon
     Else 'MBGLINEVISIBLE = FALSE/0
      DrawYlabeloff
   End If

End Sub

Private Sub DrawRotatedText(ByVal Txt As String, _
                            ByVal x As Single, _
                            ByVal y As Single, _
                            Optional font_name = "Arial", _
                            Optional size = 16, _
                            Optional weight = 100, _
                            Optional escapement = 450, _
                            Optional use_italic = False, _
                            Optional use_underline = False, _
                            Optional use_strikethrough = False)

   Const CLIP_LH_ANGLES = 16
   Const PI = 3.14159625
   Dim newfont As Long
   Dim oldfont As Long

   newfont = CreateFont(size, 0, escapement, escapement, weight, use_italic, use_underline, use_strikethrough, 0, 0, CLIP_LH_ANGLES, 0, 0, font_name)
   oldfont = SelectObject(hdc, newfont)
   CurrentX = x
   CurrentY = y
   Print Txt;
   newfont = SelectObject(hdc, oldfont)
   DeleteObject newfont

End Sub

Private Sub DrawSignleBar(ByVal x As Long, _
                          ByVal y As Long, _
                          ByVal lngWidth As Long, _
                          ByVal lngHeight As Long, _
                          ByVal ucolor As Long, _
                          Optional Txt As String = vbNullString)

   Dim TriVert(3) As TRIVERTEX
   Dim gTRi(2)    As GRADIENT_TRIANGLE

   TriVert(0).x = x
   TriVert(0).y = y
   TriVert(1).x = x + lngWidth
   TriVert(1).y = y
   TriVert(2).x = x
   TriVert(2).y = y - lngHeight
   TriVert(3).x = x + lngWidth
   TriVert(3).y = y - lngHeight
   GradientFillColor TriVert(0), GetColorLvl(ucolor, 3)
   GradientFillColor TriVert(1), GetColorLvl(ucolor, 4)
   GradientFillColor TriVert(2), GetColorLvl(ucolor, 7)
   GradientFillColor TriVert(3), GetColorLvl(ucolor, 8)
   With gTRi(0)
      .Vertex1 = 0
      .Vertex2 = 1
      .Vertex3 = 2
   End With 'gTRi(0)
   With gTRi(1)
      .Vertex1 = 1
      .Vertex2 = 2
      .Vertex3 = 3
   End With 'gTRi(1)
   GradientFillTriangle UserControl.hdc, TriVert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
   'side
   TriVert(1).x = x + lngWidth + 20
   TriVert(1).y = y - 20
   TriVert(0).x = x + lngWidth
   TriVert(0).y = y
   TriVert(2).x = x + lngWidth
   TriVert(2).y = y - lngHeight
   TriVert(3).x = x + lngWidth + 20
   TriVert(3).y = y - lngHeight - 20
   GradientFillColor TriVert(0), GetColorLvl(ucolor, 6)
   GradientFillColor TriVert(2), GetColorLvl(ucolor, 8)
   GradientFillColor TriVert(3), GetColorLvl(ucolor, 7)
   GradientFillColor TriVert(1), GetColorLvl(ucolor, 5)
   With gTRi(0)
      .Vertex1 = 0
      .Vertex2 = 1
      .Vertex3 = 2
   End With 'gTRi(0)
   With gTRi(1)
      .Vertex1 = 1
      .Vertex2 = 2
      .Vertex3 = 3
   End With 'gTRi(1)
   GradientFillTriangle UserControl.hdc, TriVert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
   'Top
   TriVert(0).x = x
   TriVert(0).y = y - lngHeight
   TriVert(1).x = x + lngWidth
   TriVert(1).y = y - lngHeight
   TriVert(2).x = x + 20
   TriVert(2).y = y - lngHeight - 20
   TriVert(3).x = x + lngWidth + 20
   TriVert(3).y = y - lngHeight - 20
   GradientFillColor TriVert(0), GetColorLvl(ucolor, 7) 'fl
   GradientFillColor TriVert(1), GetColorLvl(ucolor, 9) 'fr
   GradientFillColor TriVert(2), GetColorLvl(ucolor, 4) 'bl
   GradientFillColor TriVert(3), GetColorLvl(ucolor, 6) 'br
   With gTRi(0)
      .Vertex1 = 0
      .Vertex2 = 1
      .Vertex3 = 2
   End With 'gTRi(0)
   With gTRi(1)
      .Vertex1 = 1
      .Vertex2 = 2
      .Vertex3 = 3
   End With 'gTRi(1)
   GradientFillTriangle UserControl.hdc, TriVert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
   Reflex x, y, lngWidth, lngHeight, ucolor
   UserControl.ForeColor = GetColorLvl(ucolor, 3.5)
   DrawRotatedText Txt, 0, -100
   x = x + (lngWidth / 2) - UserControl.CurrentX
   y = UserControl.ScaleHeight - 60 + UserControl.CurrentX
   DrawRotatedText Txt, x, y

End Sub

Private Sub DrawYlabeloff()

   UserControl.ForeColor = 0
   DrawFood

End Sub

Private Sub DrawYlabelon()

   Dim tY    As Long
   Dim my    As Double
   Dim ly    As Integer
   Dim stp   As Double
   Dim CrLvl As Integer

   With UserControl
      stp = ((.ScaleHeight - 100) / 10)
      .ForeColor = &H96BAD4
      my = -1
      DrawFood
      For tY = 0 To 9
         ly = 0
         For my = .ScaleHeight - (50 + (tY * stp)) To .ScaleHeight - (50 + ((tY + 1) * stp)) Step -stp / 10
            ly = ly + 1
            If (ly Mod 10) = 0 Then
               UserControl.Line (45, my)-(50, my)
              ElseIf (ly Mod 5) = 0 Then 'NOT (LY...
               UserControl.Line (47, my)-(50, my)
              Else 'NOT (LY...
               UserControl.Line (48, my)-(50, my)
            End If
         Next my
         CrLvl = (CrLvl + 1) Mod 2
         drawBgLine 50, .ScaleHeight - (50 + (tY * stp)), stp, .ScaleWidth - 80, CrLvl
      Next tY
      UserControl.Line (50, .ScaleHeight - 50)-(50, UserControl.ScaleHeight - (50 + (tY * stp)))
   End With 'USERCONTROL

End Sub

Private Function GetColor(ByVal Value As Long) As Long

   If Value And &HFF000000 Then
      GetColor = GetSysColor(Value And &HFF)
     Else 'NOT VALUE...
      GetColor = Value
   End If

End Function

Private Function GetColorLvl(ByVal vColor As Long, _
                             ByVal lvl As Double) As Long

   lvl = lvl - 5
   lvl = (lvl * 200) - 1
   vColor = GetColor(vColor)
   GetColorLvl = ColorAdjustLuma(vColor, lvl, 1)

End Function

Private Function GetCurrBar(ByVal x As Single, _
                            ByVal y As Single) As Single

   Dim width  As Single
   Dim stp    As Double
   Dim a      As Integer
   Dim height As Integer
   Dim x1     As Integer
   Dim x2     As Integer
   Dim y1     As Integer
   Dim y2     As Integer
   
   On Error Resume Next
   If Not x < 60 Or y > UserControl.ScaleHeight - 50 Then
      width = (UserControl.ScaleWidth - 125) / ValueCount
      stp = (UserControl.ScaleHeight - 100) / GetMax
      GetCurrBar = 0
      For a = ValueCount To 1 Step -1
         height = stp * mValue(a)
         x1 = 60 + ((a - 1) * (width))
         x2 = x1 + width + 20 - mSpace
         y1 = UserControl.ScaleHeight - 50
         y2 = y1 - height - 20
         If x > x1 Then
            If x < x2 Then
               If y < y1 Then
                  If y > y2 Then
                     If Not ((Abs(y - y2) + Abs(x - x1) < 20) Or (Abs(y - y1) + Abs(x - x2) < 20)) Then
                        GetCurrBar = a
                        Exit For
                     End If
                  End If
               End If
            End If
         End If
      Next a
   End If

End Function

Private Function GetMax() As Double

   Dim Tmp As Double
   Dim a   As Integer
   
   On Error GoTo NoValue
   If mMax = -1 Then
      For a = 1 To ValueCount
         If Tmp < mValue(a) Then
            Tmp = mValue(a)
         End If
      Next a
      If Tmp = 0 Then
         Tmp = 10
      End If
      GetMax = Tmp
     Else 'NOT MMAX...
      GetMax = mMax
   End If
   
NoValue:
    If Err Then
       GetMax = 10
    End If
  
End Function

Private Sub GradientBg()

   Dim TriVert(4) As TRIVERTEX
   Dim gTRi(2)    As GRADIENT_TRIANGLE

   TriVert(0).x = 0
   TriVert(0).y = 0
   GradientFillColor TriVert(0), GetColorLvl(UserControl.BackColor, 8)
   TriVert(1).x = UserControl.ScaleWidth
   TriVert(1).y = 0
   GradientFillColor TriVert(1), GetColorLvl(UserControl.BackColor, 7)
   TriVert(2).x = 0
   TriVert(2).y = UserControl.ScaleHeight
   GradientFillColor TriVert(2), GetColorLvl(UserControl.BackColor, 6)
   TriVert(3).x = UserControl.ScaleWidth
   TriVert(3).y = UserControl.ScaleHeight
   GradientFillColor TriVert(3), GetColorLvl(UserControl.BackColor, 5)
   gTRi(0).Vertex1 = 0
   gTRi(0).Vertex2 = 1
   gTRi(0).Vertex3 = 2
   gTRi(1).Vertex1 = 1
   gTRi(1).Vertex2 = 2
   gTRi(1).Vertex3 = 3
   GradientFillTriangle UserControl.hdc, TriVert(0), 4, gTRi(0), 3, GRADIENT_FILL_TRIANGLE

End Sub

Private Sub GradientFillColor(ByRef tTV As TRIVERTEX, _
                              ByVal iColor As Long)

   Dim iRed   As Long
   Dim iGreen As Long
   Dim iBlue  As Long

   '/* Separate color into RGB
   iRed = (iColor And &HFF&) * &H100&
   iGreen = (iColor And &HFF00&)
   iBlue = (iColor And &HFF0000) \ &H100&
   '/* Make Red color a UShort
   If (iRed And &H8000&) = &H8000& Then
      tTV.Red = (iRed And &H7F00&)
      tTV.Red = tTV.Red Or &H8000
     Else 'NOT (IRED...
      tTV.Red = iRed
   End If
   '/* Make Green color a UShort
   If (iGreen And &H8000&) = &H8000& Then
      tTV.Green = (iGreen And &H7F00&)
      tTV.Green = tTV.Green Or &H8000
     Else 'NOT (IGREEN...
      tTV.Green = iGreen
   End If
   '/* Make Blue color a UShort
   If (iBlue And &H8000&) = &H8000& Then
      tTV.Blue = (iBlue And &H7F00&)
      tTV.Blue = tTV.Blue Or &H8000
     Else 'NOT (IBLUE...
      tTV.Blue = iBlue
   End If
   tTV.Alpha = 0

End Sub

Public Property Get HighLight() As OLE_COLOR

   HighLight = mHighLight

End Property

Public Property Let HighLight(ByVal nColor As OLE_COLOR)

   mHighLight = nColor
   Draw
End Property

Public Property Get MainBarColor() As OLE_COLOR

   MainBarColor = mMainBarColor

End Property

Public Property Let MainBarColor(ByVal nColor As OLE_COLOR)

   mMainBarColor = nColor
   Draw

End Property

Public Property Get max() As Double

   max = GetMax

End Property

Public Property Let max(ByVal vMax As Double)

   mMax = vMax
   Draw

End Property

Public Property Get Redraw() As Boolean

   Redraw = mRedraw

End Property

Public Property Let Redraw(ByVal NewVal As Boolean)

   mRedraw = NewVal
   Draw

End Property

Public Property Get ShowBGLine() As Boolean

   ShowBGLine = mBgLineVisible

End Property

Public Property Let ShowBGLine(ByVal NewVal As Boolean)

   mBgLineVisible = NewVal
   Draw

End Property

Public Property Get Text(ByVal index As Integer) As String

   Text = mLabel(index)

End Property

Public Property Let Text(ByVal index As Integer, _
                         ByVal NewVal As String)

   mLabel(index) = NewVal
   Draw

End Property

Private Sub UserControl_Initialize()
   mSpace = 8
   UserControl.FontName = "Arial"
   mMax = -1
   Redraw = False
   UserControl.BackColor = &HAAE7FF
   mMainBarColor = &HDA8448
   ValueCount = 10
   mHighLight = &HFF
   mBgLineVisible = True
   Redraw = True


End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)

   Dim cBar   As Single
   Dim height As Single
   Dim stp    As Single
   Dim width  As Long
   Dim Mx     As Double
   Dim Value  As Integer


   On Error Resume Next
   Line1.y1 = y
   Line1.y2 = y
   Mx = GetMax
   Value = Round((-(y - UserControl.ScaleHeight + 50) * Mx) / (UserControl.ScaleHeight - 100))
   Line1.Visible = Not (Value < 0 Or Value > Mx Or x < 50 Or x > UserControl.ScaleWidth - 50)
   crsLabel.Move 40 - crsLabel.width, y - (crsLabel.height / 2)
   cBar = GetCurrBar(x, y)
   crsLabel.Caption = Round((-(y - UserControl.ScaleHeight + 50) * Mx) / (UserControl.ScaleHeight - 100))
   crY = y
   If cBar And BarEvidenziata <> cBar Then
      BarEvidenziata = cBar
      Draw
      stp = (UserControl.ScaleHeight - 100) / Mx
      width = (UserControl.ScaleWidth - 125) / ValueCount
      UserControl.ForeColor = 0
      height = stp * mValue(cBar)
      Label1.Caption = mValue(cBar)
      With UserControl
         .ForeColor = &HFFFFFF
         .CurrentX = 71 + ((cBar - 1) * (width)) - ((Label1.width - width) / 2) - (mSpace / 2)
         .CurrentY = .ScaleHeight - 59 - height - Label1.height
         UserControl.Print mValue(cBar)
         .ForeColor = &H88
         .CurrentX = 70 + ((cBar - 1) * (width)) - ((Label1.width - width) / 2) - (mSpace / 2)
         .CurrentY = .ScaleHeight - 60 - height - Label1.height
         UserControl.Print mValue(cBar)
      End With 'UserControl
     ElseIf cBar = 0 And BarEvidenziata <> cBar Then 'NOT CBAR...
      BarEvidenziata = 0
      Draw
   End If
   crsLabel.Visible = Line1.Visible

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   Redraw = False
   With PropBag
      UserControl.BackColor = .ReadProperty("BackColor", &HAAE7FF)
      mMainBarColor = .ReadProperty("MainBarColor", &HDA8448)
      ValueCount = .ReadProperty("ValueCount", 10)
      mHighLight = .ReadProperty("HighLight", &HFF)
      mBgLineVisible = .ReadProperty("BgLine", True)
      mSpace = .ReadProperty("Space", 8)
   End With 'PropBag
   Redraw = True

End Sub

Private Sub UserControl_Resize()

   Line1.x1 = 45
   Line1.x2 = UserControl.width
   Draw

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BackColor", UserControl.BackColor
      .WriteProperty "MainBarColor", mMainBarColor
      .WriteProperty "ValueCount", ValueCount
      .WriteProperty "Highlight", mHighLight
      .WriteProperty "BgLine", mBgLineVisible
      .WriteProperty "Space", mSpace
   End With 'PropBag

End Sub

Public Property Get Value(ByVal index As Integer) As Double

   If index > ValueCount And index <= 0 Then Exit Property
   Value = mValue(index)

End Property

Public Property Let Value(ByVal index As Integer, _
                          ByVal NewVal As Double)

   mValue(index) = NewVal
   Draw

End Property

Public Property Get ValueCount() As Integer

   On Error Resume Next
   ValueCount = UBound(mValue)
   If Err Then
      ValueCount = 10
   End If
   On Error GoTo 0
End Property

Public Property Let ValueCount(ByVal NewVal As Integer)

   Dim old As Long
   Dim x   As Long

   On Error Resume Next
   old = UBound(mValue)
   If old = 0 Then
      old = 1
   End If
   ReDim Preserve mValue(1 To NewVal)
   ReDim Preserve mLabel(1 To NewVal)
   ReDim Preserve mBarColor(1 To NewVal)
   For x = old To NewVal
      mBarColor(x) = -1
   Next x
   If Ambient.UserMode = False Then
     Draw
   End If
   On Error GoTo 0

End Property

Private Sub Reflex(ByVal x As Long, _
                  ByVal y As Long, _
                  ByVal lngWidth As Long, _
                  ByVal lngHeight As Long, _
                  ByVal ucolor As Long)

   Dim TriVert(3) As TRIVERTEX
   Dim gTRi(2)    As GRADIENT_TRIANGLE

   ucolor = GetColorLvl(ucolor, 8)
   'Fronte
   TriVert(0).x = x
   TriVert(0).y = y
   TriVert(1).x = x + lngWidth
   TriVert(1).y = y
   TriVert(2).x = x
   TriVert(2).y = y + IIf(lngHeight > 35, 35, lngHeight)
   TriVert(3).x = x + lngWidth
   TriVert(3).y = y + IIf(lngHeight > 35, 35, lngHeight)
   GradientFillColor TriVert(0), GetColorLvl(ucolor, 5)
   GradientFillColor TriVert(1), GetColorLvl(ucolor, 6)
   GradientFillColor TriVert(2), GetPixel(UserControl.hdc, TriVert(2).x, TriVert(2).y)
   GradientFillColor TriVert(3), GetPixel(UserControl.hdc, TriVert(3).x, TriVert(3).y)
   With gTRi(0)
      .Vertex1 = 0
      .Vertex2 = 1
      .Vertex3 = 2
   End With 'gTRi(0)
   With gTRi(1)
      .Vertex1 = 1
      .Vertex2 = 2
      .Vertex3 = 3
   End With 'gTRi(1)
   GradientFillTriangle UserControl.hdc, TriVert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
   'side
   TriVert(1).x = x + lngWidth + 20
   TriVert(1).y = y - 20
   TriVert(0).x = x + lngWidth
   TriVert(0).y = y
   TriVert(2).x = x + lngWidth
   TriVert(2).y = y + IIf(lngHeight > 35, 35, lngHeight)
   TriVert(3).x = x + lngWidth + 20
   TriVert(3).y = y + IIf(lngHeight > 35, 35, lngHeight) - 20
   GradientFillColor TriVert(0), GetColorLvl(ucolor, 7)
   GradientFillColor TriVert(2), GetPixel(UserControl.hdc, TriVert(2).x, TriVert(2).y)
   GradientFillColor TriVert(3), GetPixel(UserControl.hdc, TriVert(3).x, TriVert(3).y)
   GradientFillColor TriVert(1), GetColorLvl(ucolor, 8)
   With gTRi(0)
      .Vertex1 = 0
      .Vertex2 = 1
      .Vertex3 = 2
   End With 'gTRi(0)
   With gTRi(1)
      .Vertex1 = 1
      .Vertex2 = 2
      .Vertex3 = 3
   End With 'gTRi(1)
   GradientFillTriangle UserControl.hdc, TriVert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE

End Sub

':)Code Fixer V3.0.9 (21/11/2008 17.28.04) 31 + 715 = 746 Lines Thanks Ulli for inspiration and lots of code.

