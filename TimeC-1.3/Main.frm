VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "TimeC"
   ClientHeight    =   600
   ClientLeft      =   465
   ClientTop       =   465
   ClientWidth     =   2400
   FillStyle       =   0  'Solid
   Icon            =   "Main.frx":0000
   ScaleHeight     =   600
   ScaleWidth      =   2400
   Begin VB.Timer TimeC_upd 
      Interval        =   990
      Left            =   5400
      Top             =   3480
   End
   Begin VB.Label TimeText 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00 ; 00 ; 00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2250
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private startX, startY As Integer
Public ScaleSize As Integer
Public OriginWidth, OriginHeight, OriginFontSize As Integer

Public Function updateScale()
    Main.TimeText.Font.Size = Main.OriginFontSize * Main.ScaleSize
    Main.Width = Main.OriginWidth * Main.ScaleSize
    Main.Height = Main.OriginHeight * Main.ScaleSize
    TimeText.Left = Main.Width / 2 - TimeText.Width / 2
    TimeText.Top = Main.Height / 2 - TimeText.Height / 2
End Function
Private Sub Form_Load()
  Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '∂•÷√
  Call SetWindowLong(Me.hwnd, -20, True)  '∂•÷√
  Call SetLayeredWindowAttributes(hwnd, 0, 180, &H2)  '∞ÎÕ∏
  Call SetClassLong(Me.hwnd, -26, &H20000) '“ı”∞
  'init
  
  OriginWidth = 2400
  OriginHeight = 600
  OriginFontSize = 24
  
  Main.Width = OriginWidth
  Main.Height = OriginHeight
  TimeText.Font.Size = OriginFontSize
  TimeText.Caption = "00 : 00 : 00"
  TimeText.Left = Main.Width / 2 - TimeText.Width / 2
  TimeText.Top = Main.Height / 2 - TimeText.Height / 2
  'TimeText.FontName = ("impact")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub TimeC_upd_Timer()
  h = Hour(Now)
  m = Minute(Now)
  s = Second(Now)
  If Len(h) = 1 Then
    h = "0" & h
  End If
  If Len(m) = 1 Then
    m = "0" & m
  End If
  If Len(s) = 1 Then
    s = "0" & s
  End If
  
  TimeText.Caption = h & " : " & m & " : " & s
  TimeText.Left = Main.Width / 2 - TimeText.Width / 2
  TimeText.Top = Main.Height / 2 - TimeText.Height / 2
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    Main.Left = Main.Left + (X - startX)
    Main.Top = Main.Top + (Y - startY)
  End If
End Sub

Private Sub TimeText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub TimeText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    Main.Left = Main.Left + (X - startX)
    Main.Top = Main.Top + (Y - startY)
  End If
End Sub
