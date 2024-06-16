VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "MainForm"
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Œ¢»Ì—≈∫⁄"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   645
   ScaleWidth      =   3855
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.Timer Timer1 
      Interval        =   61999
      Left            =   120
      Top             =   600
   End
   Begin VB.Label Text4 
      BackStyle       =   0  'Transparent
      Caption         =   "÷‹"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   160
      Width           =   495
   End
   Begin VB.Label WeekText 
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Text3 
      BackStyle       =   0  'Transparent
      Caption         =   "ÃÏ"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   160
      Width           =   495
   End
   Begin VB.Label DayText 
      BackStyle       =   0  'Transparent
      Caption         =   "985"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1300
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Text2 
      BackStyle       =   0  'Transparent
      Caption         =   "÷–øº"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   420
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "æ‡"
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   160
      Width           =   435
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetPrivateProfileInt Lib "kernel32" _
Alias "GetPrivateProfileIntA" ( _
     ByVal lpApplicationName As String, _
     ByVal lpKeyName As String, _
     ByVal nDefault As Long, _
     ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" ( _
     ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpDefault As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" ( _
     ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, _
     ByVal lpString As Any, _
     ByVal lpFileName As String) As Long
Private startX, startY As Integer
Public configINI As String
Public configDate As Date
'rw ini file: https://blog.csdn.net/surro/article/details/1751905

Public Function INIWriteString(filePath As String, Section As String, key As String, Value As String) As Boolean
     INIWriteString = (WritePrivateProfileString(Section, key, Value, filePath) = 0)
End Function

Public Function INIReadString(filePath As String, Section As String, key As String, Size As Long) As String
     Dim ReturnStr As String
     Dim ReturnLng As Long
     ReadString = vbNullString
     ReturnStr = Space(Size)
     ReturnLng = GetPrivateProfileString(Section, key, vbNullString, ReturnStr, Size, filePath)
     INIReadString = Left(ReturnStr, ReturnLng)
End Function



Private Sub Form_Load()
  Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '∂•÷√
  Call SetWindowLong(Me.hwnd, -20, True)  '∂•÷√
  Call SetLayeredWindowAttributes(hwnd, 0, 180, &H2)  '∞ÎÕ∏
  Call SetClassLong(Me.hwnd, -26, &H20000) '“ı”∞
  Me.Left = Screen.Width / 2 - Me.Width / 2
  Me.Top = 200
  
  configINI = App.Path & "\config.ini"
  configDate = CDate("2023-6-23")
  
  If (Dir(configINI) = "") Then
    ' ini is missing
    Call updateDate(configDate)
    Call saveConfig(configDate)
  Else
    Dim date1 As String
    date1 = INIReadString(configINI, "Config", "Date", 128)
    configDate = CDate(date1)
    Call updateDate(configDate)
    Board.DateInput.Text = date1
  End If
End Sub

Public Function updateDate(date1 As Date)
  DayText.Caption = DateDiff("d", Now, date1)
  WeekText.Caption = DateDiff("w", Now, date1)
End Function

Public Function saveConfig(date1 As Date)
    configDate = date1
    Call INIWriteString(configINI, "Config", "Date", Format(date1, "yyyy-mm-dd"))
End Function

Private Sub Timer1_Timer()
    If (Day(Now) <> Day(configDate)) Then Call updateDate(configDate)
End Sub

'shit dragging START ======================================================================
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MainForm.Left = MainForm.Left + (X - startX)
    MainForm.Top = MainForm.Top + (Y - startY)
  End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MainForm.Left = MainForm.Left + (X - startX)
    MainForm.Top = MainForm.Top + (Y - startY)
  End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MainForm.Left = MainForm.Left + (X - startX)
    MainForm.Top = MainForm.Top + (Y - startY)
  End If
End Sub
Private Sub DayText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub DayText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MainForm.Left = MainForm.Left + (X - startX)
    MainForm.Top = MainForm.Top + (Y - startY)
  End If
End Sub

Private Sub Text3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MainForm.Left = MainForm.Left + (X - startX)
    MainForm.Top = MainForm.Top + (Y - startY)
  End If
End Sub


Private Sub WeekText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub WeekText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MainForm.Left = MainForm.Left + (X - startX)
    MainForm.Top = MainForm.Top + (Y - startY)
  End If
End Sub

Private Sub Text4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
  If Button = 2 Then Call Board.PosWith(X, Y)
End Sub

Private Sub Text4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    MainForm.Left = MainForm.Left + (X - startX)
    MainForm.Top = MainForm.Top + (Y - startY)
  End If
End Sub
' shit dragging END


