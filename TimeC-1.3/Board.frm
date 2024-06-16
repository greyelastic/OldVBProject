VERSION 5.00
Begin VB.Form Board 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "面板"
   ClientHeight    =   1335
   ClientLeft      =   2505
   ClientTop       =   4395
   ClientWidth     =   2130
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   2130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton SmallerButton 
      Appearance      =   0  'Flat
      Caption         =   "-"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton LargerButton 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton CloseButton 
      Appearance      =   0  'Flat
      Caption         =   "关闭"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label SizeScaleText 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label SizeText 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "大小调节"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Version 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1.3"
      ForeColor       =   &H80000011&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   270
   End
End
Attribute VB_Name = "Board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private startX, startY As Integer

Public Sub PosWith(mx, my)
  Me.Left = Main.Left
  Me.Top = Main.Top + Main.Height + 50
  Me.Visible = Not (Me.Visible)
End Sub

Private Sub CloseButton_Click()
    End
End Sub

Private Sub Form_Load()
  Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3) '顶置
  Call SetWindowLong(Me.hwnd, -20, True)  '顶置
  Call SetLayeredWindowAttributes(hwnd, 0, 180, &H2)  '半透
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  startX = X
  startY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    Me.Left = Me.Left + (X - startX)
    Me.Top = Me.Top + (Y - startY)
  End If
End Sub


Private Sub LargerButton_Click()
    Main.ScaleSize = Main.ScaleSize + 1
    Call Main.updateScale
    SizeScaleText.Caption = Main.ScaleSize
End Sub

Private Sub SmallerButton_Click()
    If (Main.ScaleSize - 1 > 0) Then
        Main.ScaleSize = Main.ScaleSize - 1
        Call Main.updateScale
        SizeScaleText.Caption = Main.ScaleSize
    End If
End Sub

Private Sub Version_Click()
  'AppLog = Array( _
  '"Beta1.0 : 起源版本" _
  ', "Beta1.1 : 添加窗口阴影，支持在文字上拖动" _
  ')
  'MargeLog = ""
  'For i = 0 To UBound(AppLog)
  '  MargeLog = MargeLog & AppLog(i) & vbCr
  'Next
  'Call MsgBox("编程日志: " & vbCr & MargeLog)
 End Sub

