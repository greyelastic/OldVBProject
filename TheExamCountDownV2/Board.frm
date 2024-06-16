VERSION 5.00
Begin VB.Form Board 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "面板"
   ClientHeight    =   1845
   ClientLeft      =   2505
   ClientTop       =   4395
   ClientWidth     =   4035
   Icon            =   "Board.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CloseButton 
      Appearance      =   0  'Flat
      Caption         =   "关闭"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton SaveButton 
      Appearance      =   0  'Flat
      Caption         =   "保存"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      Caption         =   "取消"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox DateInput 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label DateText 
      BackStyle       =   0  'Transparent
      Caption         =   "日期 (以 年-月-日 的方式填入，例如 2023-6-23)"
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label SettingText 
      BackStyle       =   0  'Transparent
      Caption         =   "设定"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Version 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1.0"
      ForeColor       =   &H80000011&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1560
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

Public Sub PosWith(mx, my)
  Me.Left = MainForm.Left
  Me.Top = MainForm.Top + MainForm.Height + 20
  Me.Visible = Not (Me.Visible)
End Sub

Private Sub CancelButton_Click()
    Me.Visible = False
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

Private Sub SaveButton_Click()
    If IsDate(DateInput.Text) Then
        Dim date1 As Date
        date1 = CDate(DateInput.Text)
        Call MainForm.updateDate(date1)
        Call MainForm.saveConfig(date1)
        Me.Visible = False
    Else
        Call MsgBox("格式错误")
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

