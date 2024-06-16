VERSION 5.00
Begin VB.Form Main 
   Caption         =   "随机抽取"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5835
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton toSelect 
      Caption         =   "抽选"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   120
      Picture         =   "Main.frx":0000
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton GoToSetting 
      Caption         =   "设置"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label PlaceValue 
      AutoSize        =   -1  'True
      Caption         =   "行：   列：   "
      ForeColor       =   &H80000011&
      Height          =   180
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'窗口顶置用到的函数
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'定义宽与高
Public Line0, Col0 As Integer
'定义初始宽与高
Public PlaceLeft, PlaceTop As Integer
'定义抽选时的宽与高
Public SelectLine, SelectCol As Integer

'随机数函数
Function RndInt(m, n)
  RndInt = Int(Rnd * (m - n + 1) + n)
End Function

'初始化
Private Sub Form_Load()
  '初始化宽高
  Line0 = SettingForm.SettingLine.Text
  Col0 = SettingForm.SettingCol.Text
  '初始化起始坐标
  PlaceLeft = 120
  PlaceTop = 720
  '初始化选择
  SelectLine = 1
  SelectCol = 1
  '加载座位
  Call LoadP(Line0, Col0)
  '顶置窗口
  Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
  '初始化随机数
  Randomize
End Sub

'前往设置窗
Private Sub GoToSetting_Click()
  SettingForm.Show
End Sub

'带上设置窗一起关
Private Sub Form_Unload(Cancel As Integer)
  SettingForm.ex
End Sub

'加载座位P
Public Sub LoadP(Line, Col)
  NO = 1
  P(0).Left = PlaceLeft
  P(0).Top = PlaceTop
  P(0).Picture = LoadPicture(App.Path & "\PlaceA.jpg")
  For i2 = 1 To Line
  For i = 1 To Col
    Load P(NO)
    With P(NO)
    .Picture = P(0).Picture
    .Width = P(0).Width
    .Height = P(0).Height
    .Left = P(0).Left + (P(0).Width + 50) * i
    .Top = P(0).Top
    .Visible = True
    End With
    NO = NO + 1
  Next
  P(0).Top = P(0).Top + P(0).Height + 50
  Next
End Sub

'卸载座位P
Public Sub UnloadP(Line, Col)
  For i = 1 To Line * Col
    Unload P(i)
  Next
End Sub

'抽取座位
Private Sub toSelect_Click()
  '追加取消上个抽选的座位
  If SelectLine = 1 Then
    P(SelectCol).Picture = LoadPicture(App.Path & "\PlaceA.jpg")
  Else
    P(SelectCol + (SelectLine - 1) * Col0).Picture = LoadPicture(App.Path & "\PlaceA.jpg")
  End If
  '随机抽选
  SelectLine = RndInt(0, Line0 + 1)
  SelectCol = RndInt(0, Col0 + 1)
  '显示文本
  PlaceValue.Caption = "行：" & SelectLine & "  列：" & SelectCol
  '追加显示被抽选的座位
  If SelectLine = 1 Then
    P(SelectCol).Picture = LoadPicture(App.Path & "\PlaceB.jpg")
  Else
    P(SelectCol + (SelectLine - 1) * Col0).Picture = LoadPicture(App.Path & "\PlaceB.jpg")
  End If
End Sub
