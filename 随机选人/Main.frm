VERSION 5.00
Begin VB.Form Main 
   Caption         =   "�����ȡ"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5835
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton toSelect 
      Caption         =   "��ѡ"
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
      Caption         =   "����"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label PlaceValue 
      AutoSize        =   -1  'True
      Caption         =   "�У�   �У�   "
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
'���ڶ����õ��ĺ���
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'��������
Public Line0, Col0 As Integer
'�����ʼ�����
Public PlaceLeft, PlaceTop As Integer
'�����ѡʱ�Ŀ����
Public SelectLine, SelectCol As Integer

'���������
Function RndInt(m, n)
  RndInt = Int(Rnd * (m - n + 1) + n)
End Function

'��ʼ��
Private Sub Form_Load()
  '��ʼ�����
  Line0 = SettingForm.SettingLine.Text
  Col0 = SettingForm.SettingCol.Text
  '��ʼ����ʼ����
  PlaceLeft = 120
  PlaceTop = 720
  '��ʼ��ѡ��
  SelectLine = 1
  SelectCol = 1
  '������λ
  Call LoadP(Line0, Col0)
  '���ô���
  Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
  '��ʼ�������
  Randomize
End Sub

'ǰ�����ô�
Private Sub GoToSetting_Click()
  SettingForm.Show
End Sub

'�������ô�һ���
Private Sub Form_Unload(Cancel As Integer)
  SettingForm.ex
End Sub

'������λP
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

'ж����λP
Public Sub UnloadP(Line, Col)
  For i = 1 To Line * Col
    Unload P(i)
  Next
End Sub

'��ȡ��λ
Private Sub toSelect_Click()
  '׷��ȡ���ϸ���ѡ����λ
  If SelectLine = 1 Then
    P(SelectCol).Picture = LoadPicture(App.Path & "\PlaceA.jpg")
  Else
    P(SelectCol + (SelectLine - 1) * Col0).Picture = LoadPicture(App.Path & "\PlaceA.jpg")
  End If
  '�����ѡ
  SelectLine = RndInt(0, Line0 + 1)
  SelectCol = RndInt(0, Col0 + 1)
  '��ʾ�ı�
  PlaceValue.Caption = "�У�" & SelectLine & "  �У�" & SelectCol
  '׷����ʾ����ѡ����λ
  If SelectLine = 1 Then
    P(SelectCol).Picture = LoadPicture(App.Path & "\PlaceB.jpg")
  Else
    P(SelectCol + (SelectLine - 1) * Col0).Picture = LoadPicture(App.Path & "\PlaceB.jpg")
  End If
End Sub
