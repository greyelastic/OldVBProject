VERSION 5.00
Begin VB.Form SettingForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4125
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4125
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Apply 
      Caption         =   "Ӧ��"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Frame F4 
      Caption         =   "����"
      Height          =   2175
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   1935
      Begin VB.CheckBox CheckOthers 
         Caption         =   "�м��ʳ�������"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.Frame F3 
      Caption         =   "����"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      Begin VB.TextBox others 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Text            =   "��̨"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame F2 
      Caption         =   "����"
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   855
      Begin VB.TextBox SettingCol 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "7"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame F1 
      Caption         =   "����"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
      Begin VB.TextBox SettingLine 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   "8"
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Version 
      Caption         =   "Beta1.8"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "SettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Ӧ������
Private Sub Apply_Click()
  Call Main.UnloadP(Main.Line0, Main.Col0)
  Main.SelectLine = 1
  Main.SelectCol = 1
  Main.Line0 = SettingLine.Text
  Main.Col0 = SettingCol.Text
  Call Main.LoadP(Main.Line0, Main.Col0)
  
  Call MsgBox("��Ӧ��", 0, "����")
End Sub

'��������һ���
Public Sub ex()
  End
End Sub

'���ڶ��ú����ð汾��
Private Sub Form_Load()
  Call SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 3)
  Version.Caption = "Beta1.8"
End Sub

'��־
Private Sub Version_Click()
  AppLog = Array( _
  "Beta1.0 : ����UI" _
  , "Beta1.5 : ������λͼ�λ�" _
  , "Beta1.6 : �޸����������������ü�����ж�� - 2020.12.4" _
  , "Beta1.7 : �޸����һ�л����һ���޷������е�BUG�������ڶ��õ������ÿ��޷���ʾ��BUG - 2020.12.3" _
  , "Beta1.8 : �Ż����룬����ע�ͣ���־�Ż�Ϊ����ʽ - 2020.12.14" _
  )
  MargeLog = ""
  For i = 0 To UBound(AppLog)
    MargeLog = MargeLog & AppLog(i) & vbCr
  Next
  Call MsgBox("�����־: " & vbCr & MargeLog)
  'Call MsgBox("�����־: " & vbCr & "Beta1.0 : ����UI" & vbCr & "Beta1.5 : ������λͼ�λ�" & vbCr & "Beta1.6 : �޸����������������ü�����ж�� - 2020.12.4" & vbCr & "Beta1.7 : �޸����һ�л����һ���޷������е�BUG�������ڶ��õ������ÿ��޷���ʾ��BUG - 2020.12.3")
End Sub
