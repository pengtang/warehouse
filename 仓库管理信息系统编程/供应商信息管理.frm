VERSION 5.00
Begin VB.Form ��Ӧ����Ϣ���� 
   Caption         =   "Form1"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16095
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   16095
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   7800
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5520
      TabIndex        =   12
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ɾ����Ӧ��"
      Height          =   495
      Left            =   13440
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3300
      ItemData        =   "��Ӧ����Ϣ����.frx":0000
      Left            =   480
      List            =   "��Ӧ����Ϣ����.frx":0002
      TabIndex        =   3
      Top             =   840
      Width           =   12375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   855
      Left            =   13560
      TabIndex        =   2
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���빩Ӧ��"
      Height          =   735
      Left            =   3840
      TabIndex        =   1
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ�̵绰"
      Height          =   180
      Left            =   8280
      TabIndex        =   10
      Top             =   5520
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ�̵�ַ"
      Height          =   180
      Left            =   5640
      TabIndex        =   9
      Top             =   5520
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ������"
      Height          =   180
      Left            =   3360
      TabIndex        =   8
      Top             =   5400
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ�̱��"
      Height          =   180
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ�̱��"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   5520
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "��Ӧ�̱�"
      Height          =   255
      Left            =   6120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "��Ӧ����Ϣ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.AddItem ("��Ӧ�̱��" + Space(20) + "��Ӧ������" + Space(20) + "��Ӧ�̵�ַ" + Space(20) + "��Ӧ�̵绰")
End Sub
Private Sub Command1_Click()


m = Text1.Text
n = Text2.Text
q = Text3.Text
w = Text4.Text

List1.AddItem (m + Space(30 - Len(m)) + n + Space(30 - Len(n)) + q + Space(30 - Len(q)) + w)
MsgBox "��ӳɹ�"
End Sub

Private Sub Command2_Click()
��Ӧ����Ϣ����.Hide
End Sub

Private Sub Command3_Click()


If List1.ListIndex >= 0 Then
    List1.RemoveItem (List1.ListIndex)
End If

End Sub

Private Sub Label5_Click()

End Sub
