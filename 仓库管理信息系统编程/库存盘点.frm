VERSION 5.00
Begin VB.Form ����̵� 
   Caption         =   "����̵�"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10005
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   855
      Left            =   4920
      TabIndex        =   2
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ˢ��"
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   6000
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   5280
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "����̵�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
List1.AddItem ("��Ʒ���" + Space(25) + "��Ʒ����" + Space(25) + "��Ʒ����" + Space(25) + "�ڿ�λ��")
For iter = 0 To j - 1
List1.AddItem (goodsTotalNumber(iter) + Space(33 - Len(goodsTotalNumber(iter))) + merchantTotalName(iter) + Space(33 - Len(merchantTotalName(iter))) + Str(merchantTotalQuantity(iter)) + Space(33 - Len(Str(merchantTotalQuantity(iter)))) + goodsTotalPosition)
Next iter


End Sub

Private Sub Command2_Click()
����̵�.Hide
End Sub

Private Sub Form_Load()
List1.AddItem ("��Ʒ���" + Space(25) + "��Ʒ����" + Space(25) + "��Ʒ����" + Space(25) + "�ڿ�λ��")
End Sub
