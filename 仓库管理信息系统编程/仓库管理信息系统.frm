VERSION 5.00
Begin VB.Form �ֿ������Ϣϵͳ 
   Caption         =   "�ֿ������Ϣϵͳ"
   ClientHeight    =   6240
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10410
   LinkTopic       =   "Form2"
   Picture         =   "�ֿ������Ϣϵͳ.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   10410
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu M1 
      Caption         =   "������Ϣ����"
      Begin VB.Menu A1 
         Caption         =   "��Ӧ����Ϣ����"
      End
      Begin VB.Menu A2 
         Caption         =   "�ͻ���Ϣ����"
      End
      Begin VB.Menu A3 
         Caption         =   "����Ա��Ϣ����"
      End
   End
   Begin VB.Menu M2 
      Caption         =   "��Ʒ����"
      Begin VB.Menu B1 
         Caption         =   "������"
      End
      Begin VB.Menu B2 
         Caption         =   "�������"
      End
      Begin VB.Menu B3 
         Caption         =   "����̵�"
      End
   End
End
Attribute VB_Name = "�ֿ������Ϣϵͳ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A1_Click()
��Ӧ����Ϣ����.Show
End Sub

Private Sub A2_Click()
�ͻ���Ϣ����.Show
End Sub

Private Sub A3_Click()
����Ա��Ϣ����.Show
End Sub

Private Sub B1_Click()
������.Show
End Sub

Private Sub B2_Click()
�������.Show
End Sub

Private Sub B3_Click()
����̵�.Show
End Sub

Private Sub D1_Click()
��ⱨ��.Show
End Sub
