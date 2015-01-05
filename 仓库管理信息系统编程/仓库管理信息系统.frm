VERSION 5.00
Begin VB.Form 仓库管理信息系统 
   Caption         =   "仓库管理信息系统"
   ClientHeight    =   6240
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10410
   LinkTopic       =   "Form2"
   Picture         =   "仓库管理信息系统.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   10410
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu M1 
      Caption         =   "基本信息设置"
      Begin VB.Menu A1 
         Caption         =   "供应商信息管理"
      End
      Begin VB.Menu A2 
         Caption         =   "客户信息管理"
      End
      Begin VB.Menu A3 
         Caption         =   "管理员信息管理"
      End
   End
   Begin VB.Menu M2 
      Caption         =   "货品管理"
      Begin VB.Menu B1 
         Caption         =   "入库管理"
      End
      Begin VB.Menu B2 
         Caption         =   "出库管理"
      End
      Begin VB.Menu B3 
         Caption         =   "库存盘点"
      End
   End
End
Attribute VB_Name = "仓库管理信息系统"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A1_Click()
供应商信息管理.Show
End Sub

Private Sub A2_Click()
客户信息管理.Show
End Sub

Private Sub A3_Click()
管理员信息管理.Show
End Sub

Private Sub B1_Click()
入库管理.Show
End Sub

Private Sub B2_Click()
出库管理.Show
End Sub

Private Sub B3_Click()
库存盘点.Show
End Sub

Private Sub D1_Click()
入库报表.Show
End Sub
