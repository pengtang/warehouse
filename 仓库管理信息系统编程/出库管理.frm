VERSION 5.00
Begin VB.Form 出库管理 
   Caption         =   "出库管理"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   12645
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   8400
      TabIndex        =   13
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "查找"
      Height          =   735
      Left            =   8640
      TabIndex        =   12
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   735
      Left            =   8280
      TabIndex        =   7
      Top             =   7920
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出库"
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   7680
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   10575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "货物编号"
      Height          =   180
      Left            =   3000
      TabIndex        =   16
      Top             =   5520
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "查询关键字"
      Height          =   180
      Left            =   9120
      TabIndex        =   14
      Top             =   5640
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "出库时间"
      Height          =   180
      Left            =   4800
      TabIndex        =   11
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "出库货品数量"
      Height          =   180
      Left            =   1200
      TabIndex        =   10
      Top             =   6600
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "出库货品名称"
      Height          =   180
      Left            =   4800
      TabIndex        =   9
      Top             =   5520
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "出库编号"
      Height          =   180
      Left            =   1200
      TabIndex        =   8
      Top             =   5520
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "出库表"
      Height          =   180
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "出库管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

merchantOutNumber(k) = Text1.Text
merchantOutName(k) = Text2.Text
merchantOutQuantity(k) = Val(Text3.Text)
merchantOutTime(k) = Text4.Text
goodsOutNumber(k) = Text6.Text



Target = goodsOutNumber(k)
flag = 0 '检查是否出完
check = 0 '检查是否有这一项
For iter = 0 To j
pos = InStr(goodsTotalNumber(iter), Target)
If pos > 0 Then
    check = 1
    If merchantTotalQuantity(iter) - merchantOutQuantity(k) >= 0 Then
        List1.AddItem (Text1.Text + Space(20 - Len(Text1.Text)) + Text6.Text + Space(20 - Len(Text6.Text)) + Text2.Text + Space(22 - Len(Text2.Text)) + Text3.Text + Space(22 - Len(Text3.Text)) + Text4.Text)
        merchantTotalQuantity(iter) = merchantTotalQuantity(iter) - merchantOutQuantity(k)
        MsgBox "出库成功"
        k = k + 1
        If merchantTotalQuantity(iter) = 0 Then
            flag = 1
        End If
        Exit For
    Else
        MsgBox "出错,出库失败,出库数量大于库存总量"
        Exit For
    End If
End If

Next iter


If check = 0 Then
    MsgBox "出库失败,没有这一项"
End If

If flag = 1 Then
    If pos <> j - 1 Then
        For iter = pos - 1 To j - 2
         merchantTotalPosition(iter) = merchantTotalPosition(iter + 1)
         merchantTotalName(iter) = merchantTotalName(iter + 1)
         merchantTotalQuantity(iter) = merchantTotalQuantity(iter + 1)
         goodsTotalNumber(iter) = goodsTotalNumber(iter + 1)
        Next iter
    End If

j = j - 1
End If








End Sub

Private Sub Command2_Click()
出库管理.Hide
End Sub

Private Sub Command3_Click()
Target = Text5.Text
flag = 0
For iter = 0 To List1.ListCount - 1

pos = InStr(List1.List(iter), Target)
If pos > 0 Then
List1.Selected(iter) = True
MsgBox ("找到了! :)")
flag = 1
'Exit For
End If

Next iter

If flag = 0 Then
MsgBox ("没有找到! T.T ")
End If


End Sub


Private Sub Form_Load()
List1.AddItem ("出库编号" + Space(12) + "货品编号" + Space(12) + "出库货品名称" + Space(12) + "出库货品数量" + Space(12) + "出库时间")
End Sub
