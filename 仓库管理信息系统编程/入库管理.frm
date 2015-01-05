VERSION 5.00
Begin VB.Form 入库管理 
   AutoRedraw      =   -1  'True
   Caption         =   "入库管理"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12765
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   3360
      TabIndex        =   17
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8760
      TabIndex        =   12
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "查找"
      Height          =   735
      Left            =   8760
      TabIndex        =   11
      Top             =   6960
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   7320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   7320
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "入库"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   8160
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   615
      Left            =   8760
      TabIndex        =   1
      Top             =   7800
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   5460
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   11295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "货物编号"
      Height          =   180
      Left            =   3840
      TabIndex        =   16
      Top             =   5880
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "入库位置"
      Height          =   180
      Left            =   6240
      TabIndex        =   15
      Top             =   6960
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "查找关键字"
      Height          =   180
      Left            =   9360
      TabIndex        =   13
      Top             =   5880
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "入库时间"
      Height          =   180
      Left            =   4080
      TabIndex        =   10
      Top             =   6960
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "入库货品数量"
      Height          =   180
      Left            =   1200
      TabIndex        =   9
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "入库货品名称"
      Height          =   180
      Left            =   6120
      TabIndex        =   8
      Top             =   5880
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "入库编号"
      Height          =   180
      Left            =   1320
      TabIndex        =   7
      Top             =   5880
      Width           =   720
   End
End
Attribute VB_Name = "入库管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
入库管理.Hide
End Sub

Private Sub Command2_Click()

merchantInNumber(i) = Text1.Text
goodsInNumber(i) = Text7.Text
merchantInName(i) = Text2.Text
merchantInQuantity(i) = Val(Text3.Text)
merchantInTime(i) = Text4.Text
merchantInPosition(i) = Text6.Text


List1.AddItem (Text1.Text + Space(20 - Len(Text1.Text)) + Text7.Text + Space(20 - Len(Text7.Text)) + Text2.Text + Space(24 - Len(Text2.Text)) + Text3.Text + Space(24 - Len(Text3.Text)) + Text4.Text + Space(24 - Len(Text4.Text)) + Text6.Text)


MsgBox "入库成功"

Target = goodsInNumber(i)
flag = 0
For iter = 0 To j
    pos = InStr(goodsTotalNumber(iter), Target)
    If pos > 0 Then
        merchantTotalQuantity(iter) = merchantTotalQuantity(iter) + merchantInQuantity(i)
        flag = 1
    Exit For
    End If

Next iter

If flag = 0 Then
goodsTotalNumber(j) = Text7.Text
merchantTotalName(j) = Text2.Text
merchantTotalQuantity(j) = Val(Text3.Text)
merchantTotalPosition(j) = Text6.Text

j = j + 1
End If

i = i + 1
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
List1.AddItem ("入库编号" + Space(12) + "货物编号" + Space(12) + "入库货物名称" + Space(12) + "入库货物数量" + Space(12) + "入库货物日期" + Space(12) + "入库位置")
End Sub
