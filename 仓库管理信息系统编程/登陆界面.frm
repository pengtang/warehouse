VERSION 5.00
Begin VB.Form ��½���� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��½����"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8655
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "��Բ"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6030
      Left            =   0
      Picture         =   "��½����.frx":0000
      ScaleHeight     =   5970
      ScaleWidth      =   11685
      TabIndex        =   0
      Top             =   0
      Width           =   11745
      Begin VB.CommandButton Command2 
         Caption         =   "�˳�"
         Height          =   495
         Left            =   4800
         TabIndex        =   7
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��½"
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   4680
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "��½���룺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   3
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "�û�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   2
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SP��˾�ֿ������Ϣϵͳ"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   5220
      End
   End
End
Attribute VB_Name = "��½����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Text1.Text = "" Then
      MsgBox "�������û���!", vbOKOnly + vbExclamation, "����"
      Exit Sub
    Else
    If Text1.Text = "admin" And Text2.Text = "admin" Then
    MsgBox "��ӭʹ�ñ�ϵͳ", vbOKOnly + vbInformation, "�ֿ������Ϣϵͳ"
    �ֿ������Ϣϵͳ.Show
    Else
    
    MsgBox "�û������������", vbOKOnly + vbCritical, "����"
    Text1.Text = ""
    Text2.Text = ""
   End If
  End If
End Sub

Private Sub Command2_Click()
End
End Sub

