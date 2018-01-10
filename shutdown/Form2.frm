VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   LinkTopic       =   "Form2"
   ScaleHeight     =   1350
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label3 
      Caption         =   "取消"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "          在一分钟后关机"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "            系统已经断电。。。"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
 Form2.Left = (Screen.Width - Form1.Width) / 2
Form2.Top = (Screen.Height - Form1.Height) / 2
End Sub

Private Sub Label3_Click()
    Form1.Timer2.Enabled = False
    Shell "cmd.exe /c" & "shutdown.exe -a"
    Me.Hide
    Form1.Show
End Sub
