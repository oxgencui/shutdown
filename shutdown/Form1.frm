VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "定时关机"
   ClientHeight    =   2595
   ClientLeft      =   8115
   ClientTop       =   5070
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "隶书"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5250
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1560
      Top             =   240
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Form1.frx":2B434
      Left            =   480
      List            =   "Form1.frx":2B44A
      TabIndex        =   2
      Text            =   "选择关机时间或输入关机时间"
      Top             =   720
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消关机"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定关机"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Double '用来接收msgbox的返回值以便于判断是否设置了定时
Dim PwC As SYSTEM_POWER_STATUS
Private Sub Command1_Click()
    Dim commad As String
    Dim offtime As Double
   
    If Combo1.Text = "选择关机时间或输入关机时间" Then
        MsgBox "请选择时间！", vbCritical, "温馨提示"
    Else
        offtime = Val(Combo1.Text) * 60
        If offtime > 0 And offtime <= 18000 Then
            commad = "shutdown.exe -s -t " & offtime
            Shell "cmd.exe /c" & "shutdown.exe -a", vbHide
            Shell "cmd.exe /c" & commad, vbHide
            n = MsgBox("设置成功！将在" & Val(Combo1.Text) & "分钟后关机，谢谢使用！", , "温馨提示")
        Else
            MsgBox "请输入小于300的分钟数！", vbCritical, "温馨提示！"
        End If
    End If
End Sub

Private Sub Command2_Click()
    If n = 1 Then
        Shell "cmd.exe /c" & "shutdown.exe -a", vbHide
        MsgBox "取消关机！", vbApplicationModal, "温馨提示！"
    Else
        MsgBox "还没有定时关机！", , "温馨提示！"
    End If
End Sub

Private Sub Form_Load()
   Form1.Left = (Screen.Width - Form1.Width) / 2
   Form1.Top = (Screen.Height - Form1.Height) / 2
   Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Form1.Caption = "定时关机！当前时间:   " & Time
End Sub

Private Sub Timer2_Timer()
    
    GetSystemPowerStatus PwC
'Print "电池接通交流电元：" & PwC.ACLineStatus '1代表为交流电，0代表电池
If PwC.ACLineStatus = 0 Then
    Form1.Hide
    Form2.Show
    Shell "cmd.exe /c" & "shutdown.exe -s -t 60", vbHide
ElseIf PwC.ACLineStatus = 1 Then
    Timer2.Enabled = True
End If
End Sub
