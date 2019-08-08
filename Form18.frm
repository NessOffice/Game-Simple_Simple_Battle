VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form18 
   Caption         =   "-4级战斗游戏"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   9150
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      Height          =   1575
      Left            =   15000
      TabIndex        =   109
      Top             =   1800
      Width           =   615
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   1080
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   10
      Tab             =   4
      TabsPerRow      =   10
      TabHeight       =   529
      BackColor       =   16777215
      TabCaption(0)   =   "战斗"
      TabPicture(0)   =   "Form18.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command23"
      Tab(0).Control(1)=   "Command22"
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(3)=   "Label27"
      Tab(0).Control(4)=   "Label44"
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(12)=   "Label4"
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(15)=   "Label22(0)"
      Tab(0).Control(16)=   "Label22(1)"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "状态"
      TabPicture(1)   =   "Form18.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label22(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label34"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label37"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label38"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label39"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label40"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame4"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "修炼"
      TabPicture(2)   =   "Form18.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command2(2)"
      Tab(2).Control(1)=   "Command2(1)"
      Tab(2).Control(2)=   "Command2(0)"
      Tab(2).Control(3)=   "Command5"
      Tab(2).Control(4)=   "Label33"
      Tab(2).Control(5)=   "Label32"
      Tab(2).Control(6)=   "Label31"
      Tab(2).Control(7)=   "Label30"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "副本"
      TabPicture(3)   =   "Form18.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(9)"
      Tab(3).Control(1)=   "Label1(11)"
      Tab(3).Control(2)=   "Label1(10)"
      Tab(3).Control(3)=   "Label1(8)"
      Tab(3).Control(4)=   "Label1(7)"
      Tab(3).Control(5)=   "Label1(6)"
      Tab(3).Control(6)=   "Label1(5)"
      Tab(3).Control(7)=   "Label1(4)"
      Tab(3).Control(8)=   "Label1(3)"
      Tab(3).Control(9)=   "Label1(2)"
      Tab(3).Control(10)=   "Label1(1)"
      Tab(3).Control(11)=   "Label1(0)"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "商店"
      TabPicture(4)   =   "Form18.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label20"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame1"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame2"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Command12"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Command13"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Command14"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Command15"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Command16"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Command17"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Command20"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Command21"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).ControlCount=   11
      TabCaption(5)   =   "系统"
      TabPicture(5)   =   "Form18.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label58"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label12"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label19"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label21"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "VIP"
      TabPicture(6)   =   "Form18.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame3"
      Tab(6).Control(1)=   "Label35"
      Tab(6).Control(2)=   "Label28"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "熔炼"
      TabPicture(7)   =   "Form18.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label22(3)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Label41"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Label42"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Label46"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Label47"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "Frame5"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "Frame6"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).ControlCount=   7
      TabCaption(8)   =   "成就"
      TabPicture(8)   =   "Form18.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label24"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Label18"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "任务"
      TabPicture(9)   =   "Form18.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Label14"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Label26"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "Label25"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "Label23"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).ControlCount=   4
      Begin VB.CommandButton Command2 
         Caption         =   "180秒"
         Height          =   375
         Index           =   2
         Left            =   -72000
         TabIndex        =   120
         Top             =   1380
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "60秒"
         Height          =   375
         Index           =   1
         Left            =   -73320
         TabIndex        =   119
         Top             =   1380
         Width           =   1095
      End
      Begin VB.CommandButton Command23 
         Caption         =   "使用生命药水"
         Height          =   495
         Left            =   -69000
         TabIndex        =   110
         Top             =   3060
         Width           =   1455
      End
      Begin VB.CommandButton Command22 
         Caption         =   "我方攻击"
         Height          =   495
         Left            =   -73800
         TabIndex        =   108
         Top             =   3060
         Width           =   1455
      End
      Begin VB.CommandButton Command21 
         Caption         =   " 购买浓厚石   5元"
         Height          =   495
         Left            =   5520
         TabIndex        =   101
         Top             =   3060
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command20 
         Caption         =   "购买武器提升卷25元"
         Height          =   495
         Left            =   3840
         TabIndex        =   100
         Top             =   3060
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Caption         =   "提升武器威力"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   83
         Top             =   900
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CommandButton Command19 
            Caption         =   "提升当前武器"
            Height          =   1815
            Left            =   2160
            TabIndex        =   98
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label57 
            Height          =   255
            Left            =   1200
            TabIndex        =   99
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label56 
            Caption         =   "浓厚石数量："
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label55 
            Caption         =   "武器提升卷数量："
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label54 
            Caption         =   "武器等级：LV"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label53 
            Caption         =   "当前武器："
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label52 
            Height          =   255
            Left            =   1080
            TabIndex        =   93
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "熔炼新的物品"
         Height          =   2175
         Left            =   -72000
         TabIndex        =   82
         Top             =   900
         Visible         =   0   'False
         Width           =   6015
         Begin VB.OptionButton Option2 
            Caption         =   "成就碎片"
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   107
            Top             =   1200
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "汲血刃"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   106
            Top             =   960
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "修炼卡×3+抽奖券×2"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   105
            Top             =   720
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "浓厚石×10"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   104
            Top             =   480
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            Caption         =   "生命药水"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   103
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton Command18 
            Caption         =   "合成"
            Height          =   495
            Left            =   600
            TabIndex        =   88
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label Label51 
            Caption         =   "D数量："
            Height          =   375
            Left            =   240
            TabIndex        =   92
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label50 
            Caption         =   "C数量："
            Height          =   375
            Left            =   240
            TabIndex        =   91
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label49 
            Caption         =   "B数量："
            Height          =   375
            Left            =   240
            TabIndex        =   90
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label48 
            Caption         =   "A数量："
            Height          =   375
            Left            =   240
            TabIndex        =   89
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command17 
         Caption         =   "购买金钱大炮1500元"
         Height          =   495
         Left            =   2160
         TabIndex        =   62
         Top             =   3060
         Width           =   1455
      End
      Begin VB.CommandButton Command16 
         Caption         =   "购买火箭筒（一次性）200元"
         Height          =   495
         Left            =   480
         TabIndex        =   58
         Top             =   3060
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         Caption         =   "购买金币刀800元"
         Height          =   495
         Left            =   4920
         TabIndex        =   55
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "切换武器"
         Height          =   2895
         Left            =   -72480
         TabIndex        =   54
         Top             =   780
         Width           =   6375
         Begin VB.OptionButton Option1 
            Caption         =   "VB6(你已无敌)"
            Height          =   615
            Index           =   7
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   1560
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "枪魚枪(每次胜利后攻击-7)"
            Height          =   615
            Index           =   6
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   600
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "复仇拳套(战败后攻击+5)"
            Height          =   615
            Index           =   5
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "汲血刃(战胜后生命+2%)"
            Height          =   615
            Index           =   4
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   1200
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "钢管(攻击+40)"
            Height          =   615
            Index           =   3
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "金钱大炮(攻击时消耗金币)"
            Enabled         =   0   'False
            Height          =   615
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   2040
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "火箭筒(8倍攻击)"
            Enabled         =   0   'False
            Height          =   615
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1200
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "金币刀(获得金币+20%)"
            Enabled         =   0   'False
            Height          =   615
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command14 
         Caption         =   "修炼加速卡80元"
         Height          =   495
         Left            =   3600
         TabIndex        =   53
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "购买抽奖次数200元"
         Height          =   495
         Left            =   2160
         TabIndex        =   52
         Top             =   2340
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "购买生命药水（增加40HP）85元"
         Height          =   495
         Left            =   480
         TabIndex        =   51
         Top             =   2340
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "20秒"
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   38
         Top             =   1380
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "使用修炼加速卡"
         Height          =   735
         Left            =   -74040
         TabIndex        =   35
         Top             =   2580
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "VIP专属区域"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -74760
         TabIndex        =   29
         Top             =   1020
         Width           =   7695
         Begin VB.CommandButton Command7 
            Caption         =   "抽奖"
            Enabled         =   0   'False
            Height          =   975
            Left            =   2160
            TabIndex        =   40
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            Caption         =   "VIP副本——   神之间    （千万不要按！！！）"
            Enabled         =   0   'False
            Height          =   975
            Left            =   480
            TabIndex        =   39
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "确定"
            Height          =   735
            Left            =   6360
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   4320
            TabIndex        =   31
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "经典黑"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   4800
            TabIndex        =   76
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "蕾咪红"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   75
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "八云紫"
            Enabled         =   0   'False
            ForeColor       =   &H00800080&
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   74
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "天依蓝"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   73
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "流乃白"
            Enabled         =   0   'False
            ForeColor       =   &H00FF8080&
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   72
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "恐龙染料"
            Enabled         =   0   'False
            ForeColor       =   &H000080FF&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   71
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "验证码："
            Height          =   255
            Left            =   4320
            TabIndex        =   30
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "70金币区"
         Height          =   1095
         Left            =   3360
         TabIndex        =   25
         Top             =   1020
         Width           =   2295
         Begin VB.CommandButton Command9 
            Caption         =   "生命+30"
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "攻击+3"
            Height          =   495
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "30金币区"
         Height          =   1095
         Left            =   360
         TabIndex        =   24
         Top             =   1020
         Width           =   2295
         Begin VB.CommandButton Command11 
            Caption         =   "攻击+1"
            Height          =   495
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton Command10 
            Caption         =   "生命+10"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   -74160
         TabIndex        =   13
         Text            =   "未命名（记得改）"
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "施工中，暂未启用"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   48
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74280
         TabIndex        =   118
         Top             =   1440
         Width           =   7695
      End
      Begin VB.Label Label23 
         Caption         =   "状态：未完成"
         Height          =   255
         Left            =   -74400
         TabIndex        =   117
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label25 
         Height          =   615
         Left            =   -73320
         TabIndex        =   116
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "领取奖励"
         Height          =   615
         Left            =   -71880
         TabIndex        =   115
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "1次"
         Height          =   255
         Left            =   -73320
         TabIndex        =   114
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label21 
         Caption         =   "原作者网盘(里面还有个1.5版的)：http://pan.baidu.com/share/home?uk=4163755544&view=share#category/type=0(希望没**)"
         Height          =   375
         Left            =   -74640
         TabIndex        =   113
         Top             =   2340
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.Label Label19 
         Caption         =   $"Form18.frx":0118
         Height          =   615
         Left            =   -74640
         TabIndex        =   112
         Top             =   1260
         Visible         =   0   'False
         Width           =   8535
      End
      Begin VB.Label Label12 
         Caption         =   "关于本战斗游戏……"
         Height          =   255
         Left            =   -74640
         TabIndex        =   111
         Top             =   1020
         Width           =   1815
      End
      Begin VB.Label Label58 
         Caption         =   "原作网址：http://tieba.baidu.com/p/3006282155(点击进入)"
         Height          =   255
         Left            =   -74640
         TabIndex        =   102
         Top             =   1980
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label Label47 
         Caption         =   "您的熔炼等级："
         Height          =   240
         Left            =   -74520
         TabIndex        =   86
         Top             =   3540
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label46 
         BackColor       =   &H00FFFF00&
         Height          =   345
         Left            =   -74280
         TabIndex        =   85
         Top             =   3840
         Visible         =   0   'False
         Width           =   7005
      End
      Begin VB.Label Label42 
         Height          =   240
         Left            =   -68520
         TabIndex        =   84
         Top             =   3540
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label41 
         Caption         =   "这里空空如也，什么也没有。"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   48
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   -74760
         TabIndex        =   81
         Top             =   1380
         Width           =   8655
      End
      Begin VB.Label Label40 
         Height          =   375
         Left            =   -74880
         TabIndex        =   80
         Top             =   2940
         Width           =   1815
      End
      Begin VB.Label Label39 
         Caption         =   "战斗武器："
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "HP："
         Height          =   255
         Left            =   -74880
         TabIndex        =   78
         Top             =   1380
         Width           =   1935
      End
      Begin VB.Label Label37 
         Height          =   240
         Left            =   -68760
         TabIndex        =   77
         Top             =   3660
         Width           =   1815
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFF00&
         Height          =   345
         Left            =   -74520
         TabIndex        =   69
         Top             =   3960
         Visible         =   0   'False
         Width           =   7005
      End
      Begin VB.Label Label17 
         Caption         =   "小提示：如果火箭筒爆掉了记得要换武器！"
         Height          =   255
         Left            =   -74760
         TabIndex        =   59
         Top             =   3780
         Width           =   3495
      End
      Begin VB.Label Label27 
         Height          =   495
         Left            =   -69600
         TabIndex        =   48
         Top             =   3660
         Width           =   2415
      End
      Begin VB.Label Label44 
         Caption         =   "Label44"
         Height          =   15
         Left            =   -71280
         TabIndex        =   47
         Top             =   3900
         Width           =   375
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "男人的房间"
         Height          =   495
         Index           =   9
         Left            =   -74520
         TabIndex        =   44
         Top             =   3420
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "玉环中学"
         Height          =   495
         Index           =   11
         Left            =   -70080
         TabIndex        =   43
         Top             =   3420
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "某校宿舍"
         Height          =   495
         Index           =   10
         Left            =   -72240
         TabIndex        =   42
         Top             =   3420
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label35 
         Height          =   255
         Left            =   -69600
         TabIndex        =   41
         Top             =   780
         Width           =   2415
      End
      Begin VB.Label Label33 
         Caption         =   "可以减少50秒时间！"
         Height          =   255
         Left            =   -72600
         TabIndex        =   37
         Top             =   2940
         Width           =   1695
      End
      Begin VB.Label Label32 
         Caption         =   "您的加速卡数量"
         Height          =   255
         Left            =   -74400
         TabIndex        =   36
         Top             =   3420
         Width           =   1815
      End
      Begin VB.Label Label31 
         Caption         =   "剩余时间：请先操作"
         Height          =   255
         Left            =   -74520
         TabIndex        =   34
         Top             =   2100
         Width           =   1695
      End
      Begin VB.Label Label30 
         Caption         =   "选择修炼类型"
         Height          =   255
         Left            =   -74520
         TabIndex        =   33
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "提示：验证码打村上春树一长篇小说名(注意大小写)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   780
         Width           =   4215
      End
      Begin VB.Label Label24 
         Caption         =   "Label24"
         Height          =   15
         Left            =   -73680
         TabIndex        =   27
         Top             =   900
         Width           =   255
      End
      Begin VB.Label Label20 
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label16 
         Caption         =   "LV"
         Height          =   240
         Left            =   -74760
         TabIndex        =   23
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "攻击："
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   1860
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "未命名（记得改）"
         Height          =   255
         Left            =   -74520
         TabIndex        =   21
         Top             =   1095
         Width           =   1455
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   -69960
         TabIndex        =   20
         Top             =   2060
         Width           =   975
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   -68640
         TabIndex        =   19
         Top             =   2060
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "目前敌方HP:"
         Height          =   255
         Left            =   -70920
         TabIndex        =   18
         Top             =   2060
         Width           =   1935
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   -69960
         TabIndex        =   17
         Top             =   1100
         Width           =   975
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   -68640
         TabIndex        =   16
         Top             =   1100
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "目前我方HP:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -70920
         TabIndex        =   15
         Top             =   1100
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "姓名："
         Height          =   375
         Left            =   -74880
         TabIndex        =   14
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   -74400
         TabIndex        =   12
         Top             =   2460
         Width           =   7005
      End
      Begin VB.Label Label3 
         Caption         =   "暂无怪物"
         Height          =   255
         Left            =   -74520
         TabIndex        =   11
         Top             =   2055
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   -74400
         TabIndex        =   10
         Top             =   1500
         Width           =   7000
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "城管之家"
         Height          =   495
         Index           =   8
         Left            =   -70080
         TabIndex        =   9
         Top             =   2700
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "成都养鸡场2号"
         Height          =   495
         Index           =   7
         Left            =   -72240
         TabIndex        =   8
         Top             =   2700
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "中国食堂"
         Height          =   495
         Index           =   6
         Left            =   -74520
         TabIndex        =   7
         Top             =   2700
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "中环学玉"
         Height          =   495
         Index           =   5
         Left            =   -70080
         TabIndex        =   6
         Top             =   1860
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "成都养鸡场"
         Height          =   495
         Index           =   4
         Left            =   -72240
         TabIndex        =   5
         Top             =   1860
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "成都养鸡场"
         Height          =   495
         Index           =   3
         Left            =   -74520
         TabIndex        =   4
         Top             =   1860
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "庆澜海"
         Height          =   495
         Index           =   2
         Left            =   -70080
         TabIndex        =   3
         Top             =   1020
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "玉坎河"
         Height          =   495
         Index           =   1
         Left            =   -72240
         TabIndex        =   2
         Top             =   1020
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "果冻村"
         Height          =   495
         Index           =   0
         Left            =   -74520
         TabIndex        =   1
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   0
         Left            =   -74520
         TabIndex        =   67
         Top             =   2340
         Width           =   7215
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Index           =   1
         Left            =   -74520
         TabIndex        =   68
         Top             =   1380
         Width           =   7215
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Height          =   465
         Index           =   2
         Left            =   -74640
         TabIndex        =   70
         Top             =   3900
         Width           =   7215
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Height          =   465
         Index           =   3
         Left            =   -74400
         TabIndex        =   87
         Top             =   3780
         Visible         =   0   'False
         Width           =   7215
      End
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim nowHP, totHP, sm, jiasu, choujiang, leftime, love, tisheng, nonghou, huojianjia, ronglianLV, lucky As Integer
Dim eneHP, teHP, teneHP(16) As Integer
Dim Bratk, enatk(16) As Integer
Dim enemy(15), mission(9) As String
Dim coin, nowEXP, ronglianEXP As Long
Dim nm, Greward(16), EXPreward(16), wa(7), wl(6) As Integer
Dim xiulian(2) As Boolean
Dim MouseKill, wangweiKILL, A, B, C, D, achip As Integer
Public NowEnemy As Integer
Dim w(7) As Boolean
Dim i, cont As Integer

Private Sub Command1_Click()
If Text2.Text = "1q84" Then
MsgBox "您成功获得VIP权限"
Command1.Enabled = False
Command6.Enabled = True
Command7.Enabled = True
    For i = 0 To 4
    Label36(i).Enabled = True
    Next i
ElseIf Text2.Text = "幺六二一就是屌" Then
MsgBox "输入的验证码不正确"
Option1(7).Visible = True
ElseIf Text2.Text = "打败零二死基佬" Then
MsgBox "输入的验证码不正确"
Option1(7).Visible = True
ElseIf Text2.Text = "成绩颜值与身高" Then
MsgBox "输入的验证码不正确"
nowHP = nowHP + 50000: totHP = totHP + 50000
ElseIf Text2.Text = "你们寝室算个鸟" Then
MsgBox "输入的验证码不正确"
Command6.Visible = True
Else
MsgBox "输入的验证码不正确"
End If
End Sub

Private Sub Command10_Click()
Text3.SetFocus
If coin < 30 Then
MsgBox "请先确认是否金币足够！"
ElseIf eneHP < teHP Then
MsgBox "请不要在战斗时逛商店……"
SSTab1.Tab = 0
Else
coin = coin - 30
nowHP = nowHP + 10
totHP = totHP + 10
End If
End Sub

Private Sub Command11_Click()
Text3.SetFocus
If coin < 30 Then
MsgBox "请先确认是否金币足够！"
ElseIf eneHP < teHP Then
MsgBox "请不要在战斗时逛商店……"
SSTab1.Tab = 0
Else
coin = coin - 30
Bratk = Bratk + 1
    If Label23.Caption = "状态：未完成" And nm = 1 Then     '任务2
    MsgBox "您的任务完成！！"
    Label23.Caption = "状态：已完成"
    End If
End If
End Sub

Private Sub Command12_Click()
Text3.SetFocus
If coin < 85 Then
MsgBox "请先确认是否金币足够！"
Else
coin = coin - 85
sm = sm + 1
End If
End Sub

Private Sub Command13_Click()
Text3.SetFocus
If coin < 200 Then
MsgBox "请先确认是否金币足够！"
Else
coin = coin - 200
choujiang = choujiang + 1
End If
End Sub

Private Sub Command14_Click()
Text3.SetFocus
If coin < 80 Then
MsgBox "请先确认是否金币足够！"
Else
coin = coin - 80
jiasu = jiasu + 1
End If
End Sub

Private Sub Command15_Click()
If coin < 800 Then
MsgBox "请先确认是否金币足够！"
Else
coin = coin - 800
Option1(0).Enabled = True
Command15.Enabled = False
End If
End Sub

Private Sub Command16_Click()
If coin < huojianjia Then
MsgBox "请先确认是否金币足够！"
Else
coin = coin - huojianjia
Option1(1).Enabled = True
End If
End Sub

Private Sub Command17_Click()
If coin < 1500 Then
MsgBox "请先确认是否金币足够！"
Else
coin = coin - 1500
Option1(2).Enabled = True
Command17.Enabled = False
End If
End Sub

Private Sub Command18_Click()
If Option2(0).Value = True Then
MsgBox "熔炼生命药水需要2A1B"
    If A > 1 And B > 0 Then
    A = A - 2: B = B - 1
    Label48.Caption = "A数量：" + Str(A) + "": Label49.Caption = "B数量：" + Str(B) + ""
    MsgBox "熔炼中……": MsgBox "熔炼成功！您获得生命药水×1": sm = sm + 1
        ronglianEXP = ronglianEXP + 30                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    Else
    MsgBox "您的熔炼材料不够！"
    End If
End If

If Option2(1).Value = True Then
MsgBox "熔炼浓厚石需要3B"
    If B > 2 Then
     B = B - 3: Label49.Caption = "B数量：" + Str(B) + ""
    MsgBox "熔炼中……": MsgBox "熔炼成功！您获得浓厚石×10": nonghou = nonghou + 10
        ronglianEXP = ronglianEXP + 20                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    Else
    MsgBox "您的熔炼材料不够！"
    End If
End If

If Option2(2).Value = True Then
MsgBox "熔炼该套餐需要4C2D"
    If C > 3 And D > 1 Then
    C = C - 4: D = D - 2
    Label50.Caption = "C数量：" + Str(C) + "": Label51.Caption = "D数量：" + Str(D) + ""
    MsgBox "熔炼中……": MsgBox "熔炼成功！您获得修炼卡×3+抽奖券×2": jiasu = jiasu + 3: choujiang = choujiang + 2
        ronglianEXP = ronglianEXP + 50                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    Else
    MsgBox "您的熔炼材料不够！"
    End If
End If

If Option2(3).Value = True Then
MsgBox "熔炼汲血刃需要14B7C5D"
    If C > 6 And B > 13 And D > 4 Then
    C = C - 7: B = B - 14: D = D - 5
    Label50.Caption = "C数量：" + Str(C) + "": Label51.Caption = "D数量：" + Str(D) + "": Label49.Caption = "B数量：" + Str(B) + ""
    MsgBox "熔炼中……": MsgBox "熔炼成功！您获得武器“汲血刃”": Option1(4).Visible = True: Option2(3).Enabled = False: Option2(3).Value = False
        ronglianEXP = ronglianEXP + 100                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    Else
    MsgBox "您的熔炼材料不够！"
    End If
End If

If Option2(4).Value = True Then
MsgBox "熔炼成就碎片需要5A8B10C3D"
    If A > 1 And B > 0 And C > 9 And D > 2 Then
    A = A - 5: B = B - 8: C = C - 10: D = D - 3
    Label48.Caption = "A数量：" + Str(A) + "": Label49.Caption = "B数量：" + Str(B) + "": Label50.Caption = "C数量：" + Str(C) + "": Label51.Caption = "D数量：" + Str(D) + ""
    MsgBox "熔炼中……": MsgBox "熔炼成功！您获得成就碎片×1": achip = achip + 1
        ronglianEXP = ronglianEXP + 1                                         '加经验
    Else
    MsgBox "您的熔炼材料不够！"
    End If
End If
End Sub

Private Sub Command19_Click()
If w(0) And wl(0) = 0 Then                                                      '金币刀区
MsgBox "提升该武器需15个浓厚石！"
    If nonghou < 15 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 1 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 15: tisheng = tisheng - 1
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！金币刀提升至等级2，获得金币+30%"
    wl(0) = 1: wa(0) = 40: Bratk = Bratk + 10
    Label57.Caption = Str(wl(0) + 1)
    Option1(0).Caption = "金币刀(获得金币+30%)"
        ronglianEXP = ronglianEXP + 60                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(0) And wl(0) = 1 Then
MsgBox "提升该武器需25个浓厚石！"
    If nonghou < 25 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 2 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 25: tisheng = tisheng - 2
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！金币刀提升至等级3，获得金币+40%"
    wl(0) = 2: wa(0) = 50: Bratk = Bratk + 10
    Label57.Caption = Str(wl(0) + 1)
    Option1(0).Caption = "金币刀(获得金币+40%)"
        ronglianEXP = ronglianEXP + 70                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(0) And wl(0) = 2 Then
MsgBox "提升该武器需40个浓厚石！"
    If nonghou < 40 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 3 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 40: tisheng = tisheng - 3
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！金币刀提升至等级4，获得金币+40%，基础报酬+10%"
    wl(0) = 3: wa(0) = 50: Bratk = Bratk + 10
    Label57.Caption = Str(wl(0) + 1)
    Option1(0).Caption = "金币刀(获得金币+40%，基础报酬+10%)"
        ronglianEXP = ronglianEXP + 80                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(0) And wl(0) = 3 And nm > 5 Then
MsgBox "提升该武器需55个浓厚石！"
    If nonghou < 55 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 4 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 55: tisheng = tisheng - 4
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！金币刀提升至等级5，获得金币+50%，基础报酬+20%"
    wl(0) = 4: wa(0) = 60: Bratk = Bratk + 10
    Label57.Caption = Str(wl(0) + 1)
    Option1(0).Caption = "金币刀(获得金币+50%，基础报酬+20%)"
        ronglianEXP = ronglianEXP + 100                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(0) And wl(0) = 4 Then MsgBox "当前武器已达最高级！"


If w(1) And wl(1) = 0 Then                                                      '火箭筒区
MsgBox "提升该武器需8个浓厚石！"
    If nonghou < 8 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 1 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 8: tisheng = tisheng - 1
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！火箭筒提升至等级2，价格变为180元！": wl(1) = 1: Label57.Caption = Str(wl(1) + 1)
    Command16.Caption = "购买火箭筒（一次性）180元": huojianjia = 180
        ronglianEXP = ronglianEXP + 60                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(1) And wl(1) = 1 Then
MsgBox "提升该武器需15个浓厚石！"
    If nonghou < 15 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 2 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 15: tisheng = tisheng - 2
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！火箭筒提升至等级3，价格变为150元！": wl(1) = 2: Label57.Caption = Str(wl(1) + 1)
     Command16.Caption = "购买火箭筒（一次性）150元": huojianjia = 150
        ronglianEXP = ronglianEXP + 70                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(1) And wl(1) = 2 Then
MsgBox "提升该武器需25个浓厚石！"
    If nonghou < 25 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 3 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 25: tisheng = tisheng - 3
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！火箭筒提升至等级4，价格变为130元！": wl(1) = 3: Label57.Caption = Str(wl(1) + 1)
     Command16.Caption = "购买火箭筒（一次性）130元": huojianjia = 130
        ronglianEXP = ronglianEXP + 80                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(1) And wl(1) = 3 And nm > 5 Then
MsgBox "提升该武器需30个浓厚石！"
    If nonghou < 30 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 4 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 30: tisheng = tisheng - 4
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！火箭筒提升至等级5，价格变为120元！攻击提升10倍！(替换武器后有效)": wl(1) = 4: Label57.Caption = Str(wl(1) + 1)
     Command16.Caption = "购买火箭筒（一次性）120元": huojianjia = 120: Option1(1).Caption = "火箭筒(10倍攻击)"
        ronglianEXP = ronglianEXP + 100                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(1) And wl(1) = 4 Then MsgBox "当前武器已达最高级！"


If w(2) And wl(2) = 0 Then                                                      '金钱大炮区
MsgBox "提升该武器需10个浓厚石！"
    If nonghou < 10 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 1 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 10: tisheng = tisheng - 1
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！金钱大炮提升至等级2，消耗金币为12%": wl(2) = 1
    Label57.Caption = Str(wl(2) + 1)
        ronglianEXP = ronglianEXP + 60                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(2) And wl(2) = 1 Then
MsgBox "提升该武器需20个浓厚石！"
    If nonghou < 20 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 2 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 20: tisheng = tisheng - 2
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！金钱大炮提升至等级3，损耗金钱为8%": wl(2) = 2
    Label57.Caption = Str(wl(2) + 1)
        ronglianEXP = ronglianEXP + 70                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(2) And wl(2) = 2 Then
MsgBox "提升该武器需35个浓厚石！"
    If nonghou < 35 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 3 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 35: tisheng = tisheng - 3
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！金钱大炮提升至等级4，损耗金钱为6%，攻击增加20%(替换武器后有效)": wl(2) = 3
    Label57.Caption = Str(wl(2) + 1)
        ronglianEXP = ronglianEXP + 80                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(2) And wl(2) = 3 And nm > 5 Then
MsgBox "提升该武器需40个浓厚石！"
    If nonghou < 40 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 4 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 40: tisheng = tisheng - 4
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！金钱大炮提升至等级5，损耗金钱为5%，攻击增加10%(替换武器后有效)": wl(2) = 4
    Label57.Caption = Str(wl(2) + 1)
        ronglianEXP = ronglianEXP + 100                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(2) And wl(2) = 4 Then MsgBox "当前武器已达最高级！"


If w(3) And wl(3) = 0 Then                                                      '钢管区
MsgBox "提升该武器需6个浓厚石！"
    If nonghou < 6 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 1 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 6: tisheng = tisheng - 1
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！钢管提升至等级2，攻击增加至75！": wl(3) = 1: Label57.Caption = Str(wl(3) + 1)
    Bratk = Bratk + 35: wa(3) = 75: Option1(3).Caption = "钢管(攻击+75)"
        ronglianEXP = ronglianEXP + 60                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(3) And wl(3) = 1 Then
MsgBox "提升该武器需15个浓厚石！"
    If nonghou < 15 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 2 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 15: tisheng = tisheng - 2
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！钢管提升至等级3，攻击增加至110！": wl(3) = 2: Label57.Caption = Str(wl(3) + 1)
    Bratk = Bratk + 35: wa(3) = 110: Option1(3).Caption = "钢管(攻击+110)"
        ronglianEXP = ronglianEXP + 70                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(3) And wl(3) = 2 Then
MsgBox "提升该武器需20个浓厚石！"
    If nonghou < 20 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 3 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 20: tisheng = tisheng - 3
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！钢管提升至等级4，攻击增加至170！": wl(3) = 3: Label57.Caption = Str(wl(3) + 1)
    Bratk = Bratk + 60: wa(3) = 170: Option1(3).Caption = "钢管(攻击+170)"
        ronglianEXP = ronglianEXP + 80                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(3) And wl(3) = 3 And nm > 5 Then
MsgBox "提升该武器需30个浓厚石！"
    If nonghou < 30 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 4 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 30: tisheng = tisheng - 4
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！钢管提升至等级5，攻击增加至240！": wl(3) = 4: Label57.Caption = Str(wl(3) + 1)
    Bratk = Bratk + 70: wa(3) = 240: Option1(3).Caption = "钢管(攻击+240)"
        ronglianEXP = ronglianEXP + 100                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(3) And wl(3) = 4 Then MsgBox "当前武器已达最高级！"


If w(4) And wl(4) = 0 Then                                                      '汲血刃区
MsgBox "提升该武器需15个浓厚石！"
    If nonghou < 15 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 1 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 15: tisheng = tisheng - 1
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！汲血刃提升至等级2，战胜后加血3%": wl(4) = 1
    Label57.Caption = Str(wl(4) + 1)
    Option1(4).Caption = "汲血刃(战胜后生命+3%)": Bratk = Bratk + 5: wa(4) = 25
        ronglianEXP = ronglianEXP + 60                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(4) And wl(4) = 1 Then
MsgBox "提升该武器需25个浓厚石！"
    If nonghou < 25 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 2 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 25: tisheng = tisheng - 2
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！汲血刃提升至等级3，战胜后加血4%": wl(4) = 2
    Label57.Caption = Str(wl(4) + 1)
    Option1(4).Caption = "汲血刃(战胜后生命+4%)": Bratk = Bratk + 10: wa(4) = 35
        ronglianEXP = ronglianEXP + 70                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(4) And wl(4) = 2 Then
MsgBox "提升该武器需40个浓厚石！"
    If nonghou < 40 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 3 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 40: tisheng = tisheng - 3
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！汲血刃提升至等级4，战胜后加血5%": wl(4) = 3
    Label57.Caption = Str(wl(4) + 1)
    Option1(4).Caption = "汲血刃(战胜后生命+5%)": Bratk = Bratk + 15: wa(4) = 50
        ronglianEXP = ronglianEXP + 80                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(4) And wl(4) = 3 And nm > 5 Then
MsgBox "提升该武器需55个浓厚石！"
    If nonghou < 55 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 4 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 55: tisheng = tisheng - 4
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！汲血刃提升至等级5，战胜后加血6%": wl(4) = 4
    Label57.Caption = Str(wl(4) + 1)
    Option1(4).Caption = "汲血刃(战胜后生命+6%)": Bratk = Bratk + 30: wa(4) = 80
        ronglianEXP = ronglianEXP + 100                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(4) And wl(4) = 4 Then MsgBox "当前武器已达最高级！"


If w(5) And wl(5) = 0 Then                                                      '复仇拳套区
MsgBox "提升该武器需10个浓厚石！"
    If nonghou < 10 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 1 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 10: tisheng = tisheng - 1
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！复仇拳套提升至等级2，攻击增加35！": wl(5) = 1: Label57.Caption = Str(wl(5) + 1)
    Bratk = Bratk + 35: wa(5) = wa(5) + 35
        ronglianEXP = ronglianEXP + 60                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(5) And wl(5) = 1 Then
MsgBox "提升该武器需15个浓厚石！"
    If nonghou < 15 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 2 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 15: tisheng = tisheng - 2
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！复仇拳套提升至等级3，每次失败后增加10金钱！": wl(5) = 2: Label57.Caption = Str(wl(5) + 1)
    Option1(5).Caption = "复仇拳套(战败后攻击+5,金钱+10)"
        ronglianEXP = ronglianEXP + 70                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(5) And wl(5) = 2 Then
MsgBox "提升该武器需20个浓厚石！"
    If nonghou < 20 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 3 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 20: tisheng = tisheng - 3
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！复仇拳套提升至等级3，每次失败后增加25金钱！": wl(5) = 3: Label57.Caption = Str(wl(5) + 1)
    Option1(5).Caption = "复仇拳套(战败后攻击+5,金钱+25)"
        ronglianEXP = ronglianEXP + 80                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(5) And wl(5) = 3 And nm > 5 Then
MsgBox "提升该武器需30个浓厚石！"
    If nonghou < 30 Then
    MsgBox "您的浓厚石数量不够！"
    ElseIf tisheng < 4 Then
    MsgBox "您的武器提升卷数量不够！(武器当前几级，提升卷就要几张)"
    Else
    nonghou = nonghou - 30: tisheng = tisheng - 4
    Label56.Caption = "浓厚石数量：" + Str(nonghou) + "": Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
    MsgBox "正在熔炼……"
    MsgBox "熔炼成功！复仇拳套提升至等级3，每次失败后增加15生命！": wl(5) = 4: Label57.Caption = Str(wl(5) + 1)
    Option1(5).Caption = "复仇拳套(战败后攻击+5,金钱+10,生命+15)"
        ronglianEXP = ronglianEXP + 100                                         '加经验
        If ronglianEXP > 0 And ronglianEXP <= ronglianLV * 300 Then Label46.Visible = True: Label46.Width = ronglianEXP / (ronglianLV * 300) * 7000
        If ronglianEXP >= ronglianLV * 300 Then MsgBox "您的熔炼等级提升!!": ronglianEXP = 0: ronglianLV = ronglianLV + 1: Label46.Visible = False             '升级效果
    End If
    Exit Sub
End If

If w(5) And wl(5) = 4 Then MsgBox "当前武器已达最高级！"

If w(7) Then MsgBox "该武器过于变态，无法升级！"
End Sub

Private Sub Command2_Click(Index As Integer)
xiulian(0) = False: xiulian(1) = False: xiulian(2) = False
Select Case Index
    Case 0
    leftime = 20
    Case 1
    leftime = 60
    Case 2
    leftime = 120
End Select
Timer2.Enabled = True
xiulian(Index) = True
End Sub

Private Sub Command20_Click()
Text3.SetFocus
If coin < 25 Then
MsgBox "请先确认是否金币足够！"
ElseIf eneHP < teHP Then
MsgBox "请不要在战斗时逛商店……"
SSTab1.Tab = 0
Else
coin = coin - 25
tisheng = tisheng + 1
End If
End Sub

Private Sub Command21_Click()
Text3.SetFocus
If coin < 5 Then
MsgBox "你怎么穷到连5块钱的东西都买不起了……"
ElseIf eneHP < teHP Then
MsgBox "请不要在战斗时逛商店……"
SSTab1.Tab = 0
Else
coin = coin - 5
nonghou = nonghou + 1
End If
End Sub

Private Sub Command22_Click()
Text3.SetFocus
If teHP <> 0 Then
If w(7) Then nowHP = nowHP + 2 * enatk(NowEnemy + 1)
    If nowHP - enatk(NowEnemy + 1) > 0 And eneHP - Bratk > 0 Then
    
    If w(2) Then coin = Int(coin * 0.98)                                       '金钱大炮负效果
    nowHP = nowHP - enatk(NowEnemy + 1)
    Label2.Width = nowHP / totHP * 7000
    eneHP = eneHP - Bratk
    Label4.Width = eneHP / teHP * 7000
    ElseIf nowHP - enatk(NowEnemy + 1) <= 0 Then
    Label8.Caption = "0"
          If eneHP - Bratk > 0 Then
          Label11.Caption = Str(eneHP - Bratk)
          Label4.Width = (eneHP - Bratk) / teHP * 7000
          ElseIf eneHP - Bratk <= 0 Then Label11.Caption = "1"
          Label4.Width = 1 / teHP * 7000
          End If
    Label2.Visible = False
    MsgBox "你失败了！"
      If w(5) Then Bratk = Bratk + 5: wa(5) = wa(5) + 5
      If w(5) And wl(5) = 2 Then coin = coin + 10
      If w(5) And wl(5) > 2 Then coin = coin + 25
      If w(5) And wl(5) = 4 Then nowHP = nowHP + 15: totHP = totHP + 15
      
      If w(1) Then Bratk = Bratk - wa(1): Option1(1).Enabled = False: w(1) = False
        
    nowHP = totHP
    eneHP = teHP
    Label2.Visible = True
    Label2.Width = 7000
    Label4.Width = 7000
    ElseIf eneHP - Bratk <= 0 Then
    Label8.Caption = Str(nowHP - enatk(NowEnemy + 1))
    Label2.Width = (nowHP - enatk(NowEnemy + 1)) / totHP * 7000
    Label11.Caption = "0"
    Label4.Visible = False
    MsgBox "你胜利了！"
    If NowEnemy < 11 Then Label1(NowEnemy + 1).Visible = True
    If NowEnemy = 11 Then enatk(12) = enatk(12) + 0.5 * Bratk: wangweiKILL = wangweiKILL + 1
    Randomize                                                                                   '刷装备：钢管
    If NowEnemy = 5 And nm > 2 And Rnd < 0.1 And Option1(3).Visible = False Then Option1(3).Visible = True: MsgBox "恭喜您刷出了武器“钢管”！"
    Randomize                                                                                   '刷石头
    If Rnd < 0.5 Then nonghou = nonghou + 1
    
    If w(1) Then Bratk = Bratk - wa(1): Option1(1).Enabled = False: w(1) = False: Label40.Caption = "": Label52.Caption = "": Label57.Caption = "": Option1(1).Value = False    '火箭筒炸掉
    
    If w(0) = False And wl(0) = 3 Then MsgBox "" + Str(Int(Greward(NowEnemy) * 1.1)) + "G，" + Str(EXPreward(NowEnemy)) + "EXP"              '金币刀判定
    If w(0) = False And wl(0) = 4 Then MsgBox "" + Str(Int(Greward(NowEnemy) * 1.2)) + "G，" + Str(EXPreward(NowEnemy)) + "EXP"
    If w(0) = False And wl(0) < 3 Then MsgBox "" + Str(Greward(NowEnemy)) + "G，" + Str(EXPreward(NowEnemy)) + "EXP"
    If w(0) And wl(0) = 0 Then MsgBox "" + Str(Int(Greward(NowEnemy) * 1.2)) + "G，" + Str(EXPreward(NowEnemy)) + "EXP"
    If w(0) And wl(0) = 1 Then MsgBox "" + Str(Int(Greward(NowEnemy) * 1.3)) + "G，" + Str(EXPreward(NowEnemy)) + "EXP"
    If w(0) And wl(0) = 2 Then MsgBox "" + Str(Int(Greward(NowEnemy) * 1.4)) + "G，" + Str(EXPreward(NowEnemy)) + "EXP"
    If w(0) And wl(0) = 3 Then MsgBox "" + Str(Int(Greward(NowEnemy) * 1.4)) + "G，" + Str(EXPreward(NowEnemy)) + "EXP"
    If w(0) And wl(0) = 4 Then MsgBox "" + Str(Int(Greward(NowEnemy) * 1.5)) + "G，" + Str(EXPreward(NowEnemy)) + "EXP"
    
    Randomize
    If NowEnemy = 4 And Rnd < 0.6 And nm > 2 Then MsgBox "你刷出一个A": A = A + 1
    If NowEnemy = 5 And Rnd < 0.5 And nm > 2 Then MsgBox "你刷出一个B": B = B + 1
    If NowEnemy = 6 And Rnd < 0.4 And nm > 2 Then MsgBox "你刷出一个C": C = C + 1
    If NowEnemy = 7 And Rnd < 0.3 And nm > 2 Then MsgBox "你刷出一个D": D = D + 1
    
        If Label23.Caption = "状态：未完成" And nm = 0 And NowEnemy = 0 Then        '任务1
        MsgBox "您的任务完成！！"
        Label23.Caption = "状态：已完成"
        End If
        
        If NowEnemy = 1 Then MouseKill = MouseKill + 1
        If Label23.Caption = "状态：未完成" And nm = 2 And MouseKill = 5 Then        '任务3
        MsgBox "您的任务完成！！"
        Label23.Caption = "状态：已完成"
        End If
        
        If nm = 3 And NowEnemy = 0 And cont = 0 Then cont = cont + 1   '任务4
        If nm = 3 And NowEnemy = 1 And cont = 1 Then cont = cont + 1
        If nm = 3 And NowEnemy = 3 And cont = 2 Then cont = cont + 1: MsgBox "您的任务完成！！": Label23.Caption = "状态：已完成"

        If Label23.Caption = "状态：未完成" And nm = 5 And NowEnemy = 8 Then MsgBox "您的任务完成！！": Label23.Caption = "状态：已完成"    '任务6
    
        If Label23.Caption = "状态：未完成" And nm = 7 And NowEnemy = 11 Then MsgBox "您的任务完成！！": Label23.Caption = "状态：已完成"    '任务8
        
        If Label23.Caption = "状态：未完成" And nm = 8 And wangweiKILL = 6 Then MsgBox "您的任务完成！！": Label23.Caption = "状态：已完成"  '任务8

    coin = coin + Greward(NowEnemy)                                                     '金币刀判定
    If w(0) Then coin = coin + Int(Greward(NowEnemy) * 0.2)
    If w(0) And wl(0) > 0 Then coin = coin + Int(Greward(NowEnemy) * 0.1)
    If w(0) And wl(0) > 1 Then coin = coin + Int(Greward(NowEnemy) * 0.1)
    If w(0) = False And wl(0) = 3 Then coin = coin + Int(Greward(NowEnemy) * 0.1)
    If w(0) = False And wl(0) = 4 Then coin = coin + Int(Greward(NowEnemy) * 0.1)
    
    nowEXP = nowEXP + EXPreward(NowEnemy)                                           '加经验
    If nowEXP > 0 Then Label34.Visible = True: Label34.Width = nowEXP / (50 * love) * 7000
    If nowEXP >= love * 50 Then MsgBox "升级，获得金币!!": coin = Int(coin * 1.25): nowEXP = 0: love = love + 1: Label34.Visible = False             '升级效果
    
    If w(4) And wl(4) = 0 Then totHP = totHP + Int(totHP * 0.02)                                     '汲血刃判定
    If w(4) And wl(4) = 1 Then totHP = totHP + Int(totHP * 0.03)
    If w(4) And wl(4) = 2 Then totHP = totHP + Int(totHP * 0.04)
    If w(4) And wl(4) = 3 Then totHP = totHP + Int(totHP * 0.05)
    If w(4) And wl(4) = 4 Then totHP = totHP + Int(totHP * 0.06)
    
    nowHP = totHP
    eneHP = teHP
    Label4.Visible = True
    Label2.Width = 7000
    Label4.Width = 7000
    End If
Else
MsgBox "请先选择【副本】！"
SSTab1.Tab = 3
End If
If 11 < NowEnemy And NowEnemy < 16 Then NowEnemy = 0: Timer3.Enabled = True
End Sub

Private Sub Command23_Click()
Text3.SetFocus
If sm = 0 Then
MsgBox "你没有生命药水！"
ElseIf eneHP < teHP Then
MsgBox "生命药水请在战斗之前喝！"
Else
nowHP = nowHP + 40
totHP = totHP + 40
sm = sm - 1
End If
End Sub

Private Sub Command5_Click()
If jiasu = 0 Then
MsgBox "你没有修炼加速卡！"
Else
leftime = leftime - 50
jiasu = jiasu - 1
End If
End Sub

Private Sub Command6_Click()
MsgBox "提示：这是本游戏中最最最最最最最强大的敌人，拥有无比的神力与超越天理的战斗值，你真的要挑战他吗？"
MsgBox "真的，你真的要挑战游戏中最最最最最最最强大的敌人吗？"
MsgBox "真的，你最好还是放弃算了，你不可能打败他的，不可能，绝不可能。"
MsgBox "什么？你还不放弃吗？"
MsgBox "祝你好运，来世你还是一个好人。"
MsgBox "."
MsgBox ".."
MsgBox "..."
MsgBox "Cation!"
MsgBox "Cation!!!"
MsgBox "Cation!!!!!"
MsgBox "（神的声音）去死吧，人类"
NowEnemy = 15
SSTab1.Tab = 0
Timer3.Enabled = True
End Sub

Private Sub Command7_Click()
If choujiang = 0 Then
    MsgBox "抽奖次数不够！"
    Else
    Randomize
    lucky = Int(Rnd() * 101)
    If lucky < 31 Then
        MsgBox "320HP!"
        nowHP = nowHP + 320
        totHP = totHP + 320
    ElseIf lucky < 61 Then
        MsgBox "300EXP"
        nowEXP = nowEXP + 300                                          '加经验
        If nowEXP > 0 Then Label34.Visible = True: Label34.Width = nowEXP / (50 * love) * 7000
        If nowEXP >= love * 50 Then MsgBox "升级，获得金币!!": coin = Int(coin * 1.25): nowEXP = 0: love = love + 1: Label34.Visible = False             '升级效果
    ElseIf lucky < 91 Then
        MsgBox "8攻击"
        Bratk = Bratk + 8
    ElseIf lucky < 94 Then
        MsgBox "8攻击"
        Bratk = Bratk + 8
    ElseIf lucky < 97 Then
        MsgBox "8攻击"
        Bratk = Bratk + 8
    ElseIf lucky < 100 Then
        Bratk = Bratk + 8
    Else
    MsgBox "空的！"
    End If
    choujiang = choujiang - 1
End If
End Sub

Private Sub Command8_Click()
Text3.SetFocus
If coin < 70 Then
MsgBox "请先确认是否金币足够！"
ElseIf eneHP < teHP Then
MsgBox "请不要在战斗时逛商店……"
SSTab1.Tab = 0
Else
coin = coin - 70
Bratk = Bratk + 3
End If
End Sub

Private Sub Command9_Click()
Text3.SetFocus
If coin < 70 Then
MsgBox "请先确认是否金币足够！"
ElseIf eneHP < teHP Then
MsgBox "请不要在战斗时逛商店……"
SSTab1.Tab = 0
Else
coin = coin - 70
nowHP = nowHP + 30
totHP = totHP + 30
End If
End Sub

Private Sub Form_Load()
totHP = 20: nowHP = 20: Bratk = 5: coin = 30: love = 1: ronglianLV = 1: huojianjia = 200

enemy(0) = "水果刀": enemy(1) = "小潘": enemy(2) = "王de海": enemy(3) = "小鸡鸡"
enemy(4) = "张鸡": enemy(5) = "狂热狂热的港独分子": enemy(6) = "打菜老太": enemy(7) = "张鸡"
enemy(8) = "曹管": enemy(9) = "叶祥彪": enemy(10) = "卓海燕": enemy(11) = "王威"
enemy(12) = "鲤鱼王": enemy(13) = "若鹭姬": enemy(14) = "近海之主": enemy(15) = "蒋哥"

enatk(1) = 2: enatk(2) = 7: enatk(3) = 11: enatk(4) = 11: enatk(5) = 18: enatk(6) = 30: enatk(7) = 60: enatk(8) = 90: enatk(9) = 150: enatk(10) = 150: enatk(11) = 300: enatk(12) = 50: enatk(13) = 120: enatk(14) = 1400: enatk(15) = 12
teneHP(1) = 30: teneHP(2) = 60: teneHP(3) = 120: teneHP(4) = 200: teneHP(5) = 230: teneHP(6) = 400: teneHP(7) = 800: teneHP(8) = 1200: teneHP(9) = 1800: teneHP(10) = 2000: teneHP(11) = 2500: teneHP(12) = 3000: teneHP(13) = 20000: teneHP(14) = 4000: teneHP(15) = 200: teneHP(16) = 5

Greward(0) = 4: Greward(1) = 7: Greward(2) = 10: Greward(3) = 14: Greward(4) = 20: Greward(5) = 30: Greward(6) = 35: Greward(7) = 40: Greward(8) = 50: Greward(9) = 60: Greward(10) = 80: Greward(11) = 75: Greward(12) = 20: Greward(13) = 1500: Greward(14) = 300: Greward(15) = 9999999
EXPreward(0) = 2: EXPreward(1) = 5: EXPreward(2) = 10: EXPreward(3) = 20: EXPreward(4) = 40: EXPreward(5) = 50: EXPreward(6) = 70: EXPreward(7) = 80: EXPreward(8) = 100: EXPreward(9) = 100: EXPreward(10) = 150: EXPreward(11) = 200
mission(0) = "请清理水果村一次": mission(1) = "在商店购买一次1攻击": mission(2) = "请杀死五只鲤鱼": mission(3) = "依次干掉一只果冻、一只鲤鱼、一只变异鸡": mission(4) = "获得总计20个熔炼材料": mission(5) = "打败城管": mission(6) = "强化任意一件武器至LV5": mission(7) = "打败王威一次": mission(8) = "打败王威六次": mission(9) = "当前版本无任务"
wa(0) = 30: wa(3) = 40: wa(4) = 20: wa(6) = 170: wa(7) = 350
End Sub

Private Sub Label1_Click(Index As Integer)
MsgBox "提示：敌人生命" + Str(teneHP(Index + 1)) + "，" + "攻击" + Str(enatk(Index + 1))
MsgBox "请开始进行战斗"
NowEnemy = Index
SSTab1.Tab = 0
Timer3.Enabled = True
End Sub

Private Sub Label12_Click()
If Label12.Caption = "关于本战斗游戏……" Then
Label19.Visible = True
Label58.Visible = True
Label21.Visible = True
Label12.Caption = "隐藏"
Else
Label19.Visible = False
Label58.Visible = False
Label21.Visible = False
Label12.Caption = "关于本战斗游戏……"
End If
End Sub

Private Sub Label19_Click()
If coin < 70 Then
MsgBox "请先确认是否金币足够！"
ElseIf eneHP < teHP Then
MsgBox "请不要在战斗时逛商店……"
SSTab1.Tab = 0
Else
coin = coin - 70
nowHP = nowHP + 30
totHP = totHP + 30
End If
End Sub

Private Sub Label21_Click()
ShellExecute Me.hwnd, "open", "http://pan.baidu.com/share/home?uk=4163755544&view=share#category/type=0", "", "", 1
End Sub

Private Sub Label26_Click()
If Label23.Caption = "状态：已完成" Then
    If nm = 0 Then
    MsgBox "获得20G"
    coin = coin + 20
    End If
    
    If nm = 1 Then
    MsgBox "获得一个修炼加速卡"
    jiasu = jiasu + 1
    End If
    
    If nm = 2 Then
    MsgBox "来自鼠大将的可靠讯息：4~6副本中可能存在装备"
    MsgBox "熔炼功能已开启": Frame6.Visible = True: Frame5.Visible = True: Label47.Visible = True
    Label22(3).Visible = True: Label41.Visible = False: Label42.Visible = True: Command20.Visible = True: Command21.Visible = True
    End If
    
    If nm = 3 Then
    MsgBox "获得两瓶生命药水"
    sm = sm + 2
    End If
    
    If nm = 4 Then MsgBox "获得20个浓厚石、5张武器升级券、4*ABCD": nonghou = nonghou + 20: tisheng = tisheng + 5: A = A + 4: B = B + 4: C = C + 4: D = D + 4

    If nm = 5 Then MsgBox "获得城管没收的1000元钱": coin = coin + 1000
    
    If nm = 6 Then MsgBox "获得成就：我不知道"
    
    If nm = 7 Then MsgBox "生命*1.2、攻击*1.1！": nowHP = Int(nowHP * 1.2): totHP = Int(totHP * 1.2): Bratk = Int(Bratk * 1.1)

    If nm = 8 Then MsgBox "这个任务并没有什么奖励": Label14.Caption = "不要再往下看了，真没有": Label23.Caption = "状态：你就这么不相信我？"

nm = nm + 1
If nm = 2 Then MouseKill = 0
If nm = 7 Then wangweiKILL = 0
If nm < 9 Then Label23.Caption = "状态：未完成"
End If
End Sub

Private Sub Label36_Click(Index As Integer)
MsgBox "染料已使用"
Label13.ForeColor = Label36(Index).ForeColor
Label6.ForeColor = Label36(Index).ForeColor
Label7.ForeColor = Label36(Index).ForeColor
Label15.ForeColor = Label36(Index).ForeColor
Label16.ForeColor = Label36(Index).ForeColor
Label8.ForeColor = Label36(Index).ForeColor
End Sub

Private Sub Label58_Click()
ShellExecute Me.hwnd, "open", "http://tieba.baidu.com/p/3006282155", "", "", 1
End Sub

Private Sub Option1_Click(Index As Integer)
    For i = 0 To 7
    If w(i) = True Then w(i) = False: Bratk = Bratk - wa(i)
    Next i
If Index = 0 Then Label40.Caption = "金钱刀": Label52.Caption = "金钱刀": Label57.Caption = Str(wl(0) + 1)

If Index = 1 And wl(1) < 4 Then wa(1) = 7 * Bratk: Label40.Caption = "火箭筒": Label52.Caption = "火箭筒": Label57.Caption = Str(wl(1) + 1)
If Index = 1 And wl(1) = 4 Then wa(1) = 9 * Bratk: Label40.Caption = "火箭筒": Label52.Caption = "火箭筒": Label57.Caption = Str(wl(1) + 1)

If Index = 2 And wl(2) < 3 Then wa(2) = Bratk: Label40.Caption = "金钱大炮": Label52.Caption = "金钱大炮": Label57.Caption = Str(wl(2) + 1)
If Index = 2 And wl(2) = 3 Then wa(2) = 1.2 * Bratk: Label40.Caption = "金钱大炮": Label52.Caption = "金钱大炮": Label57.Caption = Str(wl(2) + 1)
If Index = 2 And wl(2) = 4 Then wa(2) = 1.3 * Bratk: Label40.Caption = "金钱大炮": Label52.Caption = "金钱大炮": Label57.Caption = Str(wl(2) + 1)

If Index = 3 Then Label40.Caption = "钢管": Label52.Caption = "钢管": Label57.Caption = Str(wl(3) + 1)
If Index = 4 Then Label40.Caption = "汲血刃": Label52.Caption = "汲血刃": Label57.Caption = Str(wl(4) + 1)
If Index = 5 Then Label40.Caption = "复仇拳套": Label52.Caption = "复仇拳套": Label57.Caption = Str(wl(5) + 1)
If Index = 6 Then Label40.Caption = "枪魚枪": Label52.Caption = "枪魚枪": Label57.Caption = Str(wl(6) + 1)
If Index = 7 Then Label40.Caption = "VB6": Label52.Caption = "VB6": Label57.Caption = "Max"
w(Index) = True
Bratk = Bratk + wa(Index)
End Sub

Private Sub Text1_Change()
Label13.Caption = Text1.Text
End Sub

Private Sub Timer1_Timer()
If ronglianLV = 2 And Option1(5).Visible = False Then MsgBox "熔炼等级提升，获得武器“复仇拳套”": MsgBox "新的熔炼物品解锁": Option1(5).Visible = True: Option2(2).Visible = True: Option2(3).Visible = True
If ronglianLV = 3 And Option2(4).Visible = False Then MsgBox "新的熔炼物品解锁": Option2(4).Visible = True
If achip = 3 And Option2(4).Enabled = True Then Option2(4).Enabled = False: MsgBox "解锁成就“买来的”"
Label8.Caption = Str(nowHP)
Label7.Caption = "/" + Str(totHP)
Label11.Caption = Str(eneHP)
If NowEnemy = 15 Then Label11.Caption = "????"
Label10.Caption = "/" + Str(teHP)
If NowEnemy = 15 Then Label10.Caption = "/????"
Label38.Caption = "HP：" + Str(totHP)
Label15.Caption = "攻击：" + Str(Bratk)
Label20.Caption = "金币：" + Str(coin)
Label37.Caption = Str(nowEXP) + "/" + Str(50 * love)
Label25.Caption = mission(nm)
Label27.Caption = "生命药水数量：" + Str(sm)
Label32.Caption = "您的加速卡数量 " + Str(jiasu) + ""
Label35.Caption = "剩余抽奖次数 " + Str(choujiang) + ""
Label16.Caption = "LV" + Str(love)
Label42.Caption = Str(ronglianEXP) + "/" + Str(300 * ronglianLV)
Label47.Caption = "您的熔炼等级：" + Str(ronglianLV) + ""
Label48.Caption = "A数量：" + Str(A) + ""
Label49.Caption = "B数量：" + Str(B) + ""
Label50.Caption = "C数量：" + Str(C) + ""
Label51.Caption = "D数量：" + Str(D) + ""
Label55.Caption = "武器提升卷数量：" + Str(tisheng) + ""
Label56.Caption = "浓厚石数量：" + Str(nonghou) + ""
If Timer2.Enabled = True And leftime >= 0 Then Label31.Caption = "剩余时间: " + Str(leftime) + ""
If Timer2.Enabled = True And leftime < 0 Then Label31.Caption = "剩余时间: 0"
If Timer2.Enabled = False Then Label31.Caption = "剩余时间: 请先操作"
If nm = 2 And MouseKill < 11 Then Label14.Caption = "目前进度：" + Str(MouseKill) + "/5"
If nm = 3 Then Label14.Caption = "目前进度：" + Str(cont) + "/3"

If nm = 4 And A + B + C + D <= 20 Then Label14.Caption = "目前进度：" + Str(A + B + C + D) + "/20"  '任务5
If nm = 4 And A + B + C + D >= 20 And Label23.Caption = "状态：未完成" Then MsgBox "您的任务完成！！": Label23.Caption = "状态：已完成"

If nm = 7 Or nm = 5 Or nm = 6 Then Label14.Caption = "一次"

If nm = 8 And wangweiKILL <= 6 Then Label14.Caption = "目前进度：" + Str(wangweiKILL) + "/6"

If Option1(1).Enabled Then Command16.Enabled = False
If Option1(1).Enabled = False Then Command16.Enabled = True

For i = 0 To 6
If wl(i) = 4 And nm = 6 And Label23.Caption = "状态：未完成" Then MsgBox "您的任务完成！！": Label23.Caption = "状态：已完成"
Next i
End Sub

Private Sub Timer2_Timer()
If leftime > 0 Then leftime = leftime - 1
If leftime < 0 Then leftime = 0
If leftime = 0 Then
Label31.Caption = "剩余时间: 0"
If xiulian(0) Then MsgBox "修炼完毕，增加4攻击"
If xiulian(1) Then MsgBox "修炼完毕，增加6攻击30生命"
If xiulian(2) Then MsgBox "修炼完毕，增加15攻击80生命"
If xiulian(0) Then Bratk = Bratk + 4
If xiulian(1) Then Bratk = Bratk + 6
If xiulian(1) Then nowHP = nowHP + 30
If xiulian(1) Then totHP = totHP + 30
If xiulian(2) Then Bratk = Bratk + 15
If xiulian(2) Then nowHP = nowHP + 80
If xiulian(2) Then totHP = totHP + 80
xiulian(0) = False: xiulian(1) = False: xiulian(2) = False
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
nowHP = totHP
Label2.Width = 7000
Label4.Width = 7000
Label3.Caption = enemy(NowEnemy)
teHP = teneHP(NowEnemy + 1)
eneHP = teHP
Timer3.Enabled = False
End Sub
