VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   Caption         =   "桔梗射擊遊戲 v 7.1"
   ClientHeight    =   9675
   ClientLeft      =   4380
   ClientTop       =   4800
   ClientWidth     =   13665
   Icon            =   "桔梗2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "桔梗2.frx":0CCA
   ScaleHeight     =   645
   ScaleMode       =   3  '像素
   ScaleWidth      =   911
   Begin VB.PictureBox Picture2 
      Height          =   375
      Left            =   13920
      ScaleHeight     =   315
      ScaleWidth      =   195
      TabIndex        =   31
      Top             =   12120
      Width           =   255
      Begin VB.Image Image34 
         Height          =   5055
         Left            =   360
         Picture         =   "桔梗2.frx":1847D
         Top             =   240
         Width           =   2850
      End
      Begin VB.Image Image35 
         Height          =   5055
         Left            =   4200
         Picture         =   "桔梗2.frx":1CE9B
         Top             =   240
         Width           =   2850
      End
   End
   Begin VB.Timer Timer31 
      Enabled         =   0   'False
      Interval        =   70
      Left            =   7560
      Top             =   120
   End
   Begin VB.Timer Timer30 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   7080
      Top             =   8760
   End
   Begin VB.Timer Timer29 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6360
      Top             =   8760
   End
   Begin VB.Timer Timer28 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   7080
      Top             =   8160
   End
   Begin VB.Timer Timer27 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6360
      Top             =   8160
   End
   Begin VB.Timer Timer26 
      Left            =   10320
      Top             =   7920
   End
   Begin VB.Timer Timer25 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   10320
      Top             =   7440
   End
   Begin VB.PictureBox Picture6 
      Height          =   495
      Left            =   14400
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   29
      Top             =   12120
      Visible         =   0   'False
      Width           =   495
      Begin VB.Image Image21 
         Height          =   795
         Index           =   2
         Left            =   3840
         Picture         =   "桔梗2.frx":2264A
         Top             =   4080
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.Image Image21 
         Height          =   780
         Index           =   1
         Left            =   4080
         Picture         =   "桔梗2.frx":233B7
         Top             =   3840
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.Image Image21 
         Height          =   840
         Index           =   0
         Left            =   4320
         Picture         =   "桔梗2.frx":23E55
         Top             =   3600
         Visible         =   0   'False
         Width           =   2700
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   15000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   28
      Top             =   12360
      Width           =   255
      Begin VB.Timer Timer12 
         Enabled         =   0   'False
         Interval        =   40
         Left            =   3840
         Top             =   2040
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   80
         Left            =   6840
         Top             =   3960
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   10
         Left            =   4320
         Picture         =   "桔梗2.frx":2516F
         Top             =   3600
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   9
         Left            =   4440
         Picture         =   "桔梗2.frx":2A071
         Top             =   3480
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   8
         Left            =   4560
         Picture         =   "桔梗2.frx":2AE5E
         Top             =   3360
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   855
         Index           =   7
         Left            =   4680
         Picture         =   "桔梗2.frx":2BD4F
         Top             =   3240
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   6
         Left            =   4800
         Picture         =   "桔梗2.frx":2CDBE
         Top             =   3120
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   5
         Left            =   4920
         Picture         =   "桔梗2.frx":2DEBE
         Top             =   3000
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   4
         Left            =   5040
         Picture         =   "桔梗2.frx":2F0C9
         Top             =   2880
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   3
         Left            =   5160
         Picture         =   "桔梗2.frx":302C9
         Top             =   2760
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   2
         Left            =   5280
         Picture         =   "桔梗2.frx":3146E
         Top             =   2640
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   1
         Left            =   5400
         Picture         =   "桔梗2.frx":325ED
         Top             =   2520
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image11 
         Height          =   840
         Index           =   0
         Left            =   5520
         Picture         =   "桔梗2.frx":335E6
         Top             =   2400
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   10
         Left            =   240
         Picture         =   "桔梗2.frx":342CB
         Top             =   2640
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   9
         Left            =   480
         Picture         =   "桔梗2.frx":35295
         Top             =   2400
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   8
         Left            =   720
         Picture         =   "桔梗2.frx":3629B
         Top             =   2160
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   7
         Left            =   960
         Picture         =   "桔梗2.frx":37228
         Top             =   1920
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   6
         Left            =   1200
         Picture         =   "桔梗2.frx":381DD
         Top             =   1680
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   5
         Left            =   1440
         Picture         =   "桔梗2.frx":3927F
         Top             =   1440
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   4
         Left            =   1680
         Picture         =   "桔梗2.frx":3A51E
         Top             =   1200
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   3
         Left            =   1920
         Picture         =   "桔梗2.frx":3B5C1
         Top             =   960
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   2
         Left            =   2160
         Picture         =   "桔梗2.frx":3C59A
         Top             =   720
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Image Image12 
         Height          =   600
         Index           =   1
         Left            =   2520
         Picture         =   "桔梗2.frx":3D581
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Image Image12 
         Height          =   675
         Index           =   0
         Left            =   2640
         Picture         =   "桔梗2.frx":3E538
         Top             =   240
         Visible         =   0   'False
         Width           =   2430
      End
   End
   Begin VB.Timer Timer24 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   1320
   End
   Begin VB.Timer Timer23 
      Enabled         =   0   'False
      Interval        =   650
      Left            =   5280
      Top             =   120
   End
   Begin VB.Timer Timer21 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11040
      Top             =   6480
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Interval        =   850
      Left            =   12240
      Top             =   6480
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   960
      Top             =   5520
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   180
      Left            =   7680
      Top             =   6240
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   180
      Left            =   7200
      Top             =   6240
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   180
      Left            =   6720
      Top             =   6240
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   180
      Left            =   6240
      Top             =   6240
   End
   Begin VB.Timer Timer17 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7680
      Top             =   6840
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7200
      Top             =   6840
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6720
      Top             =   6840
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6240
      Top             =   6840
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   650
      Left            =   5280
      Top             =   600
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   7800
      Top             =   4680
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   9120
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   9120
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   9120
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   120
      Top             =   9120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   960
      Top             =   7200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   960
      Top             =   6240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   180
      Left            =   1680
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   8160
   End
   Begin VB.PictureBox Picture1 
      Height          =   7215
      Left            =   14040
      Picture         =   "桔梗2.frx":3F47B
      ScaleHeight     =   7155
      ScaleWidth      =   10725
      TabIndex        =   0
      Top             =   120
      Width           =   10785
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   6360
         Width           =   4935
         Begin VB.OptionButton Option1 
            Caption         =   "輕  鬆"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   39
            Top             =   120
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "簡  單  "
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1560
            Style           =   1  '圖片外觀
            TabIndex        =   38
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "普  通"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2640
            Style           =   1  '圖片外觀
            TabIndex        =   37
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "困  難"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   3720
            Style           =   1  '圖片外觀
            TabIndex        =   36
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   240
         TabIndex        =   32
         Top             =   5520
         Width           =   4935
         Begin VB.OptionButton Option2 
            Caption         =   "桔  梗"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   34
            Top             =   120
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.OptionButton Option2 
            Caption         =   "亞爾麗塔"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   3360
            Style           =   1  '圖片外觀
            TabIndex        =   33
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080FFFF&
         Caption         =   "排行榜"
         Height          =   615
         Left            =   9600
         Style           =   1  '圖片外觀
         TabIndex        =   25
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "改版記錄"
         Height          =   615
         Left            =   8280
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "鍵盤按鍵"
         Height          =   615
         Left            =   5400
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   6360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FFFF&
         Caption         =   "儲存進度"
         Height          =   615
         Left            =   6840
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FFFF&
         Caption         =   "讀取進度"
         Height          =   615
         Left            =   5400
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "開始"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   20.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6840
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   6360
         Width           =   3735
      End
      Begin VB.Image Image33 
         Height          =   5055
         Left            =   360
         Picture         =   "桔梗2.frx":523C2
         Top             =   360
         Width           =   2850
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   21.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5355
         Left            =   5280
         TabIndex        =   2
         Top             =   120
         Width           =   5385
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Left            =   14400
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   26
      Top             =   11640
      Visible         =   0   'False
      Width           =   375
      Begin VB.Image Image5 
         Height          =   765
         Index           =   0
         Left            =   4680
         Picture         =   "桔梗2.frx":56DE0
         Top             =   0
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Image Image5 
         Height          =   840
         Index           =   1
         Left            =   4680
         Picture         =   "桔梗2.frx":580F5
         Top             =   720
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Image Image5 
         Height          =   780
         Index           =   2
         Left            =   4680
         Picture         =   "桔梗2.frx":593EE
         Top             =   1440
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.Image Image5 
         Height          =   795
         Index           =   3
         Left            =   4680
         Picture         =   "桔梗2.frx":59E8F
         Top             =   2160
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.Image Image7 
         Height          =   2775
         Index           =   6
         Left            =   0
         Picture         =   "桔梗2.frx":5ABF9
         Top             =   1560
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Image Image7 
         Height          =   2775
         Index           =   5
         Left            =   480
         Picture         =   "桔梗2.frx":5DE26
         Top             =   1200
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Image Image7 
         Height          =   2775
         Index           =   4
         Left            =   960
         Picture         =   "桔梗2.frx":6112C
         Top             =   840
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Image Image7 
         Height          =   2775
         Index           =   3
         Left            =   1440
         Picture         =   "桔梗2.frx":6430E
         Top             =   600
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Image Image7 
         Height          =   2775
         Index           =   2
         Left            =   1800
         Picture         =   "桔梗2.frx":67073
         Top             =   480
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Image Image7 
         Height          =   2775
         Index           =   1
         Left            =   2160
         Picture         =   "桔梗2.frx":69C83
         Top             =   240
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Image Image7 
         Height          =   2775
         Index           =   0
         Left            =   2520
         Picture         =   "桔梗2.frx":6C69E
         Top             =   0
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Image Image10 
         Height          =   840
         Index           =   2
         Left            =   3840
         Picture         =   "桔梗2.frx":6EB3B
         Top             =   3480
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image10 
         Height          =   840
         Index           =   1
         Left            =   4080
         Picture         =   "桔梗2.frx":73A3D
         Top             =   3240
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Image Image10 
         Height          =   840
         Index           =   0
         Left            =   4320
         Picture         =   "桔梗2.frx":7893F
         Top             =   3000
         Visible         =   0   'False
         Width           =   1785
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   255
      Left            =   15120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   11760
      Visible         =   0   'False
      Width           =   255
      Begin VB.Timer Timer22 
         Enabled         =   0   'False
         Interval        =   75
         Left            =   840
         Top             =   600
      End
      Begin VB.Timer Timer20 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3720
         Top             =   2280
      End
      Begin VB.Image Image15 
         Height          =   960
         Index           =   7
         Left            =   2160
         Picture         =   "桔梗2.frx":7D841
         Top             =   2040
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image Image15 
         Height          =   960
         Index           =   6
         Left            =   2400
         Picture         =   "桔梗2.frx":7DB5E
         Top             =   1800
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image Image15 
         Height          =   960
         Index           =   5
         Left            =   2640
         Picture         =   "桔梗2.frx":7DE3A
         Top             =   1680
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image Image15 
         Height          =   960
         Index           =   4
         Left            =   2760
         Picture         =   "桔梗2.frx":7E10C
         Top             =   1560
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image Image15 
         Height          =   960
         Index           =   3
         Left            =   2880
         Picture         =   "桔梗2.frx":7E3CE
         Top             =   1440
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image Image15 
         Height          =   960
         Index           =   2
         Left            =   3000
         Picture         =   "桔梗2.frx":7E6A0
         Top             =   1320
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image Image15 
         Height          =   960
         Index           =   1
         Left            =   3120
         Picture         =   "桔梗2.frx":7E980
         Top             =   1200
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image Image15 
         Height          =   960
         Index           =   0
         Left            =   3240
         Picture         =   "桔梗2.frx":7EC9D
         Top             =   1080
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   14
         Left            =   0
         Picture         =   "桔梗2.frx":7EF9A
         Top             =   1680
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   13
         Left            =   120
         Picture         =   "桔梗2.frx":7FEA8
         Top             =   1560
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   12
         Left            =   240
         Picture         =   "桔梗2.frx":80D89
         Top             =   1440
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   11
         Left            =   360
         Picture         =   "桔梗2.frx":81CB5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   10
         Left            =   480
         Picture         =   "桔梗2.frx":82C8B
         Top             =   1200
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   9
         Left            =   600
         Picture         =   "桔梗2.frx":83B34
         Top             =   1080
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   8
         Left            =   720
         Picture         =   "桔梗2.frx":848CE
         Top             =   960
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   7
         Left            =   840
         Picture         =   "桔梗2.frx":856B5
         Top             =   840
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   6
         Left            =   960
         Picture         =   "桔梗2.frx":86496
         Top             =   720
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   5
         Left            =   1080
         Picture         =   "桔梗2.frx":8726B
         Top             =   600
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   4
         Left            =   1200
         Picture         =   "桔梗2.frx":880B4
         Top             =   480
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   3
         Left            =   1320
         Picture         =   "桔梗2.frx":88F36
         Top             =   360
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   2
         Left            =   1440
         Picture         =   "桔梗2.frx":89E38
         Top             =   240
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   1
         Left            =   1560
         Picture         =   "桔梗2.frx":8ACC2
         Top             =   120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image17 
         Height          =   975
         Index           =   0
         Left            =   1680
         Picture         =   "桔梗2.frx":8BC5B
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.Image Image14 
      Height          =   255
      Left            =   5880
      Picture         =   "桔梗2.frx":8CBCE
      Top             =   8760
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image3 
      Height          =   675
      Left            =   2160
      Picture         =   "桔梗2.frx":8CFEC
      Top             =   8880
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image Image32 
      Height          =   450
      Left            =   11040
      Picture         =   "桔梗2.frx":8D518
      Top             =   720
      Width           =   495
   End
   Begin VB.Image Image31 
      Height          =   645
      Left            =   5400
      Picture         =   "桔梗2.frx":8DA29
      Top             =   8400
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  '透明
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  '實心
      Height          =   135
      Index           =   0
      Left            =   6840
      Top             =   8880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image30 
      Height          =   1200
      Index           =   0
      Left            =   8640
      Picture         =   "桔梗2.frx":8DFA1
      Top             =   8280
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image29 
      Height          =   960
      Index           =   0
      Left            =   9960
      Picture         =   "桔梗2.frx":8E297
      Top             =   8520
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image28 
      Height          =   1200
      Index           =   0
      Left            =   11160
      Picture         =   "桔梗2.frx":8E4C9
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image25 
      Height          =   1200
      Index           =   0
      Left            =   12360
      Picture         =   "桔梗2.frx":8E77E
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image27 
      Height          =   1200
      Index           =   5
      Left            =   1080
      Picture         =   "桔梗2.frx":8E9A5
      Top             =   12240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image27 
      Height          =   1200
      Index           =   4
      Left            =   1200
      Picture         =   "桔梗2.frx":8EC9E
      Top             =   12120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image27 
      Height          =   1200
      Index           =   3
      Left            =   1320
      Picture         =   "桔梗2.frx":8EF97
      Top             =   12000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image27 
      Height          =   1200
      Index           =   2
      Left            =   1440
      Picture         =   "桔梗2.frx":8F276
      Top             =   11880
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image27 
      Height          =   1200
      Index           =   1
      Left            =   1560
      Picture         =   "桔梗2.frx":8F56F
      Top             =   11760
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image27 
      Height          =   1200
      Index           =   0
      Left            =   1680
      Picture         =   "桔梗2.frx":8F868
      Top             =   11640
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image26 
      Height          =   960
      Index           =   4
      Left            =   2880
      Picture         =   "桔梗2.frx":8FB5E
      Top             =   12480
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image26 
      Height          =   960
      Index           =   3
      Left            =   3000
      Picture         =   "桔梗2.frx":8FD3D
      Top             =   12360
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image26 
      Height          =   960
      Index           =   2
      Left            =   3120
      Picture         =   "桔梗2.frx":8FF37
      Top             =   12240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image26 
      Height          =   960
      Index           =   1
      Left            =   3240
      Picture         =   "桔梗2.frx":90135
      Top             =   12120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image26 
      Height          =   960
      Index           =   0
      Left            =   3360
      Picture         =   "桔梗2.frx":90339
      Top             =   12000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image24 
      Height          =   1200
      Index           =   3
      Left            =   4680
      Picture         =   "桔梗2.frx":9056B
      Top             =   12240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image24 
      Height          =   1200
      Index           =   2
      Left            =   4800
      Picture         =   "桔梗2.frx":9074A
      Top             =   12120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image24 
      Height          =   1200
      Index           =   1
      Left            =   4920
      Picture         =   "桔梗2.frx":90983
      Top             =   12000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image24 
      Height          =   1200
      Index           =   0
      Left            =   5040
      Picture         =   "桔梗2.frx":90B62
      Top             =   11880
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image23 
      Height          =   1200
      Index           =   5
      Left            =   6480
      Picture         =   "桔梗2.frx":90E17
      Top             =   12240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image23 
      Height          =   1200
      Index           =   4
      Left            =   6600
      Picture         =   "桔梗2.frx":91045
      Top             =   12120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image23 
      Height          =   1200
      Index           =   3
      Left            =   6720
      Picture         =   "桔梗2.frx":91274
      Top             =   12000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image23 
      Height          =   1200
      Index           =   2
      Left            =   6840
      Picture         =   "桔梗2.frx":914A4
      Top             =   11880
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image23 
      Height          =   1200
      Index           =   1
      Left            =   6960
      Picture         =   "桔梗2.frx":916D3
      Top             =   11760
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image23 
      Height          =   1200
      Index           =   0
      Left            =   7080
      Picture         =   "桔梗2.frx":91901
      Top             =   11640
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "★"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   6840
      TabIndex        =   30
      Top             =   8280
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image22 
      Height          =   255
      Left            =   120
      Top             =   9600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image20 
      Height          =   765
      Left            =   10560
      Picture         =   "桔梗2.frx":91B28
      Top             =   7560
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "魔王："
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   405
      Left            =   3960
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "桔梗:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   405
      Left            =   3960
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "P 暫停中"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Image18 
      Height          =   975
      Index           =   0
      Left            =   10680
      Picture         =   "桔梗2.frx":92E48
      Top             =   6120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image Image16 
      Height          =   960
      Left            =   11760
      Picture         =   "桔梗2.frx":93DBB
      Top             =   6000
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "第 1 關"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11760
      TabIndex        =   22
      Top             =   720
      Width           =   1470
   End
   Begin VB.Label Label16 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11160
      TabIndex        =   21
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "靈力值"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2040
      TabIndex        =   19
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "生命值"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   18
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "★"
      ForeColor       =   &H00C0E0FF&
      Height          =   180
      Left            =   1320
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "★"
      ForeColor       =   &H00C0E0FF&
      Height          =   180
      Left            =   1680
      TabIndex        =   16
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "★"
      ForeColor       =   &H00C0E0FF&
      Height          =   180
      Left            =   2040
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Image13 
      Height          =   675
      Index           =   0
      Left            =   0
      Picture         =   "桔梗2.frx":940B8
      Top             =   5400
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "0 X"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10320
      TabIndex        =   14
      Top             =   720
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   6000
      Picture         =   "桔梗2.frx":94FFB
      Top             =   3360
      Width           =   2085
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "★"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   2760
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "★"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   3120
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "★"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   3480
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "X 1"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12840
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "X 3"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   630
   End
   Begin VB.Shape mp 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   1200
      Top             =   600
      Width           =   2700
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   1200
      Top             =   600
      Width           =   2700
   End
   Begin VB.Shape hp 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   1200
      Top             =   120
      Width           =   2700
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   1200
      Top             =   120
      Width           =   2700
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   10080
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   10080
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image6 
      Height          =   765
      Left            =   0
      Picture         =   "桔梗2.frx":97498
      Top             =   6960
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   0
      Picture         =   "桔梗2.frx":987AD
      Top             =   6360
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   26.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   525
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "擊破數 0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11520
      TabIndex        =   3
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  '實心
      Height          =   375
      Left            =   2640
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   840
      Index           =   0
      Left            =   120
      Picture         =   "桔梗2.frx":98EAD
      Top             =   7920
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Image Image19 
      Height          =   2775
      Left            =   10920
      Picture         =   "桔梗2.frx":9DDAF
      Top             =   4800
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Image Image9 
      Height          =   1500
      Left            =   1920
      Picture         =   "桔梗2.frx":A0249
      Top             =   7560
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Menu file 
      Caption         =   "檔案"
      Begin VB.Menu re 
         Caption         =   "重新遊戲"
         Shortcut        =   {F2}
      End
      Begin VB.Menu rese 
         Caption         =   "選擇難度"
         Shortcut        =   {F3}
      End
      Begin VB.Menu reading 
         Caption         =   "讀取進度"
      End
      Begin VB.Menu save 
         Caption         =   "儲存進度"
      End
      Begin VB.Menu rage 
         Caption         =   "排行榜"
      End
      Begin VB.Menu exit 
         Caption         =   "結束遊戲"
      End
   End
   Begin VB.Menu setup 
      Caption         =   "設定"
      Begin VB.Menu setupkey 
         Caption         =   "鍵盤按鍵設定"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu read 
      Caption         =   "說明"
      Begin VB.Menu readme 
         Caption         =   "遊戲改版記錄"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ast(3) As Integer '=-----可讓(0)箭矢(1)破魔(2)破邪再射出一發(3)鏡分身之破魔限制一發
Dim k(6) As Integer '=-------～畫格切換～(0)破魔之箭(1)咒術飭回(2)磚塊反光動畫(3)破邪(4)鳥的動畫切換(5)能量球的動畫(6)鏡分身之破魔
Dim n(4) As Integer '=----------(0)擊破數(1)四魂之玉數量(2)限制過關判斷的執行次數(3)魔王對話方塊之暫存(4)限制魔王計時器的執行
Dim ns(5) As Integer '能量球的出現數量
Dim ppm() As Integer  '(0)體力(1)破掉的磚塊(2)磚塊的命(3)不做磚塊特效(4)上下反向(5)左右反向(6)打擊特效(7)磚塊重生特效(修正版)
Dim reset As Integer '讀取"載入表單"的次數(開啟程式之外，只要有讀到"載入表單"則執行reset下的程式碼，反之不讀取)
Dim life(1) As Integer '(0)桔梗的命(1)磚塊的命
Dim sleep As Integer '按Enter的次數為0時執行為1時暫停
Dim ha As Integer '破魔
Dim ki As Integer '咒術飭回
Dim kiz As Integer '桔梗損血的最大值
Dim mpwer(1) As Integer '第三顆星的"威力"加成
Dim zxcv As Integer '當zxcv為1時鎖定鍵盤
Dim savX3 As Integer '星星數量
Dim mov(4) As Integer '如果為1時則為1的移動量增加(走→跑)
Dim ad As Integer '大絕招可使用的次數
Dim asdq(3) As Integer '這三個同時為1時才能啟動破邪
Dim pps As Integer '=--------被打到的磚塊(敵人)
Dim ssb As Integer '四魂之玉閃爍次數暫存
Dim color As Integer '大絕招替換顏色
Dim ma As Integer '磚塊的數量
Dim nex As Integer '為1時第二關
Dim ell(1) As Integer '反向(0)上下(1)左右
Dim nalaku As Integer '鳥的命
Dim mas As Integer '能量球的最大產生數量
Dim cpu As Integer '記錄是否為遊戲中(影響到表單的啟動動畫)
Dim king As Integer '記錄為那一關的魔王
Dim speed As Integer '小鳥啄擊的速度(依cpu的速度而定)
Dim timers12345 As Integer '第六關的５次計算
Dim stas As Integer '星際背景之暫停的標記
Dim ppaa  '產生幾顆起始(迴圈)打擊特效之數量
Dim ppbb '產生幾顆結束(迴圈)
Dim ppcc  '產生幾顆增值
Dim ppdd  '產生幾顆增值
Dim pda  'pda=ppaa
Dim pdb
Dim pa
Dim pb
Dim pc
Dim pd '同上
Dim dx
Dim rlv '難度判別(0)輕鬆(1)簡單(2)普通(3)困難
Dim hpw As Integer 'HP和MP的長度
Dim svs(2) As Integer '防止(0)打擊特效過多造成lag(1)受到小鳥啄擊的桔梗連續損血(2)修正破邪箭的怪力...
Dim people As Integer '角色判別
Private Sub Form_Activate() '表單啟動●
Call asc(a)
    e = ss
Call asc(s)
    f = ss
Call asc(d)
    g = ss
Call asc(enter)
    h = ss
Label5.Caption = "註：暫停 (" & h & ")" & vbCrLf & " 普通射擊 (" & f & ")" & vbCrLf & " 咒術飭回 (" & g & ") 可反彈磚塊" & vbCrLf & "破魔之箭 (" & e & ") 具穿刺功能" & vbCrLf & "遊戲中的東西可累加" & vbCrLf & "一顆星，靈力損耗減少" & vbCrLf & "兩顆星，生命值損耗減少" & vbCrLf & "三顆星，破壞力加強" & vbCrLf & "建議先從簡單開始玩" & vbCrLf & "可連續玩同一個難度" & vbCrLf & "※ 少一條命星星會減一顆" & vbCrLf & "※ 沒命時，一切重來"
If cpu = 0 Then
    cpu = 1
    Command1.Enabled = False
    Form1.Left = 0
    For f = 1 To 2145 Step 115 '=-----------表單X軸
        Form1.Left = f
        DoEvents
    Next
    Command1.Enabled = True
End If
End Sub
Private Sub ttms() '磚塊移除共用●
Timer31.Enabled = False
Timer2.Enabled = False
Timer5.Enabled = False

'表單回復動畫(已刪除)
'-----------------------------------------------移除磚塊

For f = 1 To ma
    Image4(f).Visible = False: Unload Image4(f)
    Image13(f).Visible = False '破邪隱藏
    Label20(f).Visible = False: Unload Label20(f)
Next
Timer1.Enabled = False: Timer2.Enabled = False: Timer3.Enabled = False
Timer4.Enabled = False: Timer6.Enabled = False: Timer7.Enabled = False
Timer8.Enabled = False: Timer9.Enabled = False: Timer10.Enabled = False
Timer13.Enabled = Falset: Timer22.Enabled = False: Timer23.Enabled = False: Timer24.Enabled = False
Timer19.Enabled = False: Timer20.Enabled = False: Timer21.Enabled = False '魔王計時器類
Timer25.Enabled = False: Timer27.Enabled = False: Timer28.Enabled = False: Timer29.Enabled = False: Timer30.Enabled = False: Timer31.Enabled = False
For f = 1 To mas '隱藏能量球
    Image18(f).Left = 0 - Image18(0).Width
Next
For f = 1 + pc * 8 To ppbb
    Shape5(f).Visible = False: Unload Shape5(f)
Next

Image2.Visible = False: Image6.Visible = False: Image3.Visible = False: Image20.Visible = False
Call Form_Load
End Sub
Private Sub start() '位置共用 ~ 簡單 ~ 普通●

Image1.Top = 224 '=--------------------------------桔梗
Image1.Left = 400

End Sub
Private Sub Form_Load() '表單載入●
setupkey.Enabled = True
readme.Enabled = True '=--------桔梗關於我(說明的功能表列)
re.Enabled = False
rese.Enabled = False
If reset = 1 Then '=-----------↓重新載入時
    Call start '=---------------------------------------將人物屬性位置調回來
    Label6.Visible = False '=---------------------------暫停中消失
    Image1.Picture = Image7(0)
    ssb = 0: pps = 0: nex = 0: nalaku = 0: king = 0: sleep = 0: stas = 0: dx = 0: ppaa = 0: ppbb = 0: ppcc = 0: ppdd = 0: pda = 0: pdb = 0: pa = 0: pb = 0: pc = 0: pd = 0
    Erase svs
    Erase n '清空陣列
    Erase asdq
    Erase ast
    Erase k
    Erase ppm
    Erase ns
    Label17 = "第 1 關"
    Label1.Caption = ""
    Label2.Caption = "擊破數 " & n(0)
    Shape2.Visible = True
    hp.Visible = True
    Picture1.Visible = True '---↑重新載入時
    Call loads
    Label3.Caption = n(1) & " X"
    Label4.Caption = "X " & life(0)
    Image16.Visible = False '鳥消失
    Timer19.Interval = 850 '鳥的話之計時器的將時間調回原來的樣子
    Label18.Visible = False '鳥的對話消失
    Label19.Visible = False '桔梗的對話消失
    Form1.Picture = Image22.Picture '還原背景圖
    Label4.ForeColor = &H80000012: Label8.ForeColor = &H80000012: Label6.ForeColor = &H80000012: Label3.ForeColor = &H80000012
    Label2.ForeColor = &H80000012: Label17.ForeColor = &H80000012: Label16.ForeColor = &H80000012: Label1.ForeColor = &HC00000
Else '=----------------------------------------------------------------------""僅第一次執行本程式時使用""↓
    hpw = 180 '桔梗的血
    Image22.Picture = Form1.Picture '將背景暫存起來，方便以後可以還原
    Select Case clock '依cpu的頻率而調整小鳥啄擊的速度單位Mhz
        Case Is < 1500
            speed = -20 '快
        Case 1500 To 3000
            speed = -10 '中
        Case 3000 To 4000
            speed = -5 '慢
        Case Is > 4000
            speed = -1 '特慢
    End Select
    reset = 1
    Open "SAVE\Keycode.sav" For Append As #1 '=---載入按鍵碼↓
    Close #1
    Open "SAVE\keycode.sav" For Input As #1
    If Not EOF(1) Then Input #1, Key_up, key_down, key_left, key_right, a, s, d, enter
    Close #1 '=-----------------------------------載入按鍵碼↑
    Call loads '=-----------讀取下一個程序
    Call Option1_Click(0) '讀取輕鬆
    Label4.Caption = "X " & life(0)
    For f = 1 To ma
        Load Image13(f)
    Next
    mas = 72
    For f = 1 To mas
        Load Image18(f)
        Image18(f).Left = 0 - Image18(f).Width
    Next
End If
End Sub
Private Sub loads() '表單讀取共用●
If life(0) = 0 Then '=------------------------------如果掛了則
    life(0) = 3 '=----------------------------------桔梗命為3
    savX3 = 0
    Label16.Caption = 0
    Call stars '讀取星星能力
    Label9.Visible = False: Label10.Visible = False: Label11.Visible = False
    Shape3.Visible = False
    Label4.Caption = "X " & life(0): ad = 0
    n(1) = 0: Label3.ForeColor = &H0& 'n(1)四魂之玉數量
    Label7.Visible = False: Label12.Visible = False: Label13.Visible = False
    Label3.Caption = n(1) & " X"
End If
ell(1) = 1 '磚瑰向左飛
Image19.Left = 0 - Image19.Width * 5
Image16.Left = 0 - Image16.Width * 5
Form1.Left = 915 + 1000 '=------------------------前遊戲ing值,前+後初始值
Form1.Top = 465 + 1000 '=-------------------------前遊戲ing值,前+後初始值
Form1.Height = 7665 + 255 '=----------------------前無功能表列時,前+後初始值
Form1.Width = 10845
Picture1.Left = 0
Picture1.Top = 0

'桔梗血條和法條的位置
hp.Width = hpw
mp.Width = hpw
Shape4.Width = hp.Width
Shape6.Width = Shape4.Width
Label15.Left = (hp.Left + hp.Width \ 2) - Label15.Width \ 2
Label14.Left = Label15.Left
Label4.Left = hp.Left + hp.Width + 12

End Sub
Private Sub Option2_Click(index As Integer) '選擇人物●
Select Case index
    Case 0 '桔梗
        Image33.Picture = Image34.Picture
        people = 0
        
        For f = 1455 To 3255 '動畫
            Option2(0).Width = f
            Option2(1).Width = 4710 - f
            Option2(1).Left = f + 105
            DoEvents
        Next
    Case 1 '亞爾麗塔
        Image33.Picture = Image35.Picture
        people = 1
        For f = 1455 To 3255 '動畫
            Option2(1).Width = f
            Option2(0).Width = 4710 - f
            Option2(1).Left = 4830 - f
            DoEvents
        Next
End Select
End Sub
Private Sub Option1_Click(index As Integer) '難度選擇●
Select Case index
    Case 0 '輕鬆●
        ma = 10: Timer2.Interval = 70: Y(0) = 10: Y(1) = 7: Y(2) = 30: life(1) = 0  '=---y(0)磚血----y(1)磚被傷最小值----y(2)磚被傷最大值----life(1)----磚塊的命
        rlv = 0: Call express(0)
    Case 1 '簡單●
        ma = 10: Timer2.Interval = 70: Y(0) = 20: Y(1) = 5: Y(2) = 20: life(1) = 0
        rlv = 1: Call express(1)
    Case 2 '普通●
        ma = 10: Timer2.Interval = 50: Y(0) = 50: Y(1) = 3: Y(2) = 10: life(1) = 1
        rlv = 2: Call express(2)
    Case 3 '困難●
        ma = 10: Timer2.Interval = 50: Y(0) = 50: Y(1) = 1: Y(2) = 5: life(1) = 2
        rlv = 3: Call express(3)
End Select
End Sub
Private Sub express(ix) '難度選擇之動畫●
For f = 0 To 3
    If Option1(f).Width <> 1095 Then betn = f
Next
Select Case ix
    Case 0
        Select Case betn
            Case 1
                For f = 1455 To 1095 Step -1
                    Option1(1).Width = f '原來的縮回去
                    Option1(1).Left = 2655 - f
                    Option1(0).Width = 2550 - f '放大寬度
                    DoEvents
                Next
            Case 2
                For f = 1455 To 1095 Step -1
                    Option1(2).Width = f '原來的縮回去
                    Option1(2).Left = 3735 - f
                    Option1(1).Left = 2655 - f
                    Option1(0).Width = 2550 - f '放大寬度
                    DoEvents
                Next
            Case 3
                For f = 1455 To 1095 Step -1
                    Option1(3).Width = f '原來的縮回去
                    Option1(3).Left = 4815 - f '原來的推回去
                    Option1(2).Left = 3735 - f '其它的推過去
                    Option1(1).Left = 2655 - f '其它的推過去
                    Option1(0).Width = 2550 - f '拉開
                    DoEvents
                Next
        End Select
    Case 1
        Select Case betn
            Case 0
                For f = 1455 To 1095 Step -1
                    Option1(0).Width = f '原來的縮回去
                    Option1(1).Left = f + 120 '其它的推過去
                    Option1(1).Width = 2550 - f '拉開
                    DoEvents
                Next
            Case 2
                For f = 1455 To 1095 Step -1
                    Option1(2).Width = f '原來的縮回去
                    Option1(2).Left = 3735 - f
                    Option1(1).Width = 2550 - f '放大寬度
                    DoEvents
                Next
            Case 3
                For f = 1455 To 1095 Step -1
                    Option1(3).Width = f '原來的縮回去
                    Option1(3).Left = 4815 - f '原來的推回去
                    Option1(2).Left = 3735 - f '其它的推過去
                    Option1(1).Width = 2550 - f '拉開
                    DoEvents
                Next
        End Select
    Case 2
        Select Case betn
            Case 0
                For f = 1455 To 1095 Step -1
                    Option1(0).Width = f '原來的縮回去
                    Option1(1).Left = f + 120 '其它的推過去
                    Option1(2).Left = f + 1215 '其它的推過去
                    Option1(2).Width = 2550 - f '拉開
                    DoEvents
                Next
            Case 1
                For f = 1455 To 1095 Step -1
                    Option1(1).Width = f '原來的縮回去
                    Option1(2).Left = f + 1215 '其它的推過去
                    Option1(2).Width = 2550 - f '拉開
                    DoEvents
                Next
            Case 3
                For f = 1455 To 1095 Step -1
                    Option1(3).Width = f '原來的縮回去
                    Option1(3).Left = 4815 - f '原來的推回去
                    Option1(2).Width = 2550 - f '拉開
                    DoEvents
                Next
        End Select
    Case 3
        Select Case betn
            Case 0
                For f = 1455 To 1095 Step -1
                    Option1(0).Width = f '原來的縮回去
                    Option1(1).Left = f + 120 '其它的推過去
                    Option1(2).Left = f + 1215 '其它的推過去
                    Option1(3).Left = f + 2310 '移動位置(核心)
                    Option1(3).Width = 2550 - f '放大寬度(核心)
                    DoEvents
                Next
            Case 1
                For f = 1455 To 1095 Step -1
                    Option1(1).Width = f '原來的縮回去
                    Option1(2).Left = f + 1215 '其它的推過去
                    Option1(3).Left = f + 2310
                    Option1(3).Width = 2550 - f '放大寬度
                    DoEvents
                Next
            Case 2
                For f = 1455 To 1095 Step -1
                    Option1(2).Width = f '原來的縮回去
                    Option1(3).Left = f + 2310
                    Option1(3).Width = 2550 - f '放大寬度
                    DoEvents
                Next
        End Select
End Select
End Sub
Private Sub Command1_Click() '開始●
zxcv = 1 '封印按鍵
Label8.Caption = "X " & life(1)
Label8.Visible = True
Picture1.Visible = False '=------------------------隱藏主選單(可選擇難度的那邊)

Shape1.Visible = False
Shape2.Visible = False

Form1.Width = 924 * 15 '=====寬=====
Form1.Height = 708 * 15 '=====高=====
Form1.PaintPicture Form1.Picture, 0, 0, 924, 708, 0, 0, 924, 708, vbSrcCopy

Shape1.Width = 171 + Y(0)
Shape2.Width = Shape1.Width
Shape1.Visible = True
Shape2.Visible = True

Shape1.Left = Form1.ScaleWidth - Shape1.Width - 90 '=----調整敵人之血的位置
Shape2.Left = Shape1.Left '=------------------------------------調整敵人之血的位置

Timer6.Enabled = True '=-------------四魂出現
Timer1.Enabled = True
Timer13.Enabled = True
Timer23.Enabled = True
Timer31.Enabled = True
End Sub
Private Sub Starr() '星星判斷●
If Label10.Visible = True Then
    savX3 = 3
    Call stars
    Exit Sub
Else
    If Label9.Visible = True Then
        savX3 = 2
        Call stars
        Exit Sub
    Else
        savX3 = 1
        Call stars
        Exit Sub
    End If
End If
End Sub
Private Sub stars() '星星的能力●
If savX3 > 3 Then savX3 = 3
If savX3 = 0 Then
    Label9.Visible = False: Label10.Visible = False: Label11.Visible = False '所有星星消失
    ha = 20 '=----------------------------------破魔所消耗的靈力
    ki = 13 '=----------------------------------咒術飭回所消耗的靈力
    kiz = 3 '=----------------------------------桔梗碰到磚塊損的血量最大值
    mpwer(0) = 0
    Shape3.Visible = False
Else
    If savX3 >= 1 Then '=---一顆星
        Shape3.Visible = True '=---------顯示星星背景
        Label9.Visible = True '=---------一顆星
        Label10.Visible = False
        Label11.Visible = False
        ha = 10 '=-----------------------破魔所消耗的靈力減少
        ki = 7 '=-----------------------咒術飭回所消耗的靈力減少
        kiz = 1
        mpwer(0) = 0
        If savX3 >= 2 Then '=------------兩顆星
            Label10.Visible = True '=--------兩顆星
            Label11.Visible = False
            kiz = 1 '桔梗的損血量最大值減少
            mpwer(0) = 0
            If savX3 >= 3 Then
                Label11.Visible = True '=------------三顆星
                mpwer(0) = 5
                Exit Sub
            End If
        End If
    End If
End If
End Sub
Private Sub Command2_Click() '讀取進度------主選單上的●
Call reading_Click
End Sub
Private Sub Command3_Click() '儲存進度------主選單上的●
Call save_Click
End Sub
Private Sub Command4_Click() '按鍵設定------主選單上的●
Form3.Show
End Sub
Private Sub Command5_Click() '說明----------主選單上的●
Call readme_Click
End Sub
Private Sub Command6_Click() '排行榜--------主選單上的●
Call rage_Click
End Sub
Private Sub rage_Click() '功能表列上的名次記錄
Form4.Show
End Sub
Private Sub re_Click() '功能表列的重新開始●
re.Enabled = False
rese.Enabled = False
zxcv = 1
life(0) = 0
Call ttms
End Sub
Private Sub rese_Click() '功能表列的重新選擇難度●
re.Enabled = False
rese.Enabled = False
zxcv = 1
Call ttms
End Sub
Private Sub reading_Click() '讀取進度●
Open "SAVE\桔梗.sav" For Append As #1
Close #1
pqi = Val(Label16.Caption)
Open "SAVE\桔梗.sav" For Input As #1
If Not EOF(1) Then Input #1, savX3, life(0), ad, pqi
Close #1
If life(0) >= 9 Then life(0) = 9
Label4.Caption = "X " & life(0)
Label16.Caption = pqi
Call multiha '破邪可使用之次數
Call stars '星星能力
MsgBox "讀取完畢!", 64, "讀取進度"
End Sub
Private Sub save_Click() '功能表列的儲存進度●
pqi = Val(Label16.Caption)
Open "SAVE\桔梗.sav" For Output As #1
Write #1, savX3, life(0), ad, pqi
Close #1
MsgBox "儲存完畢!", 64, "儲存進度"
End Sub
Private Sub exit_Click() '功能表列的離開程式●
zxcv = 1
End
End Sub
Private Sub setupkey_Click() '功能表列的按鍵設定●
Call Command4_Click
End Sub
Private Sub readme_Click() '功能表列的關於桔梗射擊遊戲●
Form2.Show
End Sub
Private Sub Timer5_Timer() '磚塊特效●
If k(2) = 10 Then
    k(2) = 0
Else
    k(2) = k(2) + 1
End If
For f = 1 To ma
    If zxcv = 1 Then Exit Sub
    If f <> ppm(1, f) And f <> ppm(3, f) Then
        Image4(f).Picture = Image11(k(2)).Picture
        DoEvents
    End If
Next
End Sub
Private Sub Timer13_Timer() '新陳代謝(回魔)●(新式演算法v7.0)

If hp.Width < hpw \ 4 Then
    smp = 40
Else
    smp = 10
End If

If mp.Width < hpw Then '第一階
    mp.Width = mp.Width + smp
    If mp.Width > hpw Then mp.Width = hpw '第二階
End If
End Sub
Private Sub Timer23_Timer() '新陳代謝(回體)●(新式演算法v7.0)
If hp.Width < hpw Then
    hp.Width = hp.Width + 10
    If hp.Width > hpw Then hp.Width = hpw
End If
End Sub
Private Sub Timer1_Timer() '磚塊產生●
ReDim ppm(7, ma)
Randomize
For f = 1 To ma
    If nex = 0 Then
        Load Image4(f)
        Image4(f).Height = Image4(0).Height
        Image4(f).Width = Image4(0).Width
        Load Label20(f)
        Label20(f).Height = Label20(0).Height
        Label20(f).Width = Label20(0).Width
    End If
    Image4(f).Visible = True
    Image4(f).Left = Int(Rnd * Form1.ScaleWidth * 1.5) + Form1.ScaleWidth
    Image4(f).Top = Int(Rnd * (Form1.ScaleHeight - Image4(0).Height - (Label2.Top + Label2.Height))) + Label2.Top + Label2.Height
    Image9.Left = Form1.ScaleWidth + Image9.Width
    Image9.Top = Int(Rnd * (Form1.ScaleHeight - Image9.Height - (Label2.Top + Label2.Height))) + Label2.Top + Label2.Height
    Label20(f).Left = Int(Rnd * Form1.ScaleWidth * 1.5) + Form1.ScaleWidth
    Label20(f).Top = Int(Rnd * (Form1.ScaleHeight - Label20(0).Height))
Next
Image16.Top = Int(Rnd * (Form1.ScaleHeight - Image16.Height - (Label2.Top + Label2.Height))) + Label2.Top + Label2.Height
For i = 1 To ma \ 2
    aa = Int(Rnd * ma) + 1
    ppm(4, aa) = aa
Next
Timer1.Enabled = False
Timer2.Enabled = True
Timer5.Enabled = True
re.Enabled = True
rese.Enabled = True
zxcv = 0 '解除按鍵
End Sub
Private Sub Timer2_Timer() '磚塊移動●(磚塊陣列採用桔梗第二代演算法v4.3,磚塊反彈採用第三代桔梗演算法v6.4.1)
'Image1.Left = Image1.Left - 15 '桔梗移動
If nex >= 5 Then Image14.Left = Image14.Left - 30
If Image1.Left < 0 Then Image1.Left = 0
Randomize
For f = 1 To ma
    If zxcv = 1 Then Exit Sub
    If f <> ppm(1, f) Then
        If nex >= 1 Then '第二關的反彈演算法..
            If nex >= 2 Then '第三關關的反彈演算法..
                If fure4(Image4(f)) Then
                    If ppm(5, f) <> f Then
                        ppm(5, f) = f
                    Else
                        ppm(5, f) = 0
                    End If
                End If
            End If
            If fure1(Image4(f)) Or fure2(Image4(f)) Then
                If ppm(4, f) <> f Then
                    ppm(4, f) = f
                Else
                    ppm(4, f) = 0
                End If
            End If
            If ppm(4, f) <> f Then
                ell(0) = -1
            Else
                ell(0) = 1
            End If
            Image4(f).Top = Image4(f).Top + 40 * ell(0)
        End If
        If fure3(Image4(f)) Then
            ell(1) = 1
            ppm(5, f) = 0
        End If
        If ppm(5, f) <> f Then
            ell(1) = 1
        Else
            ell(1) = -1
        End If
        Image4(f).Left = Image4(f).Left - 30 * ell(1) '=-------------------磚塊移動量
        'If coll(Image4(f), Image1) Or coll(Image1, Image4(f)) Then
        If coll(Image4(f), Image1) Then
            If Image4(f).Left <= Image1.Left + Image1.Width \ 2 And Image4(f).Top >= Image1.Top + 10 Then '=--碰到桔梗本體時
                Call kizzs
                Timer23.Enabled = True
            End If
        End If
        If zxcv = 1 Then Exit Sub
        If Image4(f).Left + Image4(f).Width < 0 And Image4(f).Left > -3000 Then
            Image4(f).Left = Form1.ScaleWidth
            Image4(f).Top = Int(Rnd * (Form1.ScaleHeight - Image4(0).Height - (Label2.Top + Label2.Height))) + Label2.Top + Label2.Height
        End If
        DoEvents
    End If
Next
If Image9.Left + Image9.Width > 0 Then
    If n(0) > 3 Then '1UP
        Image9.Visible = True
        'If coll(Image1, Image9) Or coll(Image9, Image1) Then
        If coll(Image1, Image9) Then
            Image9.Visible = False: Image9.Left = -11111
            If life(0) < 9 Then life(0) = life(0) + 1: Label4.Caption = "X " & life(0)
        End If
        Image9.Left = Image9.Left - 70
        DoEvents
    End If
End If
Select Case n(2)
    Case 0
        If n(0) = ma Then
            n(2) = n(2) + 1
            Call nextclass
            Label17 = "第 " & nex + 1 & " 關"
        End If
    Case 1
        If n(0) = ma * 2 Then '(所有關的總擊破數)
            n(2) = n(2) + 1
            Call nextclass
            Label17 = "第 " & nex + 1 & " 關"
        End If
    Case 2
        If n(0) = ma * 3 Then
            nex = nex + 1
            n(2) = n(2) + 1
            Label17 = "第 " & nex + 1 & " 關"
            Label18.Caption = "魔王:不愧是桔梗～妳是否能過我這關呢"
            Label19.Caption = "桔梗:我的靈力已不再受到限制.."
            Label18.Visible = True '魔王說話
            Timer5.Enabled = False ''停止磚塊反光動畫
            king = 0
            Timer19.Enabled = True '啟動魔王計時器
        End If
    Case 3
        If n(0) = ma * 3 + 1 Then
            n(3) = 0: n(4) = 0
            Timer19.Interval = 850
            nex = nex + 1
            n(2) = n(2) + 1
            Label17 = "第 " & nex + 1 & " 關"
            Label18.Caption = "獨孤寧珂:桔梗...去死吧～ ""鏡分身"""
            Label19.Caption = "桔梗:這個招術是....."
            Label18.Visible = True '魔王說話
            king = 1
            Timer19.Enabled = True '啟動魔王計時器
        End If
    Case 4
        If n(0) = ma * 3 + 2 Then
            Timer19.Enabled = False
            Image20.Left = -11111
            Label18.Caption = "桔梗：那個女人...傳說中的蝙蝠女!"
            Label18.Visible = True
            Label4.ForeColor = &HFFFFFF: Label8.ForeColor = &HFFFFFF: Label6.ForeColor = &HFFFFFF: Label3.ForeColor = &HFFFFFF '改變字顏色為白色
            Label2.ForeColor = &HFFFFFF: Label17.ForeColor = &HFFFFFF: Label16.ForeColor = &HFFFFFF: Label1.ForeColor = &HFFFFFF '改變字顏色為白色
            Timer31.Enabled = False '停止背景捲動
            Form1.Picture = LoadPicture() '清空背景
            Timer27.Enabled = True '啟動５次計算
            n(3) = 0: n(4) = 0
            nex = nex + 1
            n(2) = n(2) + 1
            Label17 = "第 " & nex + 1 & " 關"
            ssheight = Form1.Height
            sswidth = Form1.Width
            ssleft = Form1.Left
            sstop = Form1.Top
        End If
    Case 5
        If n(0) = ma * 3 + 3 Then
            Timer14.Enabled = False: Timer15.Enabled = False: Timer16.Enabled = False: Timer17.Enabled = False '=--走路計時器
            zxcv = 1
                          '--暫停↓
            sleep = 1
            Call desleep(1)
                          '--暫停↑
            n(2) = 0
            nex = 0
            MsgBox "第六關全破～～～", 64, "訊息"
            If Option1(3).Value = True Then MsgBox "你真是太神了! 連困難都破了＠＠", 64, "訊息"
            Call winner '呼叫名次
            Call ttms
            Exit Sub
        End If
End Select
End Sub
Private Sub Timer20_Timer() '鳥的畫格切換●
If k(4) = 7 Then
    k(4) = 0
Else
    k(4) = k(4) + 1
End If
Image16.Picture = Image15(k(4)).Picture
End Sub
Private Sub Timer19_Timer() '魔王計時器●
Select Case n(3)
    Case 1
        Label19.Visible = True: n(3) = n(3) + 1
    Case 2
        Label18.Visible = False: n(3) = n(3) + 1 '魔王對話消失
    Case 3
        Label19.Visible = False
        Select Case king
            Case 0 '鳥
                If n(4) = 0 Then
                    Timer19.Interval = 100
                    Timer20.Enabled = True '鳥的畫格切換
                    Shape2.Width = 171 + Y(0) '先將魔王的血補滿
                    Shape2.Visible = True '鳥的血顯示
                    nalaku = life(1) + 1 '鳥的命
                    Label8.Caption = "X " & nalaku
                    Image16.Visible = True '魔王顥現在表單右
                    For f = Form1.ScaleWidth To Form1.ScaleWidth - Image16.Width * 1.5 Step -1
                        Image16.Left = f
                        DoEvents
                    Next
                    n(4) = 1
                Else
                    Randomize
                    com = Int(Rnd * 6)
                    Select Case com
                        Case 0, 3 '上
                            Call updown1(Image16)
                            Exit Sub
                        Case 1, 4 '下
                            Call updown2(Image16)
                            Exit Sub
                        Case 2, 5  '2是能量球,5是小鳥啄擊
                            If com = 2 Then
                                Select Case rlv
                                    Case 0
                                        If ns(0) = mas Then
                                            ns(0) = 1
                                        Else
                                            ns(0) = ns(0) + 1
                                        End If
                                        Image18(ns(0)).Top = Image16.Top
                                        Image18(ns(0)).Left = Image16.Left - Image18(0).Width + Image18(0).Width \ 4
                                        Image18(ns(0)).Visible = True
                                    Case 1
                                        For f = 0 To 1 '2顆
                                            If ns(f) = mas + (f - 1) Then
                                                ns(f) = f + 1
                                            Else
                                                ns(f) = ns((f + 1) Mod 2) + 1
                                            End If
                                        Next
                                        For i = 0 To 1
                                            Image18(ns(0)).Top = Image16.Top - Image18(0).Height \ 2
                                            Image18(ns(1)).Top = Image16.Top + Image18(0).Height \ 2
                                            Image18(ns(i)).Left = Image16.Left - Image18(0).Width + Image18(0).Width \ 4
                                            Image18(ns(i)).Visible = True
                                        Next
                                    Case 2
                                        For f = 0 To 3 '4顆
                                            If ns(f) = mas + (f - 3) Then
                                                ns(f) = f + 1
                                            Else
                                                If f = 0 Then
                                                    ns(f) = ns(3) + 1
                                                Else
                                                    ns(f) = ns(f - 1) + 1
                                                End If
                                            End If
                                        Next
                                        For i = 0 To 3
                                            If i = 0 Then
                                                Image18(ns(0)).Top = Image16.Top + Image18(0).Height * 2.3
                                            Else
                                                Image18(ns(i)).Top = Image18(ns(i - 1)).Top + Image18(0).Height - 150
                                            End If
                                            Image18(ns(i)).Left = Image16.Left - Image18(0).Width + Image18(0).Width \ 4
                                            Image18(ns(i)).Visible = True
                                        Next
                                    Case 3
                                        For f = 0 To 5 '6顆
                                            If ns(f) = mas + (f - 5) Then
                                                ns(f) = f + 1
                                            Else
                                                If f = 0 Then
                                                    ns(f) = ns(5) + 1
                                                Else
                                                    ns(f) = ns(f - 1) + 1
                                                End If
                                            End If
                                        Next
                                        For i = 0 To 5
                                            If i = 0 Then
                                                Image18(ns(0)).Top = Image16.Top + Image18(0).Height * 3
                                            Else
                                                Image18(ns(i)).Top = Image18(ns(i - 1)).Top + Image18(0).Height - 150
                                            End If
                                            Image18(ns(i)).Left = Image16.Left - Image18(0).Width + Image18(0).Width \ 4
                                            Image18(ns(i)).Visible = True
                                        Next
                                End Select
                                Timer21.Enabled = True '能量球移動啟動
                                Timer22.Enabled = True '能量球的動畫啟動
                                Exit Sub
                            Else
                                If Image16.Top >= Image1.Top + 10 And Image16.Top < Image1.Top + Image1.Height Then
                                    wws = Image16.Left
                                    For f = Image16.Left To 0 Step speed
                                        If Image16.Visible = False Then Exit Sub
                                        Image16.Left = f
                                        'If coll(Image16, Image1) Or coll(Image1, Image16) Then
                                        If coll(Image16, Image1) Then
                                            If Image16.Left <= Image1.Left + Image1.Width \ 2 And Image16.Top >= Image1.Top + 10 Then
                                                If Timer10.Enabled = False Then
                                                    svs(1) = svs(1) + 1
                                                    If svs(1) Mod 3 = 0 Then
                                                        svs(1) = 0
                                                        zz = kiz
                                                        kiz = 3
                                                        Call kizzs '呼叫桔梗損血演算法
                                                        kiz = zz
                                                    End If
                                                    Timer23.Enabled = True
                                                End If
                                            End If
                                        End If
                                        DoEvents
                                    Next
                                    For f = 0 To wws Step 2
                                        Image16.Left = f
                                        DoEvents
                                    Next
                                    Image16.Left = wws
                                    Exit Sub
                                End If
                            End If
                    End Select
                End If
            Case 1 '鏡分身
                If n(4) = 0 Then
                    Timer19.Interval = 100
                    Shape2.Width = 171 + Y(0) '先將魔王的血補滿
                    Shape2.Visible = True '體力顯示
                    nalaku = life(1) + 1 '鏡分身的命
                    Label8.Caption = "X " & nalaku
                    Image19.Visible = True
                    For f = Form1.ScaleWidth To Form1.ScaleWidth - Image19.Width * 1.1 Step -1
                        Image19.Left = f
                    Next
                    n(4) = 1
                Else
                    Randomize
                    com = Int(Rnd * 6)
                    Select Case com
                        Case 0, 3 '上
                            Call updown1(Image19)
                            Exit Sub
                        Case 1, 4 '下
                            Call updown2(Image19)
                            Exit Sub
                        Case 2, 5 '2是破魔
                            If ast(3) = 0 Then
                                ast(3) = 1
                                Image20.Left = Image19.Left - Image20.Width
                                Image20.Top = Image19.Top + Image19.Height \ 2
                                Image20.Visible = True
                                Timer25.Enabled = True '啟動破魔
                                Timer26.Enabled = True '破魔畫格切換
                            End If
                    End Select
                End If
        End Select
Case Else
    n(3) = n(3) + 1
End Select
End Sub
Private Sub updown1(a As Object) '魔王的移動判斷式●
If fure2(a) Then
    a.Top = a.Top - 50 '上
Else
    a.Top = a.Top + 50 '下
End If
End Sub
Private Sub updown2(a As Object) '魔王的移動判斷式●
If a.Top <= Label2.Top + Label2.Height Then
    a.Top = a.Top + 50
Else
    a.Top = a.Top - 50
End If
End Sub
Private Sub Timer21_Timer() '能量球移動●
For f = 1 To mas
    'If coll(Image18(f), Image1) Or coll(Image1, Image18(f)) Then
    If coll(Image18(f), Image1) Then
        If Image18(f).Left <= Image1.Left + Image1.Width \ 2 And Image18(f).Top >= Image1.Top + 10 Then
            Timer13.Enabled = False '靈力回復停止
            mp.Width = mp.Width / 1.5
            zz = kiz
            kiz = kiz / 1.5
            Call kizzs '呼叫桔梗損血演算法
            kiz = zz
            Timer23.Enabled = True
            Timer13.Enabled = False '靈力回復
        End If
    End If
    Image18(f).Left = Image18(f).Left - 50
Next
End Sub
Private Sub Timer22_Timer() '能量球動畫切換●
If k(5) = 14 Then
    k(5) = 0
Else
    k(5) = k(5) + 1
End If
For f = 1 To mas
    Image18(f).Picture = Image17(k(5)).Picture
Next
End Sub
Private Sub Timer25_Timer() '鏡分身-破魔之箭●
Image20.Left = Image20.Left - 50
'If coll(Image1, Image20) Or coll(Image20, Image1) Then
If coll(Image1, Image20) Then
    If Image20.Left <= Image1.Left + Image1.Width \ 2 And Image20.Top >= Image1.Top + 10 Then
        If Timer10.Enabled = False Then
            zz = kiz
            kiz = kiz * 6
            Call kizzs
            kiz = zz
            Timer23.Enabled = True
        Else
            Image20.Left = 0 - Image20.Width: ast(3) = 0: Timer25.Enabled = False: Timer26.Enabled = False
        End If
    End If
End If
If Image20.Left < 0 - Image20.Width Then ast(3) = 0: Timer25.Enabled = False: Timer26.Enabled = False
End Sub
Private Sub Timer26_Timer() '鏡分身-破魔畫格切換●
If k(6) = 2 Then
    k(6) = 0
Else
    k(6) = k(6) + 1
End If
Image20.Picture = Image21(k(6)).Picture
End Sub
Private Sub Timer27_Timer() '第六關的５次計算●New
timers12345 = timers12345 + 1
If timers12345 >= 5 Then
    timers12345 = 0
    For f = 1 To ma '顯示星際
        Label20(f).Visible = True
    Next
    Label18.Visible = False
    Timer28.Enabled = True '啟動星際背景
    Timer27.Enabled = False
End If
End Sub
Private Sub Timer28_Timer() '星際背景●New
For f = 1 To ma
    Label20(f).Left = Label20(f).Left - 50 '星星移動
    If Label20(f).Left + Label20(f).Width < 0 And Label20(f).Left > -3000 Then
        Label20(f).Left = Int(Rnd * Form1.ScaleWidth * 1.5) + Form1.ScaleWidth
        Label20(f).Top = Int(Rnd * (Form1.ScaleHeight - Label20(0).Height))
    End If
    DoEvents
Next
'Image1.Left = Image1.Left - 10 '桔梗跟著移動
'If Image1.Left < 0 Then Image1.Left = 0 '讓桔梗不超過視窗
End Sub
Private Sub nextclass() '關卡共用●
nex = nex + 1
Timer1.Enabled = True '下一關開始
End Sub
Private Sub desleep(dis) '暫停程序共用●
If sleep Mod 2 = 1 Then
    Timer2.Enabled = False: Timer14.Enabled = False: Timer15.Enabled = False: Timer16.Enabled = False
    Timer17.Enabled = False: Timer13.Enabled = False: Timer23.Enabled = False: Timer6.Enabled = False
    Timer9.Enabled = False: Timer8.Enabled = False: Timer7.Enabled = False: Timer10.Enabled = False
    Timer3.Enabled = False: Timer4.Enabled = False: Timer12.Enabled = False: Timer18.Enabled = False
    Timer31.Enabled = False
    If dis = 0 Then Label6.Visible = True '顯示暫停標籤
    If Timer10.Enabled = True Then ssdis = 1
    If nex >= 3 And nex < 5 Then
        Timer19.Enabled = False
        Timer21.Enabled = False
    End If
    If Timer28.Enabled = True Then Timer28.Enabled = False: stas = 1 '星際背景暫停
    If nex = 4 Then Timer25.Enabled = False
    Exit Sub
Else
    sleep = 0
    If ssdis = 1 Then Timer10.Enabled = True
    Timer3.Enabled = True: Timer4.Enabled = True: Timer12.Enabled = True: Timer18.Enabled = True
    Timer2.Enabled = True: Timer13.Enabled = True: Timer23.Enabled = True: Timer6.Enabled = True: Timer7.Enabled = True
    Timer8.Enabled = True: Timer9.Enabled = True: Label6.Visible = False '隱藏暫停標籤
    If n(0) < nex = 5 Then Timer31.Enabled = True
    If nex >= 3 And nex < 5 Then
        Timer19.Enabled = True
        Timer21.Enabled = True
    End If
    If stas = 1 Then Timer28.Enabled = True '星際背景啟動
    If nex = 4 Then Timer25.Enabled = True
    If nex < 5 Then Timer31.Enabled = True
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '按鍵按下●
If zxcv = 0 Then
    If KeyCode = enter Then
        sleep = sleep + 1 '=------Enter暫停
        Call desleep(0)
        Exit Sub
    End If
    If sleep = 0 Then
        Select Case KeyCode
            Case Key_up '=--------按上
                If Timer11(1).Enabled = True Then mov(1) = 1
                If Image1.Top > hp.Top Then Timer16.Enabled = True: Exit Sub
            Case key_down '=------按下
                If Timer11(2).Enabled = True Then mov(2) = 2
                If Image1.Top + Image1.Height <> Form1.ScaleHeight Then Timer17.Enabled = True: Exit Sub
            Case key_left '=------按左
                If Timer11(3).Enabled = True Then mov(3) = 3
                If Image1.Left <> 0 Then Timer14.Enabled = True: Exit Sub
            Case key_right '=-----按右
                If Timer11(4).Enabled = True Then mov(4) = 4
                If Image1.Left + Image1.Width <> Form1.ScaleWidth Then Timer15.Enabled = True: Exit Sub
            Case a '=--------------A破魔
                asdq(0) = 1 '破邪之一
                ast(0) = ast(0) + 1
                If ast(0) = 1 Then
                    If mp.Width >= ha Then '----mp↓減少靈力
                        Timer13.Enabled = False
                        mp.Width = mp.Width - ha ' -----mp↑減少靈力
                        sndPlaySound "data\04.dat", 1
                        Label1.Caption = "～破魔之箭～"
                        Image6.Visible = True
                        Image6.Left = Image1.Left + Image1.Width
                        Image6.Top = Image1.Top + Image1.Height \ 2
                        Timer4.Enabled = True: Exit Sub
                    Else
                        ast(0) = 0
                        Label1.Caption = "靈力不足!!"
                    End If
                End If
            Case s '--------------S射箭
                asdq(1) = 1 '破邪之一
                ast(1) = ast(1) + 1
                If ast(1) = 1 Then
                    sndPlaySound "data\03.dat", 1
                    Image2.Visible = True
                    Image2.Left = Image1.Left + Image1.Width
                    Image2.Top = Image1.Top + Image1.Height \ 2
                    Timer3.Enabled = True: Exit Sub
                End If
            Case d '=--------------D咒術飭回
                asdq(2) = 1 '破邪之一
                If ad >= 1 Then
                    If asdq(0) = 1 And asdq(1) = 1 And asdq(2) = 1 Then
                        ast(2) = ast(2) + 1
                        If ast(2) = 1 Then
                            ad = ad - 1
                            Select Case ad
                                Case 2
                                    Label13.Visible = False
                                Case 1
                                    Label12.Visible = False
                                Case 0
                                    Label7.Visible = False
                                    Timer24.Enabled = False
                            End Select
                            Image13(0).Top = (Image1.Top + Image1.Height \ 2) - Image1.Height * 1.4
                            Image13(0).Left = Image1.Left + Image1.Width
                            For f = 1 To ma
                                Image13(f).Top = Image13(f - 1).Top + Image13(f - 1).Height
                                Image13(f).Left = Image1.Left + Image1.Width
                                Image13(f).Visible = True
                            Next
                            Label1.Caption = "～破邪之箭～"
                            mpwer(1) = 3
                            Timer18.Enabled = True '=---啟動破邪
                            Timer12.Enabled = True '=---破邪切換畫格
                            Exit Sub
                        End If
                    End If
                End If
                If mp.Width >= ki Then '=------------------mp↓減少靈力
                    Timer13.Enabled = False '=-------------暫停回魔
                    mp.Width = mp.Width - ki '=------------mp↑減少靈力
                    Label1.Caption = "～咒術飭回～"
                    Timer10.Enabled = True: Exit Sub
                Else
                    Label1.Caption = "靈力不足!!"
                    Exit Sub
                End If
        End Select
    End If
End If
End Sub
Private Sub Timer11_Timer(index As Integer) '跑步計時器●
Timer11(index).Enabled = False
End Sub
Private Sub Timer16_Timer() '桔梗移動計時器上●-------------------"↑"
If mov(1) = 1 Then
    If Timer14.Enabled = True Then mov(3) = 3 '目的讓斜移平均
    If Timer15.Enabled = True Then mov(4) = 4
    Image1.Top = Image1.Top - 25 '桔梗跑步
Else
    Image1.Top = Image1.Top - 10 '桔梗移動
End If
Call supicokakalasup
If Image1.Top < hp.Top Then Image1.Top = hp.Top: Timer16.Enabled = False '=--固定在上
End Sub
Private Sub Timer17_Timer() '桔梗移動計時器下●-------------------"↓"
If mov(2) = 2 Then
    If Timer14.Enabled = True Then mov(3) = 3
    If Timer15.Enabled = True Then mov(4) = 4
    Image1.Top = Image1.Top + 25 '桔梗跑步
Else
    Image1.Top = Image1.Top + 10 '桔梗移動
End If
Call supicokakalasup
If Image1.Top + Image1.Height > Form1.ScaleHeight Then Image1.Top = Form1.ScaleHeight - Image1.Height: Timer17.Enabled = False '固定在下
End Sub
Private Sub Timer14_Timer() '桔梗移動計時器左●-------------------"←"
If mov(3) = 3 Then
    If Timer14.Enabled = True Then mov(1) = 1
    If Timer15.Enabled = True Then mov(2) = 2
    Image1.Left = Image1.Left - 25 '桔梗跑步
Else
    Image1.Left = Image1.Left - 10 '桔梗移動
End If
Call supicokakalasup
If Image1.Left < 0 Then Image1.Left = 0: Timer14.Enabled = False '=--固定在左
End Sub
Private Sub Timer15_Timer() '桔梗移動計時器右●-------------------"→"
If mov(4) = 4 Then '=-----------------------------當"右"彈起時 *啟動"超過0.18秒"會自動關閉的計時器* 然後在計時器關閉前按下右時,mov(4)=4並跑步,按彈起時..mov(4)=0
    If Timer14.Enabled = True Then mov(1) = 1
    If Timer15.Enabled = True Then mov(2) = 2
    Image1.Left = Image1.Left + 25 '=--------------------------------桔梗跑步
Else
    Image1.Left = Image1.Left + 10 '=--------------------------------桔梗移動
End If
Call supicokakalasup
If Image1.Left + Image1.Width > Form1.ScaleWidth Then Image1.Left = Form1.ScaleWidth - Image1.Width: Timer15.Enabled = False '固定在右
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer) '按鍵彈起●
If KeyCode = Key_up Or KeyCode = key_down Or KeyCode = key_left Or KeyCode = key_right Then
    For f = 1 To 4
        mov(f) = 0
    Next
End If
Select Case KeyCode
    Case Key_up '上
        Timer11(1).Enabled = True
        Timer16.Enabled = False
        Exit Sub
    Case key_down '下
        Timer11(2).Enabled = True
        Timer17.Enabled = False
        Exit Sub
    Case key_left '左
        Timer11(3).Enabled = True
        Timer14.Enabled = False
        Exit Sub
    Case key_right '右
        Timer11(4).Enabled = True
        Timer15.Enabled = False
        Exit Sub
    Case a, s, d
        Timer13.Enabled = True
        Erase asdq '=-----ASD任何一個鍵起來則無法啟動破邪
        Exit Sub
End Select
End Sub
Private Sub Timer18_Timer() '破邪之箭●
For f = 1 To ma
    Image13(f).Left = Image13(f).Left + 50
    Call forimage(Image13(f))
    If Image13(f).Left >= Form1.ScaleWidth Then
        Label1.Caption = ""
        Timer12.Enabled = False: Timer18.Enabled = False
        Image13(f).Visible = False
        k(3) = 0
        ast(2) = 0
        mpwer(1) = 0
    End If
Next
End Sub
Private Sub Timer12_Timer() '破邪之箭切換畫格●
If k(3) = 10 Then
    k(3) = 0
Else
    k(3) = k(3) + 1
End If
For f = 1 To ma
    Image13(f).Picture = Image12(k(3)).Picture
Next
End Sub
Private Sub Timer3_Timer() '箭矢移動●
Image2.Left = Image2.Left + 30
Call forimage(Image2) '呼叫磚塊的碰撞迴圈
If Image2.Left > Form1.ScaleWidth Then Image2.Visible = False: Image2.Left = -111111: ast(1) = 0: Timer3.Enabled = False
End Sub
Private Sub Timer4_Timer() '破魔移動●
If k(0) = 3 Then '=---------------------------------------破魔畫格切換
    k(0) = 0
Else
    k(0) = k(0) + 1
End If
Image6.Picture = Image5(k(0)).Picture
Image6.Left = Image6.Left + 35
Call forimage(Image6) '=--------------------------------呼叫磚塊的碰撞迴圈
If Image6.Left > Form1.ScaleWidth Then Image6.Visible = False: ast(0) = 0: Timer4.Enabled = False: Image6.Left = -11111: Label1.Caption = "" 'at讓破魔可在射一發
End Sub
Private Sub Timer10_Timer() '咒術飭回畫格切換●
If k(1) = 6 Then
    k(1) = 0: Timer10.Enabled = False
    Label1.Caption = ""
Else
    k(1) = k(1) + 1
End If
Image1.Picture = Image7(k(1)).Picture
Call forimage(Image1) '=----------------------------------呼叫磚塊的碰撞迴圈
End Sub
Private Sub forimage(xxx As Object) '磚塊的碰撞迴圈●
If nex = 3 Or nex = 4 Then '魔王損血判斷式
    If nex = 3 Then '鳥
        If xxx = Image1 Then
            Label1.Caption = " 無 作 用 "
        Else
            'If coll(xxx, Image16) Or coll(Image16, xxx) Then
            If coll(xxx, Image16) Then
                If xxx = Image2 Then
                    ast(1) = 0 '=------------------------------讓箭矢可在射一發
                    Image2.Visible = False: Image2.Left = -11111: Timer3.Enabled = False
                End If
                Call power '除了咒術飭回外皆可造成敵人之損傷
            End If
        End If
    Else '鏡分身
        'If coll(xxx, Image19) Or coll(Image19, xxx) Then
        If coll(xxx, Image19) Then
            If xxx.Left >= Image19.Left + Image19.Width \ 2 And xxx.Top >= Image19.Top + 10 Then
                If xxx = Image2 Then
                    ast(1) = 0 '=------------------------------讓箭矢可在射一發
                    Image2.Visible = False: Image2.Left = -11111: Timer3.Enabled = False
                End If
                Call power
            End If
        End If
    End If
Else '下面是磚塊碰撞演算法的上式(power下式)
    For f = 1 To ma
        If zxcv = 1 Then Exit Sub
        If f <> ppm(1, f) Then
            'If coll(xxx, Image4(f)) Or coll(Image4(f), xxx) Then
            If coll(xxx, Image4(f)) Then
                Select Case xxx
                    Case Image1 '=----------------------------咒術飭回
                        If Image4(f).Left >= Image1.Left + Image1.Width \ 4 * 3 Then
                            If nex >= 1 Then
                                If ppm(5, f) = 0 Then
                                    sndPlaySound "data\02.dat", 1
                                    ppm(5, f) = f
                                End If
                            Else
                                sndPlaySound "data\02.dat", 1
                                For ff = Image4(f).Left To Form1.ScaleWidth - Image4(f).Width Step 120
                                    If zxcv = 1 Then Exit Sub
                                    Image4(f).Left = ff
                                    DoEvents
                                Next
                            End If
                        End If
                    Case Image2 '=-----------------------------箭矢
                        ast(1) = 0 '=------------------------------讓箭矢可在射一發
                        Image2.Visible = False: Image2.Left = -11111: Timer3.Enabled = False
                    Case Image13(1) '=-------------------------破邪之箭
                        svs(2) = svs(2) + 1
                        If svs(2) Mod 2 = 0 Then
                            svs(2) = 0
                        Else
                            Exit Sub
                        End If
                End Select
                If xxx <> Image1 Then '=-----------------------破邪和破魔和箭矢共用
                    pps = Image4(f).index
                    Call power
                End If
            End If
        End If
    Next
End If
End Sub
Private Sub Timer6_Timer() '四魂出現●(四魂系列採用桔梗射擊遊戲第一代之演算法v1.5之修正版)
Randomize
Image3.Left = Int(Rnd * (Form1.ScaleWidth - Image3.Width))
Image3.Top = Int(Rnd * (Form1.ScaleHeight - Image3.Height - Label2.Top - Label2.Height)) + Label2.Top + Label2.Height
Image3.Visible = True: Timer6.Enabled = False: Timer7.Enabled = True
End Sub
Private Sub Timer7_Timer() '四魂要消失前的閃爍等待時間●
Timer8.Enabled = True: Timer7 = False
End Sub
Private Sub Timer8_Timer() '四魂閃爍●
ssb = ssb + 1
Image3.Visible = False: Timer9.Enabled = True
If ssb = 7 Then Timer6.Enabled = True: Image3.Left = -11111: Timer9.Enabled = False: Timer8.Enabled = False: ssb = 0
End Sub
Private Sub Timer9_Timer() '四魂閃爍●
Image3.Visible = True
End Sub
Private Sub supicokakalasup() '有無碰到四魂之玉●
'If coll(Image1, Image3) Or coll(Image3, Image1) Then Call sicokakala '如果有碰到四魂之玉則
If coll(Image1, Image3) Then Call sicokakala '如果有碰到四魂之玉則
End Sub
Private Sub sicokakala() '吃到四魂之玉●
n(1) = n(1) + 1
Label3.Caption = n(1) & " X"
Image3.Visible = False: Image3.Left = -11111
mp.Width = mp.Width + 50
If mp.Width >= hpw Then
    mp.Width = hpw
End If
If n(1) Mod 3 = 0 Then
    Timer24.Enabled = True
    ad = ad + 1
    Label3.ForeColor = &H80FFFF '=---------"四魂之玉標籤"頻色改變
    Call multiha
Else
    Label3.ForeColor = &H80000012
End If
Call getstart
End Sub
Private Sub multiha() '破邪可使用的次數●
If ad > 3 Then ad = 3
If ad >= 1 Then
    Label7.Visible = True
    If ad >= 2 Then
        Label12.Visible = True
        If ad >= 3 Then
            Label13.Visible = True
        End If
    End If
End If
End Sub
Private Sub Timer24_Timer() '星星閃爍●
If Label7.Visible = True Then
    color = color + 1
    If color >= 5 Then color = 0
    ti = color Mod 4
    Select Case ti
        Case 0
            Label7.ForeColor = &HFF8080
            Label12.ForeColor = &HFF8080
            Label13.ForeColor = &HFF8080
        Case 1
            Label7.ForeColor = &H80FF80
            Label12.ForeColor = &H80FF80
            Label13.ForeColor = &H80FF80
        Case 2
            Label7.ForeColor = &HC0FFFF
            Label12.ForeColor = &HC0FFFF
            Label13.ForeColor = &HC0FFFF
        Case 3
            Label7.ForeColor = &HC0E0FF
            Label12.ForeColor = &HC0E0FF
            Label13.ForeColor = &HC0E0FF
    End Select
End If
End Sub
Private Sub kizzs() '桔梗損血演算法●
Timer23.Enabled = False
X = Int(Rnd * kiz) + 5
If hp.Width <= X Then
    If life(0) <> 0 Then
        If nex = 3 Then
            Timer19.Enabled = False
            Timer21.Enabled = False
        End If
        sleep = 1
        Call desleep(1)
        life(0) = life(0) - 1
        Label4.Caption = "X " & life(0)
        hp.Width = hpw
        mp.Width = hpw
        topp = Image1.Top
        zxcv = 1
        re.Enabled = False '預防Bug
        rese.Enabled = False '預防Bug
        For ff = 0 - Image1.Height To topp
            Image1.Top = ff
            DoEvents
        Next
        re.Enabled = True '預防Bug
        rese.Enabled = True '預防Bug
        zxcv = 0
        If savX3 = 3 Then '=--三顆星則
            savX3 = 2 '=-------第三顆星消失
            Call stars
        Else
            If savX3 = 2 Then '兩顆星則
                savX3 = 1 '=---第二顆星消失
                Call stars
            Else
                savX3 = 0 '=----第一顆星消失
                Call stars
            End If
        End If
        If nex = 3 Then
            Timer19.Enabled = True
            Timer21.Enabled = True
        End If
        sleep = 0
        Call desleep(1)
    Else
        Timer14.Enabled = False: Timer15.Enabled = False: Timer16.Enabled = False: Timer17.Enabled = False '走路計時器
        hp.Visible = False
        Timer13.Enabled = False
        Timer23.Enabled = False
        zxcv = 1
        '--暫停↓
        sleep = 1
        Call desleep(1)
        '--暫停↑
        MsgBox "你掛了～～～", 64, "訊息"
        Call winner '呼叫名次
        Call ttms
        Exit Sub
    End If
Else
    hp.Width = hp.Width - X
End If
End Sub
Private Sub power() '磚塊損血演算法●
'shape1 (紅色)敵人體力之底
'shape2 (黃色)敵人之體力
'ppm(0,pps) 體力
'ppm(1,pps) 破掉的磚塊
'ppm(2,pps) 磚塊的命
'ppm(3,pps) 不做磚塊特效
'ppm(7,pps) 磚塊重生特效(修正版)
If nex = 3 Or nex = 4 Then '=--------------------------------------------魔王(鳥)損血演算法
    If nex = 3 Then '鳥
        Randomize
        X = (Int(Rnd * Y(2)) + Y(1) + mpwer(0) + mpwer(1)) \ 2
        If Shape2.Width <= X Then
            If nalaku > 0 Then
                Timer19.Enabled = False
                nalaku = nalaku - 1
                Label8.Caption = "X " & nalaku
                Shape2.Width = 171 + Y(0)
                topp = Int(Rnd * (Form1.ScaleHeight - Image16.Height - (Label2.Top + Label2.Height))) + Label2.Top + Label2.Height
                For f = 0 - Image16.Height To topp Step 35
                    Image16.Top = f
                    DoEvents
                Next
                Timer19.Enabled = True
            Else
                Timer19.Enabled = False
                Timer19.Interval = 1200
                Shape2.Visible = False鳥
                Image16.Visible = False: Image16.Left = 0 - Image16.Width
                n(0) = n(0) + 1
                Label2.Caption = "擊破數 " & n(0)
            End If
        Else
            Shape2.Width = Shape2.Width - X
            Call getstart
            Call vvss(Image16) '打擊特效演算法
            DoEvents
        End If
    Else '=--------------鏡分身
        Randomize
        X = (Int(Rnd * Y(2)) + Y(1) + mpwer(0) + mpwer(1)) \ 2
        If Shape2.Width <= X Then
            If nalaku > 0 Then
                Timer19.Enabled = False
                nalaku = nalaku - 1
                Label8.Caption = "X " & nalaku
                Shape2.Width = 171 + Y(0)
                topp = Int(Rnd * (Form1.ScaleHeight - Image19.Height - (Label2.Top + Label2.Height))) + Label2.Top + Label2.Height
                For f = 0 - Image19.Height To topp Step 35
                    Image19.Top = f
                    DoEvents
                Next
                Timer19.Enabled = True
            Else
                Shape2.Visible = False
                Image19.Visible = False
                Image14.Left = Image19.Left + Image19.Width \ 2
                Image14.Top = Image19.Top + Image19.Height \ 2
                Image14.Visible = True
                
                Image19.Visible = False: Image19.Left = 0 - Image19.Width
                Timer21.Enabled = False
                Timer25.Enabled = False
                Timer26.Enabled = False
                n(0) = n(0) + 1
                Label2.Caption = "擊破數 " & n(0)
            End If
        Else
            Shape2.Width = Shape2.Width - X
            Call getstart
            Call vvss(Image19) '打擊特效演算法
            DoEvents
        End If
    End If
Else '=-------------------------------------------------------------磚塊損血演算法
    If ppm(0, pps) = 0 Then
        Shape2.Width = 171 + Y(0)
        ppm(2, pps) = life(1) '=-----------磚塊命
        Label8.Caption = "X " & ppm(2, pps)
    Else
        Shape2.Width = ppm(0, pps)
    End If
    Shape2.Visible = True
    Randomize
    X = Int(Rnd * Y(2)) + Y(1) + mpwer(0) + mpwer(1)
    If Shape2.Width <= X Then
        If ppm(2, pps) > 0 Then
            Shape2.Width = 171 + Y(0)
            ppm(0, pps) = 171 + Y(0)
            ppm(2, pps) = ppm(2, pps) - 1
            Label8.Caption = "X " & ppm(2, pps) '=---------------顯示敵人的命
            Image4(pps).Left = Form1.ScaleWidth - Image4(0).Width
            topp = Int(Rnd * (Form1.ScaleHeight - Image4(0).Height - (Label2.Top + Label2.Height))) + Label2.Top + Label2.Height
            ppm(3, pps) = 0
            Image4(pps).Picture = Image4(0).Picture
            ppm(7, axc) = pps
            For f = 0 - Image4(0).Height To topp Step 5
                Image4(ppm(7, axc)).Top = f
            Next
        Else
            Label8.Caption = "X " & ppm(2, pps)
            n(0) = n(0) + 1: Label2.Caption = "擊破數 " & n(0)
            Shape2.Visible = False
            ppm(1, pps) = pps
            Image4(ppm(1, pps)).Visible = False: Image4(ppm(1, pps)).Left = -11111
            Shape2.Width = 171 + Y(0)
            ppm(6, pps) = pps
            Call vvss(0) '打擊特效演算法
            Exit Sub
        End If
    Else
        Call getstart '打擊分數演算法
        Label8.Caption = "X " & ppm(2, pps) '=--------------------顯示敵人的命
        Shape2.Width = Shape2.Width - X
        ppm(0, pps) = Shape2.Width
        Select Case Shape2.Width
            Case Is <= (171 + Y(0)) \ 2 And Shape2.Width > (171 + Y(0)) \ 3
                Image4(pps).Picture = Image10(1).Picture
                ppm(3, pps) = pps
            Case Is <= (171 + Y(0)) \ 3
                Image4(pps).Picture = Image10(2).Picture
                ppm(3, pps) = pps
        End Select
        
        
        If Timer18.Enabled = True Or Timer4.Enabled = True Then
            svs(0) = svs(0) + 1
            If svs(0) Mod (rlv + 3) = 0 Then
                svs(0) = 0
                ppm(6, pps) = pps
                Call vvss(0) '打擊特效演算法
            End If
        Else
            ppm(6, pps) = pps
            Call vvss(0) '打擊特效演算法
        End If
        DoEvents
    End If
End If
End Sub
Private Sub vvss(axx) '打擊特效即時演算法●New (打擊特效系列採用桔梗第四代演算法v7.0)

ppaa = 1 + ppcc * 8
ppbb = 8 + ppdd * 8
Randomize

If nex >= 3 Then '如果是第四關之後(魔王)的話
    For f = ppaa To ppbb
        Load Shape5(f)
        
        Shape5(f).Left = axx.Left + axx.Width \ 2
        Shape5(f).Top = Int(Rnd * ((axx.Top + axx.Height) - axx.Top)) + axx.Top
        rrr = Int(Rnd * 255)
        ggg = Int(Rnd * 255)
        bbb = Int(Rnd * 255)
        Shape5(f).FillColor = RGB(rrr, ggg, bbb)
        Shape5(f).Visible = True
        pdb = ppbb
        
    Next
Else '以下為磚塊特效
    For f = ppaa To ppbb
        Load Shape5(f)
        
        Shape5(f).Left = Image4(ppm(6, pps)).Left + Image4(ppm(6, pps)).Width \ 2
        Shape5(f).Top = Int(Rnd * ((Image4(ppm(6, pps)).Top + Image4(ppm(6, pps)).Height) - Image4(ppm(6, pps)).Top)) + Image4(ppm(6, pps)).Top
        rrr = Int(Rnd * 255)
        ggg = Int(Rnd * 255)
        bbb = Int(Rnd * 255)
        Shape5(f).FillColor = RGB(rrr, ggg, bbb)
        Shape5(f).Visible = True
        pdb = ppbb
        
    Next
End If
ppcc = ppcc + 1
ppdd = ppdd + 1


Timer29.Enabled = True
Timer30.Enabled = True
End Sub
Private Sub Timer29_Timer() '打擊特效●New
For f = pda To pdb
    os = f Mod 8
    Select Case os
        Case 0 '上
            Shape5(f).Top = Shape5(f).Top - 60
        Case 1 '下
            Shape5(f).Top = Shape5(f).Top + 60
        Case 2 '左
            Shape5(f).Left = Shape5(f).Left - 60
        Case 3 '右
            Shape5(f).Left = Shape5(f).Left + 60
        Case 4 '左上
            Shape5(f).Top = Shape5(f).Top - 60
            Shape5(f).Left = Shape5(f).Left - 60
        Case 5 '左下
            Shape5(f).Left = Shape5(f).Left - 60
            Shape5(f).Top = Shape5(f).Top + 60
        Case 6 '右上
            Shape5(f).Left = Shape5(f).Left + 60
            Shape5(f).Top = Shape5(f).Top - 60
        Case 7 '右下
            Shape5(f).Left = Shape5(f).Left + 60
            Shape5(f).Top = Shape5(f).Top + 60
    End Select
Next
End Sub
Private Sub Timer30_Timer() '即時移除打擊特效●New
If ppbb - pb <= 2 Then
    Timer30.Interval = 800
Else
    Timer30.Interval = 300
End If
If pb < ppbb Then
    '移除程序
     pa = 1 + pc * 8
     pb = 8 + pd * 8
        For f = pa To pb
            Unload Shape5(f)
        Next
     pc = pc + 1
     pd = pd + 1
     pda = 1 + pc * 8
End If
End Sub
Private Sub getstart() '打擊分數到一定程度可獲得星星和命的演算法●
If Val(Label16.Caption) <= 50000000 Then
    If nex >= 3 Then
        Label16.Caption = Val(Label16.Caption) + 1000
    Else
        Label16.Caption = Val(Label16.Caption) + 100
    End If
    If Val(Label16.Caption) Mod 10000 = 0 Then Call Starr
    If Val(Label16.Caption) Mod 8500 = 0 Then ad = ad + 1: Call multiha: Timer24.Enabled = True
    If Val(Label16.Caption) Mod 7000 = 0 Then
        If life(0) < 9 Then
            life(0) = life(0) + 1
            Label4.Caption = "X " & life(0)
        End If
        DoEvents
    End If
End If
End Sub
Private Sub winner() '贏或輸的名次指派呼叫●
mark = Val(Label16.Caption)
win = 1
Form4.Show
End Sub
Private Sub Timer31_Timer() '主背景移動●New 註：以下為饅頭的支授(採用桔梗第四代演算法v7.0)
'924表寬
'708表高
    dx = dx + 13
    If dx >= 1024 Then dx = 0 '====圖寬=====
    If dx <= 100 Then '=====大於564後就超圖了=====
        Form1.PaintPicture Form1.Picture, 0, 0, 924, 708, dx, 0, 924, 708, vbSrcCopy
    Else '上面form1截圖 複製到          0,0,表寬,表高,dx,0,表寬,表高
        Form1.PaintPicture Form1.Picture, 0, 0, 1024 - dx, 708, dx, 0, 1024 - dx, 708, vbSrcCopy
        '上面form1截圖 複製到           0,0,圖寬-dx,表高,dx,0,圖寬-dx,表高
        Form1.PaintPicture Form1.Picture, 1024 - dx, 0, dx - 100, 708, 0, 0, dx - 100, 708, vbSrcCopy
        '上面form1截圖 複製到           圖寬-dx,0,表寬+(dx-圖寬),表高,0,0,表寬+(dx+圖寬),表高
    End If
    Form1.Refresh
End Sub
