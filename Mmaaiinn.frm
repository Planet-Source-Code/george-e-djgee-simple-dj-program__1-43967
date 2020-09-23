VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DjGee"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8790
   Icon            =   "Mmaaiinn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   6
      Left            =   2520
      ScaleHeight     =   1905
      ScaleWidth      =   1905
      TabIndex        =   71
      Top             =   0
      Width           =   1935
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   75
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         Height          =   195
         Index           =   3
         Left            =   1455
         TabIndex        =   74
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   73
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         Height          =   195
         Index           =   5
         Left            =   1455
         TabIndex        =   72
         Top             =   1680
         Width           =   480
      End
      Begin VB.Shape Option1 
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   600
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   210
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   1080
         X2              =   1320
         Y1              =   960
         Y2              =   1080
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   12
         Left            =   960
         Shape           =   3  'Circle
         Top             =   840
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2520
      ScaleHeight     =   1905
      ScaleWidth      =   1905
      TabIndex        =   70
      Top             =   1920
      Width           =   1935
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   1
         X1              =   1440
         X2              =   1920
         Y1              =   1080
         Y2              =   960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   480
         Y1              =   960
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   2160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00808080&
         BorderWidth     =   3
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   3000
         Left            =   -1080
         Shape           =   3  'Circle
         Top             =   -495
         Width           =   3000
      End
   End
   Begin VB.CommandButton Cmd_All 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Speed Normal"
      Height          =   255
      Index           =   6
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_All 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mean All Speed"
      Height          =   255
      Index           =   5
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_All 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Volume Max"
      Height          =   255
      Index           =   4
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_All 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Volume = 0"
      Height          =   255
      Index           =   3
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_All 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mean All Volume"
      Height          =   255
      Index           =   2
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_All 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "All Position = 0"
      Height          =   255
      Index           =   1
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_All 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mean All Position"
      Height          =   255
      Index           =   0
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   5
      Left            =   -120
      ScaleHeight     =   4905
      ScaleMode       =   0  'User
      ScaleWidth      =   5985
      TabIndex        =   54
      Top             =   3840
      Width           =   6015
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   0
         Interval        =   500
         Left            =   840
         Top             =   720
      End
      Begin VB.CommandButton Cmd_play 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton Cmd_pause 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   ";"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Cmd_stop 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   720
         Width           =   1815
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   1
         Interval        =   500
         Left            =   1320
         Top             =   720
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   2
         Interval        =   500
         Left            =   4200
         Top             =   1920
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   3
         Interval        =   500
         Left            =   6360
         Top             =   3360
      End
      Begin VB.CommandButton Cmd_Open 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Cmd_Clear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear List"
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   3480
         Width           =   4215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mp3 Info"
         Height          =   2295
         Left            =   0
         TabIndex        =   55
         Top             =   1440
         Width           =   1815
         Begin VB.ListBox Lst_info 
            Height          =   1815
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   3495
         Left            =   1800
         TabIndex        =   62
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Path."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Count."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File Name."
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Time"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Played"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   4
      Left            =   5880
      ScaleHeight     =   4905
      ScaleMode       =   0  'User
      ScaleWidth      =   2865
      TabIndex        =   32
      Top             =   3840
      Width           =   2895
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "EQ"
         Height          =   3855
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   2895
         Begin VB.CheckBox Chk_CH_Right 
            BackColor       =   &H00E0E0E0&
            Caption         =   "On"
            Height          =   375
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox Chk_CH_Left 
            BackColor       =   &H00E0E0E0&
            Caption         =   "On"
            Height          =   375
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   720
            Width           =   495
         End
         Begin MSComctlLib.Slider Sli_CH_Left 
            Height          =   2175
            Left            =   1080
            TabIndex        =   36
            Top             =   1320
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   3836
            _Version        =   393216
            BorderStyle     =   1
            Orientation     =   1
            LargeChange     =   100
            Max             =   995
            SelStart        =   500
            TickStyle       =   3
            TickFrequency   =   100
            Value           =   500
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   2175
            Left            =   600
            TabIndex        =   37
            Top             =   1320
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   3836
            _Version        =   393216
            BorderStyle     =   1
            Orientation     =   1
            LargeChange     =   100
            Max             =   1500
            SelStart        =   500
            TickStyle       =   3
            TickFrequency   =   100
            Value           =   500
         End
         Begin MSComctlLib.Slider Sli_volume 
            Height          =   2175
            Left            =   120
            TabIndex        =   38
            Top             =   1320
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   3836
            _Version        =   393216
            BorderStyle     =   1
            Orientation     =   1
            LargeChange     =   100
            Max             =   995
            SelStart        =   500
            TickStyle       =   3
            TickFrequency   =   100
            Value           =   500
         End
         Begin MSComctlLib.Slider Sli_CH_Right 
            Height          =   2175
            Left            =   1560
            TabIndex        =   39
            Top             =   1320
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   3836
            _Version        =   393216
            BorderStyle     =   1
            Orientation     =   1
            LargeChange     =   100
            Max             =   995
            SelStart        =   500
            TickStyle       =   3
            TickFrequency   =   100
            Value           =   500
         End
         Begin MSComctlLib.Slider Sli_Bass 
            Height          =   2175
            Left            =   2040
            TabIndex        =   40
            Top             =   1320
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   3836
            _Version        =   393216
            BorderStyle     =   1
            Orientation     =   1
            LargeChange     =   100
            Max             =   995
            SelStart        =   500
            TickStyle       =   3
            TickFrequency   =   100
            Value           =   500
         End
         Begin MSComctlLib.Slider Sli_Treble 
            Height          =   2175
            Left            =   2520
            TabIndex        =   41
            Top             =   1320
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   3836
            _Version        =   393216
            BorderStyle     =   1
            Orientation     =   1
            LargeChange     =   100
            Max             =   995
            SelStart        =   500
            TickStyle       =   3
            TickFrequency   =   100
            Value           =   500
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Treble"
            ForeColor       =   &H80000008&
            Height          =   3615
            Index           =   6
            Left            =   2400
            TabIndex        =   53
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   47
            Top             =   3600
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   46
            Top             =   3600
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Channels"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   44
            Top             =   3600
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   43
            Top             =   3600
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Norm"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   480
            TabIndex        =   42
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Right"
            ForeColor       =   &H80000008&
            Height          =   3375
            Index           =   2
            Left            =   1440
            TabIndex        =   48
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bass"
            ForeColor       =   &H80000008&
            Height          =   3615
            Index           =   3
            Left            =   1920
            TabIndex        =   52
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Left"
            ForeColor       =   &H80000008&
            Height          =   3375
            Index           =   1
            Left            =   960
            TabIndex        =   49
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Spd"
            ForeColor       =   &H80000008&
            Height          =   3615
            Index           =   5
            Left            =   480
            TabIndex        =   50
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vol"
            ForeColor       =   &H80000008&
            Height          =   3615
            Index           =   4
            Left            =   0
            TabIndex        =   51
            Top             =   240
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   3
      Left            =   4440
      ScaleHeight     =   2442.883
      ScaleMode       =   0  'User
      ScaleWidth      =   2505
      TabIndex        =   24
      Top             =   1920
      Width           =   2535
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mp3 4"
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   3
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   2535
         Begin VB.CheckBox Chk_pick 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "4"
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   2295
         End
         Begin MSComctlLib.Slider Sli_progress 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            LargeChange     =   1000
            SmallChange     =   10
            TickStyle       =   3
         End
         Begin VB.Shape Shape1 
            Height          =   735
            Index           =   11
            Left            =   0
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label lblLengthA 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   31
            Top             =   960
            Width           =   1410
         End
         Begin VB.Label lblPositionA 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   30
            Top             =   1200
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Index           =   10
            Left            =   0
            Top             =   960
            Width           =   2535
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Index           =   9
            Left            =   0
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Lbl_No 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   29
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Lbl_No 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   28
            Top             =   1200
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   2
      Left            =   0
      ScaleHeight     =   2442.883
      ScaleMode       =   0  'User
      ScaleWidth      =   2505
      TabIndex        =   16
      Top             =   1920
      Width           =   2535
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mp3 3"
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   2
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2535
         Begin VB.CheckBox Chk_pick 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "3"
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   2295
         End
         Begin MSComctlLib.Slider Sli_progress 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            LargeChange     =   1000
            SmallChange     =   10
            TickStyle       =   3
         End
         Begin VB.Shape Shape1 
            Height          =   735
            Index           =   8
            Left            =   0
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label lblLengthA 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   23
            Top             =   960
            Width           =   1410
         End
         Begin VB.Label lblPositionA 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   22
            Top             =   1200
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Index           =   7
            Left            =   0
            Top             =   960
            Width           =   2535
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Index           =   6
            Left            =   0
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Lbl_No 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   21
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Lbl_No 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   20
            Top             =   1200
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   1
      Left            =   4440
      ScaleHeight     =   2442.883
      ScaleMode       =   0  'User
      ScaleWidth      =   2505
      TabIndex        =   8
      Top             =   0
      Width           =   2535
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mp3 2"
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2535
         Begin VB.CheckBox Chk_pick 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "2"
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2295
         End
         Begin MSComctlLib.Slider Sli_progress 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            LargeChange     =   1000
            SmallChange     =   10
            TickStyle       =   3
         End
         Begin VB.Shape Shape1 
            Height          =   735
            Index           =   5
            Left            =   0
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label lblLengthA 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   15
            Top             =   960
            Width           =   1410
         End
         Begin VB.Label lblPositionA 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   1200
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Index           =   4
            Left            =   0
            Top             =   960
            Width           =   2535
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Index           =   3
            Left            =   0
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Lbl_No 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Lbl_No 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   1200
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   0
      Left            =   0
      ScaleHeight     =   2442.883
      ScaleMode       =   0  'User
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mp3 1"
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2535
         Begin VB.CheckBox Chk_pick 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "1"
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   555
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   2295
         End
         Begin MSComctlLib.Slider Sli_progress 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   1440
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            LargeChange     =   1000
            SmallChange     =   10
            TickStyle       =   3
         End
         Begin VB.Label Lbl_No 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   7
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Lbl_No 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Length"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   6
            Top             =   960
            Width           =   735
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Index           =   2
            Left            =   0
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Shape Shape1 
            Height          =   255
            Index           =   1
            Left            =   0
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lblPositionA 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   5
            Top             =   1200
            Width           =   1410
         End
         Begin VB.Label lblLengthA 
            BackColor       =   &H00400000&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   4
            Top             =   960
            Width           =   1410
         End
         Begin VB.Shape Shape1 
            Height          =   735
            Index           =   0
            Left            =   0
            Top             =   240
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------------------
'' you know, i cant really copywrite this code because _
    '' hell, anyone could write it, its just not too complicated _
    '' so have fun with it, and i hope you make a better dj _
    '' program than i did, if you do then send the vb source to _
    '' ggolemg@yahoo.com
'' btw:: i hate skins, so if you just skin this, dont _
    '' bother sending it and this project took 5 ciggs in appeasement _
    '' to my computer
'' thanks george
''-------------------------------------------------------------
Option Explicit
Private blah As Long
Private v1 As Long 'vertical scratch above mid
Private v2 As Long 'vertical scratch below mid
Private traceLeftRight As Double 'quad fader horizontal
Private traceTopBottom As Double 'quad fader vertical
Private sx As Long 'quad fader original pos x
Private sy As Long 'quad fader original pos y

Private Sub Chk_CH_Left_Click() 'left channel on/off

    With Chk_CH_Left
        If .Value = 1 Then
            .Caption = "OFF"
            setMP3channelState wichMp3, "left", "off" 'set left channel off
          Else 'NOT .VALUE...
            .Caption = "ON"
            setMP3channelState wichMp3, "left", "on" 'set left channel on
            setMp3channelVolume wichMp3, "left", Sli_CH_Left.Value 'show left channel volume
        End If
    End With 'CHK_CH_LEFT

End Sub

Private Sub Chk_CH_Right_Click() 'right channel on/off

    With Chk_CH_Right
        If .Value = 1 Then
            .Caption = "OFF"
            setMP3channelState wichMp3, "right", "off" 'set right channel off
          Else 'NOT .VALUE...
            .Caption = "ON"
            setMP3channelState wichMp3, "right", "on" 'set right channel on
            setMp3channelVolume wichMp3, "right", Sli_CH_Left.Value 'show right channel volume

        End If
    End With 'CHK_CH_RIGHT

End Sub

Private Sub Chk_pick_Click(Index As Integer) 'pick wichmp3 you are currently on

    wichMp3 = Index 'if i keep everything within the context of 'index' then min/max wichmp3 is never < > index
    Sli_volume.Value = statusVolume(Index) 'volume
    Slider1.Value = statusSpeed(Index) 'speed
    Sli_Treble.Value = statusTreble(Index) 'treble
    Sli_Bass.Value = statusBass(Index) 'bass

    If Index = 0 Then 'wichmp3 = 0
        Chk_pick(1).Value = 0
        Chk_pick(2).Value = 0
        Chk_pick(3).Value = 0
      ElseIf Index = 1 Then 'NOT INDEX... and wichmp3 = 1
        Chk_pick(0).Value = 0
        Chk_pick(2).Value = 0
        Chk_pick(3).Value = 0
      ElseIf Index = 2 Then 'NOT INDEX... and wichmp3 = 2
        Chk_pick(0).Value = 0
        Chk_pick(1).Value = 0
        Chk_pick(3).Value = 0
      ElseIf Index = 3 Then 'NOT INDEX... and wichmp3 = 3
        Chk_pick(0).Value = 0
        Chk_pick(1).Value = 0
        Chk_pick(2).Value = 0
    End If

    If Chk_pick(Index).Value = 0 Then
        Chk_pick(Index).BackColor = &HC0FFC0
      Else 'NOT CHK_PICK(INDEX).VALUE...
        Chk_pick(Index).BackColor = vbGreen
    End If

End Sub

Private Sub Cmd_All_Click(Index As Integer)

  Dim i As Integer 'just used in for/next loops
  Dim P(0 To 3) As Long 'easier than p1,p2,p3

    If Index = 0 Then 'mean all position
        For i = 0 To 3
            P(i) = statusPosition(i)
        Next i
        For i = 0 To 3
            setMp3CurrentTime i, ((P(0) + P(1) + P(2) + P(3)) \ 4), "repeat"
        Next i

      ElseIf Index = 1 Then 'NOT INDEX... all position = 0
        For i = 0 To 3
            setMp3CurrentTime i, 0, "repeat"
        Next i

      ElseIf Index = 2 Then 'NOT INDEX... mean all volume
        For i = 0 To 3
            P(i) = statusVolume(i)
        Next i
        For i = 0 To 3
            setMp3Volume i, ((P(0) + P(1) + P(2) + P(3)) \ 4) 'set mean volume
        Next i
        Sli_volume.Value = statusVolume(wichMp3) 'display mean volume

      ElseIf Index = 3 Then 'NOT INDEX... all volume = 0
        For i = 0 To 3
            setMp3Volume i, 0 '0 is 0 :)
        Next i
        Sli_volume.Value = statusVolume(wichMp3) 'display volume

      ElseIf Index = 4 Then 'NOT INDEX... all volume max
        For i = 0 To 3
            setMp3Volume i, 1000 '1000 is max value
        Next i
        Sli_volume.Value = statusVolume(wichMp3) 'display volume

      ElseIf Index = 5 Then 'NOT INDEX... mean all speed
        For i = 0 To 3
            P(i) = statusSpeed(i)
        Next i
        For i = 0 To 3
            setMp3Speed i, ((P(0) + P(1) + P(2) + P(3)) \ 4) 'set mean speed
        Next i

      ElseIf Index = 6 Then 'NOT INDEX... all speed normal
        For i = 0 To 3
            setMp3Speed i, 1000 '1000 is normal value
        Next i

    End If

End Sub

Private Sub Cmd_Clear_Click() 'clear list window

    lstFiles.ListItems.Clear

End Sub

Private Sub Cmd_Open_Click() 'open tree view

    Me.MousePointer = 11 'hourglass
    frmExplore.Show (1) 'show it
    Me.MousePointer = 0 'default

End Sub

Private Sub Cmd_pause_Click()

    If lstFiles.ListItems.Count > 0 Then 'make sure theres something there to pause
        PauseMp3 wichMp3 'pause selected
        Timer1(wichMp3).Enabled = False 'disable according timer
        Frame1(wichMp3).BackColor = vbWhite 'let user know
    End If

End Sub

Private Sub Cmd_play_Click(Index As Integer)

  Dim lShortPath As Long
  Dim sShortPath As String * 260
  Dim sShortPathName As String

    If lstFiles.ListItems.Count > 0 Then 'make sure theres something there to play
        Me.MousePointer = 11 'hourglass
        Chk_pick(wichMp3).Value = 1 'set picked

        With lstFiles
            lShortPath = GetShortPathName(.ListItems(.SelectedItem.Index).Text & .ListItems(.SelectedItem.Index).ListSubItems(2).Text, sShortPath, 260) 'get path to mp3
            sShortPathName = ftnStripNullChar(sShortPath) 'get wrid of null character
        End With 'LSTFILES
        openMp3 sShortPathName, wichMp3 'open mp3
        setMp3State "on", wichMp3 'set this mp3 on
        setMp3TimeFormat wichMp3, "tmsf" 'set this mp3s time format
        PlayMp3 wichMp3, "repeat" 'play this mp3
        lblLengthA(wichMp3).Caption = statusLength(wichMp3) 'display length of it
        Sli_progress(wichMp3).Max = statusLength(wichMp3) 'display length of it
        Sli_Bass.Value = statusBass(wichMp3) 'display bass
        Sli_Treble.Value = statusTreble(wichMp3) 'display treble
        Sli_CH_Right.Value = statusChannelVolume(wichMp3, "right") 'display right channel volume
        Sli_CH_Left.Value = statusChannelVolume(wichMp3, "left") 'display left channel volume

        With lstFiles
            Chk_pick(wichMp3).Caption = "(" & Mid$(.ListItems(.SelectedItem.Index).ListSubItems(1).Text, 1, 2) & ") " & .ListItems(.SelectedItem.Index).ListSubItems(2).Text 'display name in checkbox
            .ListItems(.SelectedItem.Index).ListSubItems(3).Text = "P" 'already selected
        End With 'LSTFILES
        Timer1(wichMp3).Enabled = True 'enable according timer
        Me.MousePointer = 0 'default
    End If
    Chk_pick_Click (wichMp3)

End Sub

Private Sub Cmd_stop_Click()

    Timer1(wichMp3).Enabled = False 'disable selected mp3
    StopMp3 wichMp3 'stop selected mp3
    CloseMp3 wichMp3 'close selected mp3
    lblLengthA(wichMp3).Caption = "0" 'display nothing
    lblPositionA(wichMp3).Caption = "0" 'display nothing
    Chk_pick(wichMp3).Caption = "No Track loaded." 'display nothing
    Frame1(wichMp3).BackColor = vbWhite 'let user know not on

End Sub

Private Sub DisAllMCI(Index As Integer) 'display all for timer

    lblPositionA(Index).Caption = statusPosition(Index) 'display position
    Sli_progress(Index).Value = statusPosition(Index) 'display position

    If statusMp3state(Index) = "on" Then
        Frame1(Index).BackColor = &HC000& 'let user know
      Else 'NOT STATUSMP3STATE(INDEX)...
        Frame1(Index).BackColor = &H80FF80 'let user know
    End If
    'mciSendString "status mp3" & Index & " play speed", sReturnBuffer4$, Len(sReturnBuffer4$), 0
    'Lbl_PS(Index).Caption = Val(sReturnBuffer4$)
    'mciSendString "status mp3" & Index & " audio breaks", sReturnBuffer$, Len(sReturnBuffer$), 0
    'Lbl_breaks(Index).Caption = Val(sReturnBuffer$)

End Sub

Private Sub Form_Load()

  Dim i As Integer

    InitCommonControls 'win xp
    Option1.Left = (Picture1(6).Width \ 2) - (Option1.Width \ 2) 'positioning
    Option1.Top = (Picture1(6).Height \ 2) - (Option1.Height \ 2) 'positioning
    Line2.X1 = Picture1(6).Width \ 2 'positioning
    Line2.X2 = Picture1(6).Width \ 2 'positioning
    Line2.Y1 = Picture1(6).Height \ 2 'positioning
    Line2.Y2 = Picture1(6).Height \ 2 'positioning
    Shape1(12).Left = (Picture1(6).Width \ 2) - (Shape1(12).Width \ 2) 'positioning
    Shape1(12).Top = (Picture1(6).Height \ 2) - (Shape1(12).Height \ 2) 'positioning
    For i = 0 To 1
        With Line3(i)
            .X1 = 0 'positioning
            .Y1 = 0 'positioning
            .X2 = 0 'positioning
            .Y2 = 0 'positioning
        End With 'LINE3(I)
    Next i

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim i As Integer

    For i = 0 To 3
        wichMp3 = i
        Call Cmd_stop_Click 'stop all
    Next i

End Sub

Private Sub GetMP3Inf()

  Dim accMP3Info As MP3Info

    getMP3Info MP3FileName, accMP3Info
    With Lst_info
        .Clear
        .AddItem "Size: " & accMP3Info.SIZE
        .AddItem "Length: " & accMP3Info.LENGTH
        .AddItem accMP3Info.MPEG & " " & accMP3Info.LAYER
        .AddItem accMP3Info.BITRATE
        .AddItem accMP3Info.FREQ & " " & accMP3Info.CHANNELS
        .AddItem "CRC: " & accMP3Info.CRC
        .AddItem "Copy: " & accMP3Info.COPYRIGHT
        .AddItem "Emphasis: " & accMP3Info.EMPHASIS
        .AddItem "Original: " & accMP3Info.ORIGINAL
    End With 'LST_INFO
    Me.MousePointer = 0

End Sub

Private Sub Label1_Click()

    setMp3Speed wichMp3, 1000 'set normal speed
    Slider1.Value = 1000 'display speed
    Timer1(wichMp3).Enabled = True 'enable timer

End Sub

Private Sub lstFiles_Click()

  Dim lShortPath As Long
  Dim sShortPath As String * 260
  Dim sShortPathName As String

    If lstFiles.ListItems.Count > 0 Then
        Me.MousePointer = 11

        With lstFiles
            lShortPath = GetShortPathName(.ListItems(.SelectedItem.Index).Text & .ListItems(.SelectedItem.Index).ListSubItems(2).Text, sShortPath, 260)
            sShortPathName = ftnStripNullChar(sShortPath)
        End With 'LSTFILES
        MP3FileName = sShortPathName
        GetMP3Inf 'get mp3 info

    End If

End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

    Picture1(6).Cls 'clear picture
    sx = Picture1(6).Width \ 2 'set old pos
    sy = Picture1(6).Height \ 2 'set old pos

End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

  Dim r As Integer, g As Integer, b As Integer, val As Integer

    If Index = 6 Then 'quad fader/mixer blah :)
        If x > 0 And x < Picture1(6).Width And Y > 0 And Y < Picture1(6).Height Then 'stay inside
            If Button = 1 Then 'user hold the left mouse button down
                val = 7 'trial and error haha
                Option1.Left = x - (Option1.Width \ 2) 'set red circles left
                Line2.X1 = Option1.Left + (Option1.Width \ 2)
                Option1.Top = Y - (Option1.Height \ 2) 'set red circles top
                Line2.Y1 = Option1.Top + (Option1.Height \ 2)
                r = (Option1.Top \ val) 'red
                g = (Option1.Left \ val) 'green
                b = (((Option1.Left \ val) + (Option1.Top \ val)) \ 2) 'blue
                If r < 255 And g < 255 And b < 255 And r > 0 And g > 0 And b > 0 Then 'so no errors
                    Picture1(6).Line (sx, sy)-(x, Y), RGB(r, g, b) 'draw all the lines
                End If
                traceLeftRight = (Option1.Left + 90) \ 2 'set variables
                traceTopBottom = (Option1.Top + 90) \ 2 'set variables

                setMp3Volume 0, ((((-traceLeftRight) - traceTopBottom)) \ 2) + 1000
                Label4(2).Caption = (((((-traceLeftRight) - traceTopBottom)) \ 2) + 1000)  ' correct

                setMp3Volume 1, ((traceLeftRight - traceTopBottom) + 1000) \ 2
                Label4(3).Caption = (((traceLeftRight - traceTopBottom) + 1000) \ 2)  'correct

                setMp3Volume 2, ((-traceLeftRight + traceTopBottom) + 1000) \ 2
                Label4(4).Caption = (((-traceLeftRight + traceTopBottom) + 1000) \ 2)  'correct

                setMp3Volume 3, ((traceLeftRight + traceTopBottom) \ 2)
                Label4(5).Caption = (((traceLeftRight + traceTopBottom) \ 2))  'correct

                Sli_volume.Value = statusVolume(wichMp3) 'show volume
            End If
        End If
    End If

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    With Line1
        .X1 = 0 'positioning
        .X2 = Picture2.Width 'positioning
        .Y1 = Picture2.Height \ 2 'positioning
        .Y2 = Picture2.Height \ 2 'positioning
    End With 'LINE1

    If Button = 1 Then 'user has to hold down button 1
        If Y < Picture2.Height \ 2 Then 'top half
            setMp3CurrentTime wichMp3, statusPosition(wichMp3) + v1, "repeat" 'ahead
            v1 = v1 + 500
            v2 = 1
            With Line3(0)
                .X1 = Line1.X1
                .Y1 = Line1.Y1
                .X2 = x
                .Y2 = Y
            End With 'LINE3(0)
            With Line3(1)
                .X2 = Line1.X2
                .Y2 = Line1.Y2
                .X1 = x
                .Y1 = Y
            End With 'LINE3(1)

          Else 'NOT Y... bottom half
            setMp3CurrentTime wichMp3, statusPosition(wichMp3) - v2, "repeat" 'behind
            v2 = v2 + 500
            v1 = 1
            With Line3(0)
                .X1 = Line1.X1
                .Y1 = Line1.Y1
                .X2 = x
                .Y2 = Y
            End With 'LINE3(0)
            With Line3(1)
                .X2 = Line1.X2
                .Y2 = Line1.Y2
                .X1 = x
                .Y1 = Y
            End With 'LINE3(1)

        End If

      Else 'NOT BUTTON...
        v1 = 0
        v2 = 0
    End If

End Sub


Private Sub Sli_Bass_Scroll()

    setMp3Bass wichMp3, Sli_Bass.Value 'sets bass

End Sub

Private Sub Sli_CH_Left_Scroll() 'slider left channel scroll

    setMp3channelVolume wichMp3, "left", Sli_CH_Left.Value 'set the left channel volume
    Sli_volume.Value = statusVolume(wichMp3) 'display volume
    Sli_CH_Right.Value = statusChannelVolume(wichMp3, "right") 'display left channel volume
    Label2(5).Caption = Sli_volume.Value 'display volume
    Label2(2).Caption = Sli_CH_Left.Value 'display left channel volume
    Label2(1).Caption = Sli_CH_Right.Value 'display right channel volume
    Label2(6).Caption = Slider1.Value 'display speed

End Sub

Private Sub Sli_CH_Right_Scroll() 'slider right channel scroll

    setMp3channelVolume wichMp3, "right", Sli_CH_Right.Value 'set the right channel volume
    Sli_volume.Value = statusVolume(wichMp3) 'display volume
    Sli_CH_Left.Value = statusChannelVolume(wichMp3, "left") 'display left channel volume
    Label2(5).Caption = Sli_volume.Value 'display volume
    Label2(2).Caption = Sli_CH_Left.Value 'display left channel volume
    Label2(1).Caption = Sli_CH_Right.Value 'display right channel volume
    Label2(6).Caption = Slider1.Value 'display speed

End Sub

Private Sub Sli_progress_Scroll(Index As Integer) 'change progress of mp3

    setMp3CurrentTime Index, Sli_progress(Index).Value, "repeat" 'set current time to slider
    lstFiles.SetFocus

End Sub

Private Sub Sli_Treble_Scroll() 'slider treble scroll

    setMp3Treble wichMp3, Sli_Treble.Value 'set treble

End Sub

Private Sub Sli_volume_Scroll()

    setMp3Volume wichMp3, Sli_volume.Value 'set volume
    Sli_CH_Left.Value = statusChannelVolume(wichMp3, "left") 'display left channel volume
    Sli_CH_Right.Value = statusChannelVolume(wichMp3, "right") 'display right channel volume
    Label2(5).Caption = Sli_volume.Value 'display volume
    Label2(2).Caption = Sli_CH_Left.Value 'display left channel volume
    Label2(1).Caption = Sli_CH_Right.Value 'display right channel volume
    'Sli_CH_Left.Value = Sli_volume.Value
    'Sli_CH_Right.Value = Sli_volume.Value
    Label2(6).Caption = Slider1.Value 'display speed

End Sub

Private Sub Slider1_Scroll() 'set speed

    setMp3Speed wichMp3, Slider1.Value 'set speed
    Label2(5).Caption = Sli_volume.Value 'display volume
    Label2(2).Caption = Sli_CH_Left.Value 'display left channel volume
    Label2(1).Caption = Sli_CH_Right.Value 'display right channel volume
    Label2(6).Caption = Slider1.Value 'display speed

End Sub

Private Sub Timer1_Timer(Index As Integer)

    DisAllMCI Index 'call to disallmci

End Sub



