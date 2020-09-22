VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BAT-Man Batch file Maker                         - By 'GeekFreek' Anil"
   ClientHeight    =   7980
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9990
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7940.297
   ScaleMode       =   0  'User
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Control Statments"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Index           =   4
      Left            =   8145
      TabIndex        =   82
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox ctrlPrompt 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   135
         TabIndex        =   92
         Text            =   "Your Prompt"
         Top             =   585
         Width           =   1920
      End
      Begin VB.TextBox Ctrl3 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   1560
         TabIndex        =   88
         Top             =   1350
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Ctrl2 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   1560
         TabIndex        =   87
         Top             =   1095
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Ctrl1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   1560
         TabIndex        =   86
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cboCtrl 
         Height          =   315
         ItemData        =   "Main.frx":0D7A
         Left            =   120
         List            =   "Main.frx":0D90
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   240
         Width           =   1935
      End
      Begin Project1.GoldButton ctrlOK 
         Height          =   285
         Left            =   360
         TabIndex        =   83
         Tag             =   "Contains commands used to manipulate Files"
         Top             =   1575
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Caption         =   "OK"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton ctrlCancel 
         Height          =   375
         Left            =   2160
         TabIndex        =   84
         Top             =   180
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin VB.Label lbl3 
         Caption         =   "---"
         Height          =   240
         Left            =   135
         TabIndex        =   91
         Top             =   1320
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label lbl2 
         Caption         =   "---"
         Height          =   240
         Left            =   135
         TabIndex        =   90
         Top             =   1080
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label lbl1 
         Caption         =   "---"
         Height          =   240
         Left            =   135
         TabIndex        =   89
         Top             =   870
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Frame FrmOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Line"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Index           =   3
      Left            =   7920
      TabIndex        =   73
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtCustLine 
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   77
         Text            =   "29"
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         ItemData        =   "Main.frx":0DC8
         Left            =   240
         List            =   "Main.frx":0DE1
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   720
         Width           =   2175
      End
      Begin Project1.GoldButton LineOK 
         Height          =   285
         Left            =   840
         TabIndex        =   74
         Tag             =   "Contains commands used to manipulate Files"
         Top             =   1590
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         Caption         =   "OK"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton LineCancel 
         Height          =   375
         Left            =   2160
         TabIndex        =   75
         Top             =   180
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin VB.Label Label9 
         Caption         =   "Select the style of line to draw :"
         Height          =   375
         Left            =   360
         TabIndex        =   78
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmMsgBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Message box"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   120
      TabIndex        =   52
      Top             =   6000
      Width           =   9735
      Begin VB.TextBox txtLineMSG 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   69
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox lstMSG 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   3840
         TabIndex        =   68
         ToolTipText     =   "Double-Click an item to edit it"
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtBoxWidth 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   66
         Text            =   "79"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   63
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Frame FrmCustSet 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Enter Ascii Values"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         Begin VB.TextBox CustSet 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1170
            MaxLength       =   3
            TabIndex        =   58
            ToolTipText     =   "Enter the ASCII value of the character you want in the Middle-Right Of the Box"
            Top             =   555
            Width           =   495
         End
         Begin VB.TextBox CustSet 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   645
            MaxLength       =   3
            TabIndex        =   60
            ToolTipText     =   "Enter the ASCII value of the character you want in the Middle-Bottom Of the Box"
            Top             =   870
            Width           =   495
         End
         Begin VB.TextBox CustSet 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   120
            MaxLength       =   3
            TabIndex        =   62
            ToolTipText     =   "Enter the ASCII value of the character you want in the Middle-Left of the box"
            Top             =   550
            Width           =   495
         End
         Begin VB.TextBox CustSet 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   650
            MaxLength       =   3
            TabIndex        =   56
            ToolTipText     =   "Enter the ASCII value of the character you want in the Middle-Top of the box"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox CustSet 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1170
            MaxLength       =   3
            TabIndex        =   59
            ToolTipText     =   "Enter the ASCII value of the character you want in the Botton-Right of the Box"
            Top             =   870
            Width           =   495
         End
         Begin VB.TextBox CustSet 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   120
            MaxLength       =   3
            TabIndex        =   61
            ToolTipText     =   "Enter the ASCII value of the character you want in the Bottom-Left Of the Box"
            Top             =   870
            Width           =   495
         End
         Begin VB.TextBox CustSet 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1170
            MaxLength       =   3
            TabIndex        =   57
            ToolTipText     =   "Enter the ASCII value of the character you want in the Top-Right Of the Box"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox CustSet 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   120
            MaxLength       =   3
            TabIndex        =   55
            ToolTipText     =   "Enter the ASCII value of the character you want in the Top-Left Of the Box"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ComboBox SelectStyle 
         Height          =   315
         ItemData        =   "Main.frx":0E3A
         Left            =   600
         List            =   "Main.frx":0E4A
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   240
         Width           =   1335
      End
      Begin Project1.GoldButton MsgDown 
         Height          =   285
         Left            =   8400
         TabIndex        =   71
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   16576
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":0E7C
      End
      Begin Project1.GoldButton MsgCancel 
         Height          =   405
         Left            =   9240
         TabIndex        =   72
         ToolTipText     =   "Remove Current Line"
         Top             =   135
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   714
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton MsgOK 
         Height          =   405
         Left            =   9000
         TabIndex        =   79
         Tag             =   "Contains commands used to manipulate Files"
         Top             =   1440
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   714
         Caption         =   "OK"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton msgTest 
         Height          =   285
         Left            =   9000
         TabIndex        =   80
         Tag             =   "Contains commands used to manipulate Files"
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   503
         Caption         =   "Test"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton MsgHelp 
         Height          =   255
         Left            =   8880
         TabIndex        =   94
         Top             =   600
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   450
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   16576
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Style :"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   280
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   2
         X1              =   3720
         X2              =   2160
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   1
         X1              =   3720
         X2              =   3720
         Y1              =   120
         Y2              =   1920
      End
      Begin VB.Label Label7 
         Caption         =   "Box Properties :"
         Height          =   255
         Left            =   2280
         TabIndex        =   67
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Box Width (15-79)"
         Height          =   255
         Left            =   2280
         TabIndex        =   65
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Box Title :"
         Height          =   255
         Left            =   2280
         TabIndex        =   64
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   2160
         X2              =   2160
         Y1              =   120
         Y2              =   1920
      End
   End
   Begin VB.Frame FrmOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "DIR Command"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Index           =   2
      Left            =   7680
      TabIndex        =   43
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CheckBox WideView 
         Caption         =   "Wide List Format"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox PagePause 
         Caption         =   "Pause on each page"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox lookInSub 
         Caption         =   "Look in Sub-Directories"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Wild 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   47
         Text            =   "< Criteria >"
         Top             =   525
         Width           =   1815
      End
      Begin VB.CheckBox Criteria 
         Caption         =   "Criteria (can use *,?) :"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   1815
      End
      Begin Project1.GoldButton DirOk 
         Height          =   285
         Left            =   840
         TabIndex        =   44
         Tag             =   "Contains commands used to manipulate Files"
         Top             =   1590
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         Caption         =   "OK"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton DirCancel 
         Height          =   375
         Left            =   2160
         TabIndex        =   45
         Top             =   180
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
   End
   Begin VB.Frame FrmOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Index           =   1
      Left            =   7440
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtSrc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtDes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   2415
      End
      Begin Project1.GoldButton Apply 
         Height          =   315
         Left            =   840
         TabIndex        =   37
         Tag             =   "Contains commands used to manipulate Files"
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "OK"
         Enabled         =   0   'False
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton CDSource 
         Height          =   255
         Left            =   1710
         TabIndex        =   38
         Top             =   315
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton CDDes 
         Height          =   255
         Left            =   1710
         TabIndex        =   39
         Top             =   915
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton SDFCancel 
         Height          =   375
         Left            =   2160
         TabIndex        =   40
         Top             =   180
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin VB.Label Label1 
         Caption         =   "From :"
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "To :"
         Height          =   255
         Left            =   585
         TabIndex        =   41
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame FrmOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Attribute Options"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Index           =   0
      Left            =   7200
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin Project1.GoldButton AttrGo 
         Height          =   615
         Left            =   2040
         TabIndex        =   30
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1085
         Caption         =   "Go !"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin VB.Frame F 
         Appearance      =   0  'Flat
         Caption         =   "S"
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   1560
         TabIndex        =   26
         Top             =   960
         Width           =   375
         Begin VB.OptionButton Y 
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   195
         End
         Begin VB.OptionButton N 
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   195
         End
      End
      Begin VB.Frame F 
         Appearance      =   0  'Flat
         Caption         =   "R"
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   1200
         TabIndex        =   23
         Top             =   960
         Width           =   375
         Begin VB.OptionButton Y 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   195
         End
         Begin VB.OptionButton N 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   195
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "H"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   375
         Begin VB.OptionButton Y 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   195
         End
         Begin VB.OptionButton N 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   195
         End
      End
      Begin VB.Frame F 
         Appearance      =   0  'Flat
         Caption         =   "A"
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   0
         Left            =   480
         TabIndex        =   17
         Top             =   960
         Width           =   375
         Begin VB.OptionButton N 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   195
         End
         Begin VB.OptionButton Y 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.TextBox Fn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2415
      End
      Begin Project1.GoldButton AttrCancel 
         Height          =   375
         Left            =   2160
         TabIndex        =   31
         Top             =   165
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton CDButton 
         Height          =   255
         Left            =   1680
         TabIndex        =   33
         Top             =   315
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":0F5E
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Y N"
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   1200
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "File Name :"
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Frame CMDFrm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Commands"
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   2655
      Begin Project1.GoldButton StaticButton 
         Height          =   525
         Left            =   240
         TabIndex        =   7
         Tag             =   "Contains commonly used DOS commands"
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   926
         Caption         =   "Static Commands"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         Picture         =   "Main.frx":14F8
      End
      Begin Project1.GoldButton DirCmds 
         Height          =   525
         Left            =   240
         TabIndex        =   8
         Tag             =   "Contains commands used to manipulate directories"
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   926
         Caption         =   "Folder Commands"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         Picture         =   "Main.frx":1652
      End
      Begin Project1.GoldButton FilCmds 
         Height          =   525
         Left            =   240
         TabIndex        =   9
         Tag             =   "Contains commands used to manipulate Files"
         Top             =   1680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   926
         Caption         =   "  File Commands  "
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         Picture         =   "Main.frx":17AC
      End
      Begin Project1.GoldButton AG 
         Height          =   525
         Left            =   240
         TabIndex        =   10
         Tag             =   "Contains graphical stuff you can add to the batch file"
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   926
         Caption         =   "ASCII Graphics :)"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         Picture         =   "Main.frx":188E
      End
      Begin Project1.GoldButton Advanced 
         Height          =   525
         Left            =   240
         TabIndex        =   81
         Tag             =   "Contains graphical stuff you can add to the batch file"
         Top             =   2880
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   926
         Caption         =   "Advanced Stuff"
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnDown          =   2
         Picture         =   "Main.frx":19E8
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Batch File Code:"
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin MSComDlg.CommonDialog CD2 
         Left            =   4920
         Top             =   4080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Batch File Name"
         Filter          =   "Batch Files (*.bat)|*.bat|"
         InitDir         =   "C:\"
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   4440
         Top             =   4080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select File"
         Filter          =   "All Files (*.*)|*.*|"
         InitDir         =   "C:\"
      End
      Begin VB.ListBox LB 
         Appearance      =   0  'Flat
         Height          =   4320
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   360
         Width           =   6375
      End
      Begin Project1.GoldButton Up 
         Height          =   405
         Left            =   1200
         TabIndex        =   2
         Top             =   4800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   714
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   16576
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":1B42
      End
      Begin Project1.GoldButton Down 
         Height          =   405
         Left            =   1680
         TabIndex        =   3
         Top             =   4800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   714
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   16576
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":1C24
      End
      Begin Project1.GoldButton DelCurLine 
         Height          =   405
         Left            =   2160
         TabIndex        =   4
         ToolTipText     =   "Remove Current Line"
         Top             =   4800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   714
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":1D06
      End
      Begin Project1.GoldButton EditManually 
         Height          =   405
         Left            =   3960
         TabIndex        =   5
         ToolTipText     =   "Remove Current Line"
         Top             =   4800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Caption         =   "Edit"
         Alignment       =   2
         HoverColor      =   255
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton Test 
         Height          =   405
         Left            =   5640
         TabIndex        =   11
         ToolTipText     =   "Remove Current Line"
         Top             =   4800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":22A0
      End
      Begin Project1.GoldButton Make 
         Height          =   405
         Left            =   5640
         TabIndex        =   12
         ToolTipText     =   "Remove Current Line"
         Top             =   5280
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":283A
      End
      Begin Project1.GoldButton Insert 
         Height          =   405
         Left            =   2880
         TabIndex        =   13
         ToolTipText     =   "Remove Current Line"
         Top             =   4800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   714
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":2994
      End
      Begin Project1.GoldButton CDbatch 
         Height          =   255
         Left            =   4560
         TabIndex        =   51
         Top             =   5400
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   450
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
      End
      Begin Project1.GoldButton ListHelp 
         Height          =   255
         Left            =   6480
         TabIndex        =   93
         Top             =   120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   450
         Caption         =   ""
         Alignment       =   2
         ForeColor       =   -2147483630
         SkinDisabledText=   -2147483632
         SkinHighlight   =   -2147483628
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OnUp            =   0
         OnDown          =   2
         Picture         =   "Main.frx":2A76
      End
      Begin VB.Label lblSaveFile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   32
         Top             =   5400
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   240
         Top             =   4920
         Width           =   855
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnudfsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "&Test"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuMake 
         Caption         =   "Ma&ke"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnusdh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Batch file.."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "&Insert Batch File.."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuasdfd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuClip 
      Caption         =   "C&lip"
      Visible         =   0   'False
      Begin VB.Menu mnuClipCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuCmds 
      Caption         =   "&Commands"
      Begin VB.Menu mnuStaticCmds 
         Caption         =   "Static Commands"
         Tag             =   "Contains commonly used DOS commands"
         Begin VB.Menu mnuc1 
            Caption         =   "Cancel"
         End
         Begin VB.Menu mnudjfh 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEchoOff 
            Caption         =   "Echo Off"
            Shortcut        =   {F1}
            Tag             =   "This turns off the 'echoing' of commands by DOS"
         End
         Begin VB.Menu mnuEchoOn 
            Caption         =   "Echo On"
            Shortcut        =   {F2}
            Tag             =   "This turns on the 'echoing' of commands by DOS"
         End
         Begin VB.Menu mnudf 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCls 
            Caption         =   "Cls (Clear Screen)"
            Shortcut        =   {F3}
            Tag             =   "Insert a Clear Screen commnad"
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Pause"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnujf 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDate 
            Caption         =   "Date (View only)"
            Shortcut        =   {F5}
            Tag             =   "echo. | date | find ""Cu"""
         End
         Begin VB.Menu mnuDateset 
            Caption         =   "Date (View/Set)"
         End
         Begin VB.Menu mnuTime 
            Caption         =   "Time (View only)"
            Shortcut        =   {F6}
            Tag             =   "echo. | time | find ""Cu"""
         End
         Begin VB.Menu mnuTimeset 
            Caption         =   "Time (View/Set)"
         End
         Begin VB.Menu mnuTree 
            Caption         =   "Tree"
            Shortcut        =   {F7}
            Tag             =   "Displays Tree structure of current directory"
         End
         Begin VB.Menu mnudfhg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuWin 
            Caption         =   "Win (Start Windows)"
            Shortcut        =   {F8}
            Tag             =   "Starts Windows"
         End
         Begin VB.Menu mnuShutDown 
            Caption         =   "Shut Down"
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnuRestart 
            Caption         =   "Restart"
            Shortcut        =   {F11}
            Tag             =   "Restarts the computer"
         End
         Begin VB.Menu mnuVer 
            Caption         =   "Ver (Shows OS version)"
         End
         Begin VB.Menu mnuSelfDel 
            Caption         =   "Self-Delete command"
         End
         Begin VB.Menu mnusdfsdas 
            Caption         =   "-"
         End
         Begin VB.Menu mnuComment 
            Caption         =   "Comment"
            Shortcut        =   {F12}
         End
      End
      Begin VB.Menu mnuDirCmds 
         Caption         =   "Folder Commands"
         Tag             =   "Contains commands used to manipulate directories"
         Begin VB.Menu mnuC2 
            Caption         =   "Cancel"
         End
         Begin VB.Menu dfg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMD 
            Caption         =   "MD (Make Directory)"
            Shortcut        =   ^{F1}
            Tag             =   "Creates a directory"
         End
         Begin VB.Menu mnuCD 
            Caption         =   "CD (Change Directory)"
            Shortcut        =   ^{F2}
            Tag             =   "Insert a Change Directory Command"
         End
         Begin VB.Menu mnuRD 
            Caption         =   "RD (Remove Directory)"
            Shortcut        =   ^{F3}
            Tag             =   "Removes a directory"
         End
         Begin VB.Menu mnuRenDir 
            Caption         =   "Ren (Rename Directory)"
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu bh 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDir 
            Caption         =   "DIR"
            Shortcut        =   ^{F5}
            Tag             =   "Insert a Directory command (Lists files and Directories)"
         End
         Begin VB.Menu mnuLabel 
            Caption         =   "Label (Labels a Drive)"
            Shortcut        =   ^{F6}
            Tag             =   "Displays/Sets the label of a directory"
         End
         Begin VB.Menu dfgs 
            Caption         =   "-"
         End
         Begin VB.Menu mnuXcopy 
            Caption         =   "Xcopy (Copy Directory)"
            Shortcut        =   ^{F7}
            Tag             =   "Copies a directory"
         End
         Begin VB.Menu mnudgx 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDelTree 
            Caption         =   "DelTree"
            Shortcut        =   ^{F8}
            Tag             =   "Insert a Delete Directory command"
         End
         Begin VB.Menu mnuFormat 
            Caption         =   "Format"
            Shortcut        =   ^{F9}
            Tag             =   "Formats (Clears) a Disc/Directory"
         End
      End
      Begin VB.Menu mnuFileCmds 
         Caption         =   "File Commands"
         Tag             =   "Contains commands used to manipulate Files"
         Begin VB.Menu mnuC3 
            Caption         =   "Cancel"
         End
         Begin VB.Menu fgvc 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAttrib 
            Caption         =   "Attrib"
            Shortcut        =   +{F1}
            Tag             =   "Changes the attributes of a file"
         End
         Begin VB.Menu mnuRen 
            Caption         =   "Ren (Renames file)"
            Shortcut        =   +{F2}
            Tag             =   "Renames a file"
         End
         Begin VB.Menu fgfd 
            Caption         =   "-"
         End
         Begin VB.Menu mnuType 
            Caption         =   "Type (Print File fully)"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mnuMore 
            Caption         =   "More (Print pagewise)"
            Shortcut        =   +{F4}
         End
         Begin VB.Menu mnuEdit 
            Caption         =   "Edit File"
            Shortcut        =   +{F5}
         End
         Begin VB.Menu mnudsfs 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCopy 
            Caption         =   "Copy File"
            Shortcut        =   +{F6}
            Tag             =   "Insert a File Copy command"
         End
         Begin VB.Menu mnuMove 
            Caption         =   "Move (Cut 'n Paste)"
            Shortcut        =   +{F7}
            Tag             =   "Moves a file from one location to another"
         End
         Begin VB.Menu dfgdf 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDel 
            Caption         =   "Del"
            Shortcut        =   +{F8}
            Tag             =   "Insert a Delete File command"
         End
      End
      Begin VB.Menu mnuAscii 
         Caption         =   "Ascii Graphics"
         Tag             =   "Contains graphical stuff you can add to the batch file"
         Begin VB.Menu msf 
            Caption         =   "Cancel"
         End
         Begin VB.Menu sdk 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMsgBox 
            Caption         =   "Message Box..."
            Shortcut        =   ^M
            Tag             =   "Draws a custom box with your message in it"
         End
         Begin VB.Menu jhk 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLine 
            Caption         =   "Line..."
            Shortcut        =   ^L
            Tag             =   "Draws a line on the screen"
         End
         Begin VB.Menu jhp 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEcho 
            Caption         =   "Echo"
            Shortcut        =   ^E
         End
      End
      Begin VB.Menu mnuAdvCmds 
         Caption         =   "Advanced Commands"
         Begin VB.Menu mnuGOTOst 
            Caption         =   "Goto statment"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnuAddGoto 
            Caption         =   "Add GOTO point"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuCall 
            Caption         =   "Call Batchfile"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnusdhbsa 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSet 
            Caption         =   "Set Statement"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuCtrl 
            Caption         =   "Control Statments.."
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuExitCmd 
            Caption         =   "Exit Command"
            Shortcut        =   ^X
         End
      End
   End
   Begin VB.Menu mnuChart 
      Caption         =   "&ASCII Chart"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+-------------------------------------------------------+
'|     /---  ____ === /=|=\  ____ \   |                  |
'|     |   \ |  |  |  | | |  |  | |\  |    ver 1.00      |
'|     |===  |  |  |  | | |  |  | | \ |    (c)2004       |
'|     |   | |==|  |  | | |  |==| |  \|       Anil       |
'|     ----  |  |  |  |   |  |  | |   \       Gulecha    |
'|                                                       |
'|   (I really suck at this Text-LOGO drawing thing )    |
'+-------------------------------------------------------|
'| ABOUT BAT-Man AND IT'S AUTHOR                         |
'| -----------------------------                         |
'| Hi! I am Anil (GEEKFREEK).I live in a country called  |
'| INDIA.Look it up closely and you may find a city named|
'| Bangalore.Look even closer and you'll find me waving  |
'| at you:)                                              |
'| AGE : 18                                              |
'| Passion : Computers (DUHH!)                           |
'| WHAT I KNOW : C/C++ (not VC++), Flash & actionscript, |
'|               HTML ,BASIC (the DOS one),VB :) ,etc.   |
'|                                                       |
'| WHY BAT-Man : I made batman to help (i'm Kind-Hearted)|
'|               beginners understand many concepts of VB|
'|               The code is written in FULLY UNOPTIMISED|
'|               form so as to help anyone to understand |
'|               it's working without a hitch(or hatch). |
'|             (& also coz i wanna win code-of-the-month)|
'| The stuff you can learn:                              |
'| 1.Popup menus                                         |
'| 2.Using frames instead of space-hogging forms         |
'| 3.Making smooth Interfaces with use  of Keboard/Mouse |
'| 4.A lot of stuff about managing listbox ctrl          |
'| 5.Simple Easter-egg                                   |
'| 6.Making cute reading pages like these                |
'| 7.A spirit of Unending Euphoric Freedom (???)         |
'+-------------------------------------------------------+
'| DISCLAIMER                                            |
'| ----------                                            |
'| This program/Source code is provided asis and by using|
'| it you agree that i will hold no responsibility for   |
'| damage occuring to your computer by using it          |
'|                                                       |
'+-------------------------------------------------------|
'| HOW MAY I USE THIS SOURCE CODE  (LICENSE)             |
'| -----------------------------------------             |
'| YOU MAY CHANGE/COPY/ADD/REMOVE CODE FROM THIS FILE AS |
'| LONG AS YOU AGREE THAT :                              |
'| 1.I WILL NOT AND CANNOT HOLD ANY RESPONSIBILITY FOR   |
'|   ANY KIND OF DAMAGE THAT MAY HAPPEN TO YOUR COMPUTER |
'|   WHILE USING THIS.                                   |
'| 2.YOU WONT GO AROUND BOASTING THAT YOU WROTE THIS CODE|
'| 3.YOU WILL ASK FOR MY PERMISSION BEFORE YOU GO AROUND |
'|   DISTRIBUTING A CHANGED/UNCHANGED VERSION OF BAT-Man |
'| You will vote for me (or Die trying :)                |
'+-------------------------------------------------------|
'| ACKNOWLEDGEMENTS (I hope i spelt that right)          |
'| ----------------                                      |
'| 1.Night Wolf     - night_wolf_god@hotmail.com         |
'|    The GoldButton Used Throughout this project was    |
'|    made by him ( or her?).I obtained it from PSC.     |
'|    I have made few minor changes(added the SimClick   |
'|    method).He (she?) deserves a pat for the amazing   |
'|    button.(If you're reading this Send me a mail)     |
'| 2.TDavis   -At whose site i learnt a lot about DOS    |
'|    http://gearbox.maem.umr.edu/batch/batchtoc.htm     |
'|    Or if you prefer zipped:                           |
'|    http://gearbox.maem.umr.edu/batch/batchbk.zip      |
'+-------------------------------------------------------|
'| My Email - anilgulecha@yahoo.com                      |
'+-------------------------------------------------------|
'| MUSING                                                |
'| ------                                                |
'| I have always noticed code submitters take it for     |
'| granted that their audience is purely male. They go:  |
'| 'Hey Guys' and 'Hi dudes' and whatnot.Why is this?,I  |
'| ask myself.And  then i realise that proggrammming is  |
'| for guys only dudes.Can anyone of you seriously muse  |
'| of a gal at a computer.TOTALLY IMPOSSIBLE             |
'|                                                       |
'| (By the way,i first thought i wold gain the gal's sym-|
'| pathy & votes if i talked about their rights and stuff|
'| but my MALE mojo stopped me right.)                   |
'+-------------------------------------------------------|
'|BAT-Man Daily LOG                                      |
'+-------------------------------------------------------+
'| 25-08-04 Day-1                                        |
'| --------------                                        |
'| The idea of making a batchfile maker pops after i see |
'| a similar program on Planet-Source-Code Hall of fame  |
'| But that's nothing compared to what i have in mind.   |
'| I popup some Linkin Park in the background.Start the  |
'| project and make a list of commands i know and find a |
'| lot on the web.Make the menu systems,make the cool    |
'| icon and make an interface.                           |
'+-------------------------------------------------------+
'| 26-08-04 Day-2                                        |
'| --------------                                        |
'| Play Diablo-2 all day.Add the listbox Editing buttons |
'| Finish another project called Fast Split-DX completely|
'+-------------------------------------------------------|
'| 27-08-04 Day-3                                        |
'| --------------                                        |
'| It's raining softly on the window next to me.The birds|
'| are chirping softly,the lush grass is  inviting and   |
'| the landscape is rich as far as the eye can see.      |
'| ..                                                    |
'| ...                                                   |
'| ....                                                  |
'| .....                                                 |
'| ......                                                |
'| .....                                                 |
'| ...                                                   |
'| .                                                     |
'| ..                                                    |
'| ...                                                   |
'| ....                                                  |
'| BUT THAT AIN'T GONNA STOP GEEKFREEK FROM 'FREEKING'   |
'| Today is the coding day(Oh boy!).I slog 3 hours coding|
'| the static & directory commands.Find n Fix some bugs  |
'+-------------------------------------------------------+
'| 28-08-04 Day 4                                        |
'| --------------                                        |
'| Found the Horadic Staff in Diablo.                    |
'| BAT-man 'FILE'commands were coded today.              |
'| The 'Test' & 'Make' were                             |
'| implemented.I expect to finish everything before 01-09|
'+-------------------------------------------------------+
'| 29-08-04 Day 5                                        |
'| --------------                                        |
'| Sunday : Couldn't work on the system.                 |
'+-------------------------------------------------------|
'| 30-08-04 Day 6                                        |
'| --------------                                        |
'| Did the interface for Message Box and line commands.  |
'| Very Very tiring.Bugs creeping up unexpectedly like   |
'| Poor jokes in Hindi movies.Bye bye VB.Hello Diablo2   |
'| < 10 mins later >                                     |
'| Hi! i'm back.Can't stay away for long,can I?          |
'| OK,I now code the ASCII Chart generator.              |
'+-------------------------------------------------------+
'| 31-08-04 Day 7                                        |
'| --------------                                        |
'| Just watched Shrek 2.It ROCKS.The animation those guys|
'| made is simply awesome.                               |
'| Coded the msg box completely.Also the Line command    |
'| Thought of an Easter-Egg (those of you who dunno what |
'| EE is , well DUHH!!                                   |
'+-------------------------------------------------------|
'| 01-09-04 Day 8                                        |
'| --------------                                        |
'| Didn't do much on bat-man.Searched for DOS-Batch file |
'| reference on the net.If you're looking for something  |
'| like that i would recommend:                          |
'| http://gearbox.maem.umr.edu/batch/batchtoc.htm        |
'| Or if you prefer zipped:                              |
'| http://gearbox.maem.umr.edu/batch/batchbk.zip         |
'+-------------------------------------------------------|
'| 02-09-04 Day 9                                        |
'| --------------                                        |
'| OK,so i guess today's the last day.Added a few cmds   |
'| like load/open batch,call,comments,date/time-viewonly |
'| Did the Boring (Yawwn) Comment-adding stuff to every- |
'| thing.Excuse Me if it isn't enough.                   |
'| I already have an idea for the improvements in the nxt|
'| version but i giving it out depends upon your responce|
'| Go ahead ! Vote for me                                |
'+-------------------------------------------------------|
'| 03-09-04 Day 10                                       |
'| ---------------                                       |
'| Well I'm back.I'm what you call a typical programer-  |
'| Always back to improve .I wanna add copy/paste clpibrd|
'| support to the list box as that was the only thing i  |
'| found missing yesterday when I tested BAT-Man         |
'| I guess i'll submit today ?                           |
'+-------------------------------------------------------+
'| Code submitted on 03-09-04 Day 10                     |
'+-------------------------------------------------------+
  
 
'The code starts here (Oh BOY :)

Dim This_Is_The_Longest_Variable_In_The_Whole_World As Boolean, Hwnd_Actv As Byte
'below is one of the most imp. variable
Dim CurPos As Integer
Dim AttrStr As String
Dim curIdx As Byte
Dim Reply As String
Dim BatchFileName As String
Dim Temp As String
Dim ShowFile As Boolean
'fsys used for all file reading-Writing purposes
Dim Fsys As New FileSystemObject

Dim ClipText As String
Dim ClipLines(1 To 1000) As String
Dim clipNum As Integer


Private Sub Advanced_Click()
PopupMenu mnuAdvCmds, , CMDFrm.Left + Advanced.Left, CMDFrm.Top + Advanced.Height + Advanced.Top

End Sub

Private Sub AG_Click()
PopupMenu mnuAscii, , CMDFrm.Left + AG.Left, CMDFrm.Top + AG.Height + AG.Top
End Sub



Private Sub Apply_Click()
Select Case FrmOpt(curIdx).Tag
 Case "RF": AddStatic "Ren " & Trim(txtSrc) & " " & Trim(txtDes)
 Case "RD": AddStatic "Ren " & Trim(txtSrc) & " " & Trim(txtDes)
 Case "XC": AddStatic "Xcopy " & Trim(txtSrc) & " " & Trim(txtDes)
 Case "CF": AddStatic "Copy " & Trim(txtSrc) & " " & Trim(txtDes)
 Case "MF": AddStatic "Move " & Trim(txtSrc) & " " & Trim(txtDes)
 Case Else: MsgBox "error"

End Select

End Sub

Private Sub AttrCancel_Click()
OptClose
End Sub

Private Sub AttrGo_Click()
Dim AttrLet(0 To 3) As String
AttrLet(0) = ("a")
AttrLet(1) = ("h")
AttrLet(2) = ("r")
AttrLet(3) = ("s")

AttrStr = ""

For i = 0 To 3

If Y(i).Value = False And N(i).Value = True Then
 AttrStr = AttrStr & "-" & AttrLet(i) & " "
End If

If Y(i).Value = True And N(i).Value = False Then
 AttrStr = AttrStr & "+" & AttrLet(i) & " "
End If

Next i

If AttrStr = "" Then MsgBox "Nothing to do", vbInformation, "BAT-Man Batch-file maker": AttrCancel.SimClick: Exit Sub

'MsgBox "Attrib " & AttrStr & " " & Fn.Text
AddStatic "Attrib " & AttrStr & " " & Fn.Text

End Sub



Private Sub cboCtrl_Click()
'Display Required number of textboxes depending on
'selected control statment type
'code can be highly optimised but it is plainly
'written for beginners
Select Case cboCtrl.ListIndex

Case 0: lbl1.Visible = True: lbl2.Visible = True: lbl3.Visible = False: Ctrl1.Visible = True: Ctrl2.Visible = True: Ctrl3.Visible = False
        lbl1.Caption = "If Yes Goto :": Ctrl1.Text = ""
        lbl2.Caption = "If No Goto :": Ctrl2.Text = ""
Case 1: lbl1.Visible = True: lbl2.Visible = True: lbl3.Visible = True: Ctrl1.Visible = True: Ctrl2.Visible = True: Ctrl3.Visible = True
        lbl1.Caption = "If Abort Goto :": Ctrl1.Text = ""
        lbl2.Caption = "If Retry Goto :": Ctrl2.Text = ""
        lbl3.Caption = "If Cancel Goto :": Ctrl3.Text = ""
Case 2: lbl1.Visible = True: lbl2.Visible = True: lbl3.Visible = False: Ctrl1.Visible = True: Ctrl2.Visible = True: Ctrl3.Visible = False
        lbl1.Caption = "If 1 Goto :": Ctrl1.Text = ""
        lbl2.Caption = "If 2 Goto :": Ctrl2.Text = ""
Case 3: lbl1.Visible = True: lbl2.Visible = True: lbl3.Visible = True: Ctrl1.Visible = True: Ctrl2.Visible = True: Ctrl3.Visible = True
        lbl1.Caption = "If 1 Goto :": Ctrl1.Text = ""
        lbl2.Caption = "If 2 Goto :": Ctrl2.Text = ""
        lbl3.Caption = "If 3 Goto :": Ctrl3.Text = ""
Case 4: lbl1.Visible = True: lbl2.Visible = True: lbl3.Visible = False: Ctrl1.Visible = True: Ctrl2.Visible = True: Ctrl3.Visible = False
        lbl1.Caption = "If A Goto :": Ctrl1.Text = ""
        lbl2.Caption = "If B Goto :": Ctrl2.Text = ""
Case 5: lbl1.Visible = True: lbl2.Visible = True: lbl3.Visible = True: Ctrl1.Visible = True: Ctrl2.Visible = True: Ctrl3.Visible = True
        lbl1.Caption = "If A Goto :": Ctrl1.Text = ""
        lbl2.Caption = "If B Goto :": Ctrl2.Text = ""
        lbl3.Caption = "If C Goto :": Ctrl3.Text = ""
End Select

End Sub

Private Sub CboLine_Click()
'Show txtCustLine only if listindex=6
If CboLine.ListIndex = 6 Then
 txtCustLine.Visible = True
 txtCustLine.SetFocus
Else
 txtCustLine.Visible = False
End If

End Sub

Private Sub CDbatch_Click()
 'allow user to select a file to save batchfile code to
 CD2.ShowOpen
 If CD2.FileName = "" Then Exit Sub
 If Fsys.FileExists(CD2.FileName) Then r = MsgBox("You selected an existing file.OverWrite it?", vbYesNo, "Overwrite File")
 If r = vbNo Then Exit Sub

 
 BatchFileName = CD2.FileName
 lblSaveFile.Caption = GetFile(BatchFileName)
 WriteBatchFile
End Sub

Private Sub CDButton_Click()

CD1.ShowOpen
Fn.Text = CD1.FileName
Fn.SetFocus
End Sub



Private Sub CDDes_Click()
'let user select file/Folder depending on tag
Select Case FrmOpt(curIdx).Tag
 Case "RF": ShowFile = True
 Case "CF": ShowFile = True
 Case "MF": ShowFile = True
 Case Else: ShowFile = False
End Select


If ShowFile Then
CD1.ShowOpen
txtDes.Text = CD1.FileName
txtDes.SetFocus
Else
txtDes.Text = BrowseForDirectory("Select Directory")
txtDes.SetFocus
End If

End Sub

Private Sub CDSource_Click()
'Let user select file/folder depending on tag
Select Case FrmOpt(curIdx).Tag
 Case "RF": ShowFile = True
 Case "CF": ShowFile = True
 Case "MF": ShowFile = True
 Case Else: ShowFile = False
End Select

If ShowFile Then
CD1.ShowOpen
txtSrc.Text = CD1.FileName
txtSrc.SetFocus
Else
txtSrc.Text = BrowseForDirectory("Select Directory")
txtSrc.SetFocus
End If

End Sub



Private Sub Criteria_Click()
'Enable wild textbox if criteria is checked
If Criteria.Value = 0 Then
 Wild.Enabled = False
Else
 Wild.Enabled = True
 Wild.SelStart = 0
 Wild.SelLength = Len(Wild.Text)
 Wild.SetFocus
End If
End Sub

Private Sub ctrlCancel_Click()
'close mini-frame
OptClose
End Sub

Private Sub ctrlOK_Click()
'Add the appropriate lines of code
'(letting BAT-Man do the dirty code work)
'Pretty self-Explanatory
'If your all scared reading ERRORLEVEL everywhere you
'better goto the site i have mentioned in the log(Day-8)

Select Case cboCtrl.ListIndex
Case 0: If ctrlPrompt.Text <> "" Then AddStatic "echo " & ctrlPrompt.Text
        AddStatic "Choice /c:yn Choose an option"
        AddStatic "If ERRORLEVEL 2 goto " & Trim(Ctrl2.Text)
        AddStatic "If ERRORLEVEL 1 goto " & Trim(Ctrl1.Text)
        Exit Sub

Case 1: If ctrlPrompt.Text <> "" Then AddStatic "echo " & ctrlPrompt.Text
        AddStatic "Choice /c:arc Choose an option"
        AddStatic "If ERRORLEVEL 3 goto " & Trim(Ctrl3.Text)
        AddStatic "If ERRORLEVEL 2 goto " & Trim(Ctrl2.Text)
        AddStatic "If ERRORLEVEL 1 goto " & Trim(Ctrl1.Text)
        Exit Sub

Case 2: If ctrlPrompt.Text <> "" Then AddStatic "echo " & ctrlPrompt.Text
        AddStatic "Choice /c:12 Choose an option"
        AddStatic "If ERRORLEVEL 2 goto " & Trim(Ctrl2.Text)
        AddStatic "If ERRORLEVEL 1 goto " & Trim(Ctrl1.Text)
        Exit Sub

Case 3: If ctrlPrompt.Text <> "" Then AddStatic "echo " & ctrlPrompt.Text
        AddStatic "Choice /c:123 Choose an option"
        AddStatic "If ERRORLEVEL 3 goto " & Trim(Ctrl3.Text)
        AddStatic "If ERRORLEVEL 2 goto " & Trim(Ctrl2.Text)
        AddStatic "If ERRORLEVEL 1 goto " & Trim(Ctrl1.Text)
        Exit Sub

Case 4: If ctrlPrompt.Text <> "" Then AddStatic "echo " & ctrlPrompt.Text
        AddStatic "Choice /c:ab Choose an option"
        AddStatic "If ERRORLEVEL 2 goto " & Trim(Ctrl2.Text)
        AddStatic "If ERRORLEVEL 1 goto " & Trim(Ctrl1.Text)
        Exit Sub

Case 5: If ctrlPrompt.Text <> "" Then AddStatic "echo " & ctrlPrompt.Text
        AddStatic "Choice /c:abc Choose an option"
        AddStatic "If ERRORLEVEL 3 goto " & Trim(Ctrl3.Text)
        AddStatic "If ERRORLEVEL 2 goto " & Trim(Ctrl2.Text)
        AddStatic "If ERRORLEVEL 1 goto " & Trim(Ctrl1.Text)
        Exit Sub
End Select

End Sub

Private Sub CustSet_Change(Index As Integer)
'Check for valid entry b/w 1-255
Dim Chek As Integer
On Error GoTo a:
Chek = Int(CustSet(Index))
If Chek < 256 And Chek > 0 Then Exit Sub
a:
CustSet(Index).SelStart = 0
CustSet(Index).SelLength = Len(CustSet(Index).Text)
MsgBox "Enter valid Ascii Code ( 1 - 255 )", vbInformation, "Invalid Entry"

End Sub

Private Sub DelCurLine_Click()
'Delete the current selected line (curpos)

If LB.ListIndex = -1 Then Beep: Exit Sub

LB.RemoveItem LB.ListIndex

If CurPos < LB.ListCount Then
LB.ListIndex = CurPos
Else
LB.ListIndex = LB.ListCount - 1
CurPos = LB.ListIndex
End If

End Sub

Private Sub DirCancel_Click()
'close mini-frame
OptClose
End Sub

Private Sub DirCmds_Click()
'popup the directory commands

PopupMenu mnuDirCmds, , CMDFrm.Left + DirCmds.Left, CMDFrm.Top + DirCmds.Height + DirCmds.Top
End Sub

Private Sub DirOk_Click()
'Create the command string depending on user
'selections and add them to the list box

Dim DirStr As String
DirStr = "Dir"
If lookInSub.Value = 1 Then DirStr = DirStr & "/s"
If WideView.Value = 1 Then DirStr = DirStr & "/w"
If PagePause.Value = 1 Then DirStr = DirStr & "/p"
If Criteria.Value = 1 Then DirStr = DirStr & " " & Trim(Wild.Text)
AddStatic DirStr
End Sub

Private Sub Down_Click()
'Move current selected line down
If LB.ListIndex < 0 Then Beep: Exit Sub
If LB.ListIndex = LB.ListCount - 1 Then Beep: Exit Sub

Temp = LB.List(LB.ListIndex)
LB.List(LB.ListIndex) = LB.List(LB.ListIndex + 1)
LB.List(LB.ListIndex + 1) = Temp
CurPos = CurPos + 1
LB.ListIndex = CurPos

End Sub

Private Sub EditManually_Click()
'Allow the user to edit the selected line
If LB.ListIndex < 1 Then Beep: Exit Sub
Dim DifStr As String
DifStr = "Edit the line." & vbCrLf & "Note : In case of a syntax error the command may not work as expected."

DifStr = InputBox(DifStr, "Edit Manually", LB.List(LB.ListIndex))
If DifStr = "" Then Exit Sub
LB.List(LB.ListIndex) = DifStr


End Sub

Private Sub FilCmds_Click()
'popup the file commands
PopupMenu mnuFileCmds, , CMDFrm.Left + FilCmds.Left, CMDFrm.Top + FilCmds.Height + FilCmds.Top
End Sub

Private Sub Fn_GotFocus()
'select everything when fn gets focus
Fn.SelStart = 0
Fn.SelLength = Len(Fn.Text)
End Sub


Private Sub Form_Load()

If Fsys.FileExists("Test.bat") Then Fsys.DeleteFile "Test.bat", True

'Curpos is the current position in the code list(LB) to
'enter lines.Will be used by ADDSTATIC command
CurPos = -1
Image1.Picture = frmMain.Icon

'SET the pictures for all buttons in the program
'The pictures are loaded once during design
'The rest are copied.Shortens the exe size (thats what i think)

Set AttrCancel.Picture = DelCurLine.Picture
Set CDDes.Picture = CDButton.Picture
Set CDSource.Picture = CDButton.Picture
Set SDFCancel.Picture = DelCurLine.Picture
Set DirCancel.Picture = DelCurLine.Picture
Set CDbatch.Picture = CDButton.Picture
Set MsgCancel.Picture = DelCurLine.Picture
Set LineCancel.Picture = DelCurLine.Picture
Set ctrlCancel.Picture = DelCurLine.Picture
Set MsgHelp.Picture = ListHelp.Picture

'select the first option in the combo boxes
SelectStyle.ListIndex = 0
cboCtrl.ListIndex = 0
BoxWidth = 79
frmMain.Height = 6660
SetLen

End Sub

Private Sub Form_Unload(Cancel As Integer)
'DO NOT EVEN ATTEMPT TO DARE TO THINK ABOUT REMOVING
'THE LINES BELOW. :(
MsgBox "I worked hard.Vote for ME,Plzz", vbInformation + vbOKOnly, "Very Important and Pleazing Message"
MsgBox "You're voting for me, OK.", vbInformation + vbOKOnly, "Utterly Devastatingly important message"
Cancel = 0
End Sub


Private Sub Image1_DblClick()
'ACTIVATE WINDOW OPTION .USED LATER
Hwnd_Actv = Hwnd_Actv + 1
If Hwnd_Actv > 10 Then Hwnd_Actv = 0:    MsgBox "You are the greatest!", vbInformation, "Ego Blower"
End Sub

Private Sub Insert_Click()

'The code to insert a line of code at the current position(cur-pos
If LB.ListIndex < 0 Then Beep: Exit Sub
LB.AddItem "", LB.ListIndex
LB.ListIndex = CurPos

Dim DifStr As String
DifStr = "Type code" & vbCrLf & "Note : In case of a syntax error the command may not work as expected."
DifStr = InputBox(DifStr, "Insert", LB.List(LB.ListIndex))
'Do nothing if user cancelled
If DifStr = "" Then Exit Sub
LB.List(LB.ListIndex) = DifStr

End Sub






Private Sub LB_Click()
'Set curpos when user changes selection
CurPos = LB.ListIndex
End Sub

Private Sub LB_DblClick()
'Bring up the edit box when user double-clicks LB
EditManually.SimClick
End Sub



Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
'The keyboard interface for editing,deleting,moving
'stuff in LB

Select Case KeyCode

'Move line up
Case 38: 'shift up
  If Shift = 1 Then
  If LB.ListIndex < 1 Then Beep: Exit Sub
  Temp = LB.List(LB.ListIndex)
  LB.List(LB.ListIndex) = LB.List(LB.ListIndex - 1)
  LB.List(LB.ListIndex - 1) = Temp
 End If
 
'Move line down
Case 40: 'Shift down
 If Shift = 1 Then
  If LB.ListIndex < 0 Then Beep: Exit Sub
  If LB.ListIndex = LB.ListCount - 1 Then Beep: Exit Sub
  Temp = LB.List(LB.ListIndex)
  LB.List(LB.ListIndex) = LB.List(LB.ListIndex + 1)
  LB.List(LB.ListIndex + 1) = Temp
 End If


Case 45: 'Insert
 'Insert a line and ask for editing it
 If Shift = 0 Then
   Insert.SimClick
 
 Else 'shift insert
 'Just insert a plain line
  If LB.ListIndex < 0 Then Beep: Exit Sub
  LB.AddItem "", LB.ListIndex
  LB.ListIndex = CurPos

 End If

'Delete current selected line
Case 46: 'del
 'Simulate as if DELCurLine were clicked
 DelCurLine.SimClick
 Exit Sub

Case 13: 'Enter
 'Simulate as if Edit manually were clicked
 EditManually.SimClick
 Exit Sub


Case Else
 Exit Sub

End Select
End Sub


Private Sub LB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 2 Then Exit Sub
ClipText = Clipboard.GetText

If ClipText = "" Then
 mnuPaste.Enabled = False
 If LB.SelCount = 0 Then mnuClipCopy.Enabled = False
 PopupMenu mnuClip
 Exit Sub
End If

mnuPaste.Enabled = True
If LB.SelCount = 0 Then mnuClipCopy.Enabled = False

If Right(ClipText, 2) <> vbCrLf Then ClipText = ClipText + vbCrLf
Dim Pos As Integer
clipNum = 0


Do
 Pos = InStr(ClipText, vbCrLf)
 If Pos = 0 Then Exit Do
 clipNum = clipNum + 1
 ClipLines(clipNum) = Mid(ClipText, 1, Pos - 1)
 ClipText = Mid(ClipText, Pos + 2)
Loop Until Pos = 0

PopupMenu mnuClip

End Sub

Private Sub LineCancel_Click()
' Close Line mini-Frame
OptClose

End Sub

Private Sub LineOK_Click()
'Insert a Line to LB
Select Case CboLine.ListIndex
 Case 0: AddStatic "Echo " & String(79, Chr(223))
 Case 1: AddStatic "Echo " & String(79, Chr(196))
 Case 2: AddStatic "Echo " & String(79, Chr(205))
 Case 3: AddStatic "Echo " & String(79, Chr(240))
 Case 4: AddStatic "Echo " & String(79, Chr(250))
 Case 5: AddStatic "Echo " & String(79, Chr(45))
 Case 6: AddStatic "Echo " & String(79, Chr(CInt(txtCustLine.Text)))
End Select

End Sub

Private Sub ListHelp_Click()
'Show the user the keyboard shortcuts in a msgbox
Dim Hlp As String
Hlp = "Select any item in the code list.Then press:" & vbCrLf _
& "-------------------------------" & vbCrLf _
& "Enter                   : To edit it" & vbCrLf _
& "Delete                 : To remove it" & vbCrLf _
& "-------------------------------" & vbCrLf _
& "Shift + Insert       : To add a blank line" & vbCrLf _
& "Shift + Up Arrow : To move line up" & vbCrLf _
& "Shift + Down      : To move line down" & vbCrLf _
& "-------------------------------" & vbCrLf _
& "Insert                  : Insert your own code" & vbCrLf _
& "-------------------------------" & vbCrLf _
& "You can also Copy & Paste lines in BAT-Man" & vbCrLf _
& "and also from other text editing programs." & vbCrLf
MsgBox Hlp, vbInformation, "Help"

End Sub

Private Sub lstMSG_DblClick()

'Allow the user to edit the selected line
If lstMSG.ListIndex < 0 Then Beep: Exit Sub
Dim DifStr As String
DifStr = "Edit the line." & vbCrLf & "Note : The string must not be longer than the (BoxWidth-3)"

DifStr = InputBox(DifStr, "Edit Manually", lstMSG.List(lstMSG.ListIndex))
If DifStr = "" Then Exit Sub
lstMSG.List(lstMSG.ListIndex) = DifStr



End Sub

Private Sub lstMSG_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

 Case 46: 'del
'Delete the cur entry
 If lstMSG.ListCount = 1 Then lstMSG.Clear
 Dim TemIdx As Integer
 TemIdx = lstMSG.ListIndex
 
 If lstMSG.ListIndex = -1 Then Beep: Exit Sub

 lstMSG.RemoveItem lstMSG.ListIndex
 On Error GoTo e:

 lstMSG.ListIndex = TemIdx
 Exit Sub
e:
 If lstMSG.ListCount = 0 Then Exit Sub
 lstMSG.ListIndex = lstMSG.ListCount - 1
 Exit Sub

End Select

End Sub

Private Sub Make_Click()

'Make the batch file
If BatchFileName = "" Then
 'If user already hasn't selected a file let him
 CD2.ShowOpen
 'Exit sub if user pressed cancelled
 If CD2.FileName = "" Then Exit Sub
 
 BatchFileName = CD2.FileName
 If Fsys.FileExists(CD2.FileName) Then r = MsgBox("You selected an existing file.OverWrite it?", vbYesNo, "Overwrite File")
 If r = vbNo Then Exit Sub
 lblSaveFile.Caption = GetFile(BatchFileName)
 'Write the batch file
 WriteBatchFile
Else
 'If the user has selected a file befor overwrite it
 WriteBatchFile
End If

End Sub



Private Sub mnuAbout_Click()
Dim ABT As String
ABT = "   BAT-Man Batch File Maker/Editor" & vbCrLf _
& "-----------------------------------------------------------" & vbCrLf _
& "-----------------------------------------------------------" & vbCrLf _
& "(c)2004 by Anil 'GeeKFreeK' Gulecha (Programer)" & vbCrLf _
& "" & vbCrLf _
& "This program is disrtibuted along with it's source code." & vbCrLf _
& "For Disclaimer/License view code." & vbCrLf _
& "" & vbCrLf _
& "For bugs / suggestion / anything else :" & vbCrLf _
& "Email : anilgulecha@yahoo.com" & vbCrLf _
& "-------------------------------------------------" & vbCrLf

MsgBox ABT, vbInformation, "About BAT_Man"

End Sub

Private Sub mnuAddGoto_Click()

'Add a goto point
Reply = InputBox("Name of GOTO point" & vbCrLf & "GOTO points are not case sensitive" & vbCrLf & "BAT-Man will also add an exit command for your convenience.", "GOTO")
'Exit sub if user cancelled
If Reply = "" Then Exit Sub

'Note : I have not used ADDSTATIC as i want to add the
'goto point at the end of the list and not at current position

LB.AddItem ""
LB.AddItem ":" & UCase(Trim(Reply))
LB.AddItem "Exit"

End Sub

Private Sub mnuAttrib_Click()
'Show the Change attrib mini-Frame
OptShow (0)
For i = 0 To 3
N(i).Value = False
Y(i).Value = False
Next i
Fn.SetFocus
End Sub

Private Sub mnuCall_Click()

'Insert a call command

Reply = InputBox("Enter batchfile Path/Name [args] :", "Call")
If Reply = "" Then Exit Sub
AddStatic "Call " & Reply

End Sub

Private Sub mnuCD_Click()
'Insert a cd command

Reply = InputBox("Change to which directory" & vbCrLf & "Note: Enter valid directory", "Enter Directory Name")
If Reply = "" Then Exit Sub
AddStatic "CD " & Reply
End Sub

Private Sub mnuChart_Click()
'The cool ASCII-CHART generator

'If the files have been created and have not been changed
'then just call them.
If Fsys.FileExists("Ascii.bat") And Fsys.FileExists("Ascii.asc") Then
 If (FileLen("Ascii.asc") = 1782 And FileLen("Ascii.bat") = 72) Then Shell "Ascii.bat", vbNormalFocus: Exit Sub
End If


Dim AA, BB, XX$, YY$
Set AA = Fsys.CreateTextFile("Ascii.asc", True)
Dim Oneline As String
'The first line using doubleline chars
AA.WriteLine Chr(213) & String(77, Chr(205)) & Chr(184)
'The second line with name
AA.WriteLine Chr(179) & "  ASCII Code Chart generated by BAT-Man" & String(38, " ") & Chr(179)

XX = String(5, Chr(196)) & Chr(194)
YY = String(5, Chr(196)) & Chr(194)
For i = 1 To 11
YY = YY + XX
Next i
'The third line
'I dont have the time to explain
AA.WriteLine Chr(195) & YY & String(5, 196) & Chr(180)


'Generate the numbers and code
'You will notice that the chart starts from 27 onwards
'I would have made the chart from 0 to 255 but DOS freaks
'out when it sees some character below 57 and starts showing
'errors.If you still insist try changing the below from
'1 to 18 and see what happens.Dont say i didn't warn you
For i = 2 To 18
 Oneline = Chr(179)
 For j = 1 To 13
 Oneline = Oneline + Format((i * 13) + j, "000") & " " & Chr((i * 13) + j) & Chr(179)
 Next j
AA.WriteLine Oneline

Next i
Oneline = Chr(179)
'Generate last line with following spaces
For i = 248 To 255
 Oneline = Oneline + Format(i, "000") & " " & Chr(i) & Chr(179)
Next i
For i = 1 To 5
 Oneline = Oneline & "     " & Chr(179)
Next i

AA.WriteLine Oneline

YY = ""
XX = String(5, Chr(205)) & Chr(207)
For i = 1 To 12
YY = YY + XX
Next i

AA.WriteLine Chr(212) & YY & String(5, 205) & Chr(190)

'Close ASCII.asc
AA.Close

'Type out the ascii.asc using a batch file Ascii.bat
Set BB = Fsys.CreateTextFile("Ascii.Bat", True)

BB.WriteLine "@echo off"
BB.WriteLine "cls"
BB.WriteLine "type Ascii.asc"
BB.WriteLine "Echo Hit a key to exit"
BB.WriteLine "Pause>nul"
BB.WriteLine "cls"
BB.Close

'Open ascii.bat
Shell "Ascii.bat", vbNormalFocus

End Sub

Private Sub mnuClipCopy_Click()

ClipText = ""

For i = 0 To LB.ListIndex
If LB.Selected(i) Then ClipText = ClipText & LB.List(i) & vbCrLf
Next i

Clipboard.SetText (ClipText)

End Sub

Private Sub mnuCls_Click()
'Add cls at curpos
AddStatic "Cls"
End Sub





Private Sub mnuComment_Click()
'insert a comment at curpos
Reply = InputBox("Enter Comment", "Comment")
If Reply = "" Then Exit Sub
AddStatic "::" & Reply

End Sub

Private Sub mnuCopy_Click()
'Show the mini-copy frame
'Change its title and tag to help the open file/Folder btn
OptShow 1, "CF", "Copy File"
txtSrc.Text = ""
txtDes.Text = ""
txtSrc.SetFocus
End Sub

Private Sub mnuCtrl_Click()
'Show mini control statment-frame
OptShow 4
cboCtrl.SetFocus
End Sub

Private Sub mnuDate_Click()
'Add the date-view only command
'To view the command select mnuDate in the properties
'window and look under the 'Tag' Property
AddStatic mnuDate.Tag
End Sub

Private Sub mnuDateset_Click()
'Add normal date command
AddStatic "Date"
End Sub

Private Sub mnuDel_Click()
'Add the delete file command
Reply = InputBox("Enter the file to delete." & vbCrLf _
& vbCrLf & "Note : USE THIS COMMAND WITH CAUTION.IT CAN" _
& " POTENTIALLY DELETE IMPORTANT DATA ON THE SYSTEM.", "Del")
'Exit sub if user cancelled
If Reply = "" Then Exit Sub
AddStatic "Del " & Trim(Reply)
End Sub

Private Sub mnuDelTree_Click()
'Add the delete Directory command
Reply = InputBox("Enter the directory to delete." & vbCrLf _
& vbCrLf & "Note : USE THIS COMMAND WITH CAUTION.IT CAN" _
& " POTENTIALLY DELETE IMPORTANT DATA ON THE SYSTEM.", "DelTree")
'Exit sub if user cancelled
If Reply = "" Then Exit Sub
AddStatic "Deltree " & Trim(Reply)
End Sub

Private Sub mnuDir_Click()
'Show the dir mini-Frame
OptShow (2)
End Sub

Private Sub mnuEcho_Click()
'add the echo command
Reply = InputBox("Echo What ?" & vbCrLf & "For an empty line enter a dot '.'", "Echo")
'Exit sub if user acncelled
If Reply = "" Then Exit Sub
If Reply = "." Then AddStatic "Echo.": Exit Sub
AddStatic "Echo " & Reply

End Sub

Private Sub mnuEchoOff_Click()
'Add the echo off command
'TidBit : Someone tells me this is the most commonly
'         used command in batch files
AddStatic "@echo off"
End Sub

Private Sub mnuEchoOn_Click()
'Add the echo on command
AddStatic "@echo on"
End Sub

Private Sub mnuEdit_Click()
'Ad the edit command
Reply = InputBox("Enter the file to edit using the MS-DOS editor.", "Edit")
If Reply = "" Then Exit Sub
AddStatic "Edit " & Trim(Reply)
End Sub

Private Sub mnuExit_Click()
'Exit BAT-Man
Unload Me
End Sub

Private Sub mnuExitCmd_Click()
'Add the exit command at curpos
AddStatic "Exit"
End Sub

Private Sub mnuFormat_Click()
'Add the for mat command
'I dare you you make a batch like this:
'Be sure to read the license at the top befor trying
'
'@echo off
'Echo Y | Format C:
'

Reply = InputBox("Enter the Drive to format. Eg: 'A:'" & vbCrLf _
& vbCrLf & "NOTE : USE THIS COMMAND WITH CAUTION.IT CAN" _
& " POTENTIALLY DELETE IMPORTANT DATA ON THE SYSTEM.", "Format", "A:")
If Reply = "" Then Exit Sub
AddStatic "Format " & Trim(Reply)
End Sub

Private Sub mnuGOTOst_Click()
'Add goto statement
Reply = InputBox("Goto where", "GOTO")
If Reply = "" Then Exit Sub
AddStatic "Goto " & Trim(Reply)
End Sub

Private Sub mnuInsert_Click()
'Inserts the contents of another batchfile at the
'current positinn
CD2.ShowOpen
If CD2.FileName = "" Then Exit Sub
BatchFileName = CD2.FileName
Dim Lfil
Set Lfil = Fsys.OpenTextFile(BatchFileName, ForReading)
'read from file until the stream ends
Do While Lfil.AtEndOfStream = False
AddStatic Lfil.ReadLine
Loop

End Sub

Private Sub mnuLabel_Click()
'Adds Label command
Reply = InputBox("Enter Drivename (eg - 'C:')" & vbCrLf _
& "Leave plain for default drive" & vbCrLf & vbCrLf _
& "Advanced : Enter '[Drive] [Label To Set]' below", "Label")
AddStatic Trim("Label " & Reply)
End Sub

Private Sub mnuLine_Click()
'Show line mini-frame
OptShow (3)
End Sub

Private Sub mnuMake_Click()
'Call make button click event
Make.SimClick
End Sub

Private Sub mnuMD_Click()
'add Make directory
Reply = InputBox("Name of directory to create" & vbCrLf & "Note: Enter valid alpha-numeric name.'_' is also valid", "Enter Directory Name")
If Reply = "" Then Exit Sub
AddStatic "MD " & Reply

End Sub

Private Sub mnuMore_Click()
'Add MORE command to read file page-by-page
Reply = InputBox("File to read page-by-page." & vbCrLf & "Note: File must be valid", "File to read pagewise")
If Reply = "" Then Exit Sub
AddStatic "More< " & Reply
End Sub

Private Sub mnuMove_Click()
'Show move file frame
'Set tag to MF to enable working of browse file/Folder btn
OptShow 1, "MF", "Move file"
txtSrc.Text = ""
txtDes.Text = ""
txtSrc.SetFocus
End Sub

Private Sub mnuMsgBox_Click()
'Show the Message box frame to add a COOL box to the file
frmMain.Height = 8700
frmMsgBox.Enabled = True
SelectStyle.ListIndex = 0

End Sub

Private Sub mnuNew_Click()

If LB.ListCount <> 0 Then
 Reply = MsgBox("The current code will be cleared.Continue?", vbYesNo + vbInformation, "Confirm?")
 If Reply = vbNo Then Exit Sub
End If

LB.Clear
BatchFileName = ""
lblSaveFile.Caption = ""
LB.SetFocus

End Sub

Private Sub mnuOpen_Click()
'if somethin is present in List then confirm if user
'wants to remove them
If LB.ListCount <> 0 Then
 Reply = MsgBox("By opening another file, you will lose the current work. Do you want to continue?", vbInformation + vbYesNo, "Confirm")
 If Reply = vbNo Then Exit Sub
End If

'Open a file
CD2.ShowOpen
If CD2.FileName = "" Then Exit Sub
BatchFileName = CD2.FileName
Dim Lfil
Set Lfil = Fsys.OpenTextFile(BatchFileName, ForReading)

LB.Clear
'read contents to LB

Do While Lfil.AtEndOfStream = False
LB.AddItem Lfil.ReadLine
Loop

LB.ListIndex = LB.ListCount - 1
CurPos = LB.ListIndex

End Sub

Private Sub mnuPaste_Click()

For i = 1 To clipNum
AddStatic ClipLines(i)
Next i

End Sub

Private Sub mnuPause_Click()
'Add the command
AddStatic "Pause"
End Sub

Private Sub mnuRD_Click()
'Add the command
Reply = InputBox("Name of directory to remove" & vbCrLf & "Note: Directory must be valid.It should also be empty. (Else use DelTree command)", "Enter Directory Name")
If Reply = "" Then Exit Sub
AddStatic "RD " & Reply
End Sub

Private Sub mnuRen_Click()

'Add the command
OptShow 1, "RF", "Rename file"
txtSrc.Text = ""
txtDes.Text = ""
txtSrc.SetFocus
End Sub

Private Sub mnuRenDir_Click()
'Add the command
OptShow 1, "RD", "Rename Directory"
txtSrc.Text = ""
txtDes.Text = ""
txtSrc.SetFocus
End Sub

Private Sub mnuRestart_Click()
'Add the command
AddStatic "Restart"
End Sub


Private Sub mnuSelfDel_Click()
'Let the batch file delete itself
AddStatic "del %0.bat"
End Sub

Private Sub mnuSet_Click()
'Add the command
Reply = InputBox("Enter Set Statement" & vbCrLf & "Syntax : Set [var]=[value]", "Set", "Set var=??")
If Reply = "" Then Exit Sub
AddStatic Trim(Reply)

End Sub

Private Sub mnuShutDown_Click()
'Add the command
' Doesn't seem to work on some DOS vers
AddStatic "RUNDLL.EXE user.exe,exitwindows"
End Sub

Private Sub mnuTest_Click()
'Add the command
Test.SimClick
End Sub

Private Sub mnuTime_Click()
'Add the command
AddStatic mnuTime.Tag
End Sub

Private Sub mnuTimeset_Click()
'Add the command
AddStatic "Time"
End Sub

Private Sub mnuTree_Click()
'Add the command
' Doesn't work in all vers.
AddStatic "Tree"
End Sub

Private Sub mnuType_Click()
'Add the command
Reply = InputBox("Type" & vbCrLf & "Note: File must be valid and small enough to appear completely on screen. (Else use 'More')", "File to type")
If Reply = "" Then Exit Sub
AddStatic "Type " & Reply
End Sub

Private Sub mnuVer_Click()
'Add the command
AddStatic "Ver"

End Sub

Private Sub mnuWin_Click()
'Add the command
AddStatic "Win"
End Sub

Private Sub mnuXcopy_Click()
'show the copy dir mini-frame
OptShow 1, "XC", "Copy directory"
txtSrc.Text = ""
txtDes.Text = ""
txtSrc.SetFocus
End Sub

Private Sub MsgCancel_Click()
'Add the command
frmMain.Height = 6660
frmMsgBox.Enabled = False
End Sub

Private Sub MsgDown_Click()
'Add the command
lstMSG.AddItem txtLineMSG.Text, lstMSG.ListIndex + 1
lstMSG.ListIndex = lstMSG.ListIndex + 1
txtLineMSG.Text = ""
txtLineMSG.SetFocus
End Sub

Private Sub MsgHelp_Click()
Dim Hlp As String
'Show the user the keyboard shortcuts in a msgbox
Hlp = "Select any item in the code list.Then press:" & vbCrLf _
& "-------------------------------" & vbCrLf _
& "Delete                 : To remove it" & vbCrLf _
& "-------------------------------" & vbCrLf _
& "Double-Click         : To Edit the line" & vbCrLf _
& "-------------------------------" & vbCrLf _
& "After selecting an item,the next item will be added below it" & vbCrLf

MsgBox Hlp, vbInformation, "Help"



End Sub

Private Sub MsgOK_Click()
  'First set the correspondin ascii characters
  Select Case SelectStyle.ListIndex
  Case 0: SetBoxStuff 219, 223, 219, 219, 219, 220, 219, 219, txtTitle.Text
  Case 1: SetBoxStuff 218, 196, 191, 179, 217, 196, 192, 179, txtTitle.Text
  Case 2: SetBoxStuff 201, 205, 187, 186, 188, 205, 200, 186, txtTitle.Text
  'In case of custron set the user entered chars
  Case 3:   r = SetBoxStuff(CInt(CustSet(0).Text), CInt(CustSet(1).Text), _
            CInt(CustSet(2).Text), CInt(CustSet(3).Text), CInt(CustSet(4).Text), _
            CInt(CustSet(5).Text), CInt(CustSet(6).Text), CInt(CustSet(7).Text), _
            txtTitle.Text)
            If r = False Then
            CustSet(0).SetFocus
            MsgBox "Wrong values", vbInformation, "Invalid entries"
            Exit Sub
            End If
            
  End Select

'add the first line
AddStatic "Echo " & Lline("`")
'add a plain line
AddStatic "Echo " & Lline("")

' add the message
For i = 0 To lstMSG.ListCount - 1
AddStatic "Echo " & Lline(lstMSG.List(i))
Next i

' add a plain line
AddStatic "Echo " & Lline("")
'add the last line
AddStatic "Echo " & Lline("~")

End Sub

Private Sub msgTest_Click()
  'First set the ascii chars
  Select Case SelectStyle.ListIndex
  Case 0: SetBoxStuff 219, 223, 219, 219, 219, 220, 219, 219, txtTitle.Text
  Case 1: SetBoxStuff 218, 196, 191, 179, 217, 196, 192, 179, txtTitle.Text
  Case 2: SetBoxStuff 201, 205, 187, 186, 188, 205, 200, 186, txtTitle.Text
  'If custom set the user entered chars
  Case 3:   r = SetBoxStuff(CInt(CustSet(0).Text), CInt(CustSet(1).Text), _
            CInt(CustSet(2).Text), CInt(CustSet(3).Text), CInt(CustSet(4).Text), _
            CInt(CustSet(5).Text), CInt(CustSet(6).Text), CInt(CustSet(7).Text), _
            txtTitle.Text)
            If r = False Then
            CustSet(0).SetFocus
            MsgBox "Wrong values", vbInformation, "Invalid entries"
            Exit Sub
            End If
            
  End Select

Dim Mtest
'open the test file
Set Mtest = Fsys.CreateTextFile("Mtest.bat")
Mtest.WriteLine "@Echo off"
Mtest.WriteLine "cls"
Mtest.WriteLine "echo This is how the message box will look (If you've turned echo off) :"
Mtest.WriteLine "echo."
'Add the msgbox line
Mtest.WriteLine "Echo " & Lline("`")
Mtest.WriteLine "Echo " & Lline("")

For i = 0 To lstMSG.ListCount - 1
Mtest.WriteLine "Echo " & Lline(lstMSG.List(i))
Next i

Mtest.WriteLine "Echo " & Lline("")
Mtest.WriteLine "Echo " & Lline("~")
Mtest.WriteLine "echo."
Mtest.WriteLine "Echo Hit a key to exit"
'wait for user to press a key
Mtest.WriteLine "Pause>nul"
Mtest.WriteLine "cls"
Mtest.Close
'run the test file
Shell "mtest.bat", vbNormalFocus
MsgBox "Press OK after testing", vbInformation, "Testing"
Fsys.DeleteFile "mtest.bat", True

End Sub


Private Sub SDFCancel_Click()
'close the frame
Apply.Enabled = False
OptClose
End Sub


Private Sub SelectStyle_Click()
'set visibility of frmcustset depending on ss.listindex
If SelectStyle.ListIndex = 3 Then
  FrmCustSet.Visible = True
  CustSet(0).SetFocus
Else
  FrmCustSet.Visible = False
End If
End Sub

Private Sub StaticButton_Click()
'popup the static commands
PopupMenu mnuStaticCmds, , CMDFrm.Left + StaticButton.Left, CMDFrm.Top + StaticButton.Height + StaticButton.Top
End Sub

Public Function AddStatic(ByVal CmdStr As String)
'The most used sub in the program
'Originally made to add static commands i found
'it could be used everywhre.
'
'Adds the cmdstr at the current position (curpos) if
' selected by the user or adds it at the end
If CurPos = -1 Then
LB.AddItem CmdStr
Else
CurPos = CurPos + 1
LB.AddItem CmdStr, CurPos
LB.ListIndex = CurPos
End If

End Function

Private Sub Test_Click()
'Make a temporary batch file and runs it
'Deletes it after the user has finised

If LB.ListCount = 0 Then MsgBox "Nothing to test", , "BAT-Man": Exit Sub

Set ts = Fsys.CreateTextFile("Test.bat", True)
  Dim a%, count%
  count = LB.ListCount
  Do While a < count
  ts.WriteLine (LB.List(a))
  a = a + 1
  Loop
ts.Close
Shell "Test.bat", vbNormalFocus

MsgBox "Press OK after Testing", vbInformation, "BAT-Man"
'Delete file
If Fsys.FileExists("Test.bat") Then Fsys.DeleteFile "Test.bat", True

End Sub

Private Sub txtBoxWidth_Change()
'Check if number entered by user is valid
'if it is then check if it is in the range 15-79
'If not show error msg
If txtBoxWidth.Text = "" Then SetLen (False): Exit Sub

On Error GoTo er:

Dim tem As Byte
tem = Int(txtBoxWidth.Text)
If tem < 14 Then SetLen (False): Exit Sub
If tem < 80 Then BoxWidth = tem: SetLen: Exit Sub

er:
SetLen (False)
MsgBox "Enter a valid number", vbInformation, "Invalid Number"
txtBoxWidth.SetFocus
txtBoxWidth.SelStart = 0
txtBoxWidth.SelLength = Len(txtBoxWidth.Text)

End Sub



Private Sub txtBoxWidth_LostFocus()
'does not let the user to do anything unless he (She??)
'enters a valid number
If txtBoxWidth.Text = "" Then SetLen (False): Exit Sub

On Error GoTo er:

Dim tem As Byte
tem = Int(txtBoxWidth.Text)

If tem > 14 And tem < 80 Then BoxWidth = tem: SetLen: Exit Sub

er:
SetLen (False)
MsgBox "Enter a valid number", vbInformation, "Invalid Number"
txtBoxWidth.SetFocus
txtBoxWidth.SelStart = 0
txtBoxWidth.SelLength = Len(txtBoxWidth.Text)

End Sub

Private Sub txtCustLine_Change()

'Shows error on wrong entry

On Error GoTo ee:
a = CInt(txtCustLine.Text)
If a > 0 And a < 256 Then Exit Sub

ee:
MsgBox "Enter valid number between 1-255", vbInformation, "Invalid entry"
txtCustLine.SetFocus
txtCustLine.SelStart = 0
txtCustLine.SelLength = Len(txtCustLine.Text)

End Sub

Private Sub txtCustLine_LostFocus()
'same as above
On Error GoTo ee:
a = CInt(txtCustLine.Text)
If a > 0 And a < 256 Then Exit Sub

ee:
MsgBox "Enter valid number between 1-255", vbInformation, "Invalid entry"
txtCustLine.SetFocus
txtCustLine.SelStart = 0
txtCustLine.SelLength = Len(txtCustLine.Text)

End Sub

Private Sub txtDes_Change()
'disable apply btn if text=""
If txtSrc <> "" And txtDes <> "" Then Apply.Enabled = True Else Apply.Enabled = False

End Sub


Private Sub txtLineMSG_KeyDown(KeyCode As Integer, Shift As Integer)
'Move the line into the list box if user hits enter
If KeyCode = 13 Then MsgDown.SimClick

End Sub


Private Sub txtSrc_Change()
'disable apply btn if text=""
If txtSrc <> "" And txtDes <> "" Then Apply.Enabled = True Else Apply.Enabled = False
End Sub

Private Sub Up_Click()
'Move the selected line up
'if no line selected then beep

If LB.ListIndex < 1 Then Beep: Exit Sub
Temp = LB.List(LB.ListIndex)
LB.List(LB.ListIndex) = LB.List(LB.ListIndex - 1)
LB.List(LB.ListIndex - 1) = Temp
CurPos = CurPos - 1
LB.ListIndex = CurPos
End Sub


Private Sub OptShow(Idx As Byte, Optional Extra As String = "", Optional Title As String = "")

'The second highest used sub in program (after addstatic)
'Show the mini-frame identified by idx.
'there are 5 mini-frames
'0-Attribute
'1-Copy/Rename/move/etc
'2-DIR
'3-Line
'4-Control statments

frmMain.Height = 6660
frmMsgBox.Enabled = False

For i = 0 To 4
If i = Idx Then
 FrmOpt(i).Visible = True
 'set curidx for use by OptClose
 curIdx = Idx
 FrmOpt(i).Top = 3940
 FrmOpt(i).Left = 7200
 FrmOpt(i).Tag = Extra
 If Title <> "" Then FrmOpt(i).Caption = Title
Else
 FrmOpt(i).Visible = False
End If
Next i

End Sub

Private Sub OptClose()
'Close the current shown mini-fram identified by curIdx
FrmOpt(curIdx).Visible = False
End Sub

Private Sub WriteBatchFile()
'Write code to -BatchFileName-
Set ts = Fsys.CreateTextFile(BatchFileName, True)
 If LB.ListCount = 0 Then ts.Close: Exit Sub
 Dim a%, count%
 count = LB.ListCount
  Do While a < count
  ts.WriteLine (LB.List(a))
  a = a + 1
  Loop
ts.Close
End Sub
Private Sub SetLen(Optional bool As Boolean = True)
'set the max-length of title & text in msgbox depending
'on box width
txtTitle.MaxLength = BoxWidth - 3
txtLineMSG.MaxLength = BoxWidth - 3
txtTitle.Enabled = bool
txtLineMSG.Enabled = bool
End Sub
