VERSION 5.00
Begin VB.Form Statistics 
   Caption         =   "Statistical Analysis"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   9480
   Icon            =   "Statistics.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Statistics"
   ScaleHeight     =   5355
   ScaleWidth      =   9480
   Begin VB.ComboBox TagName 
      Height          =   315
      Left            =   7200
      TabIndex        =   102
      Text            =   "TagName"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Button 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   7120
      Style           =   1  'Graphical
      TabIndex        =   98
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1360
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   3
      Left            =   7640
      Style           =   1  'Graphical
      TabIndex        =   97
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1360
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   5
      Left            =   7120
      Style           =   1  'Graphical
      TabIndex        =   96
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   920
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   6
      Left            =   7640
      Style           =   1  'Graphical
      TabIndex        =   95
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   920
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   8
      Left            =   7120
      Style           =   1  'Graphical
      TabIndex        =   94
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   480
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   9
      Left            =   7640
      Style           =   1  'Graphical
      TabIndex        =   93
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   480
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   10
      Left            =   7125
      Style           =   1  'Graphical
      TabIndex        =   92
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1800
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "&C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   11
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   91
      TabStop         =   0   'False
      Tag             =   "48"
      ToolTipText     =   "Clear All Data"
      Top             =   480
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "C&E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   12
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   90
      TabStop         =   0   'False
      Tag             =   "48"
      ToolTipText     =   "Cear Entry"
      Top             =   920
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "C&D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   13
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   89
      TabStop         =   0   'False
      Tag             =   "48"
      ToolTipText     =   "Clear Digit"
      Top             =   1360
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   15
      Left            =   7635
      Style           =   1  'Graphical
      TabIndex        =   88
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1800
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "C&P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   16
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   87
      TabStop         =   0   'False
      Tag             =   "48"
      ToolTipText     =   "Clear Previous Entry"
      Top             =   1800
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   7
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   86
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   480
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   4
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   85
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   920
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   84
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1360
      Width           =   400
   End
   Begin VB.CommandButton Button 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   83
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   1800
      Width           =   400
   End
   Begin VB.ComboBox DataName 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1665
   End
   Begin VB.ComboBox Categories 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1665
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   26
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   25
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   24
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   23
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   22
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   21
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   20
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   19
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   18
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   17
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   16
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   15
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   14
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4920
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   13
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   12
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   11
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   10
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   9
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   8
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   7
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   6
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   5
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   4
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   3
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   2
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.TextBox Value 
      Height          =   315
      Index           =   1
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4320
      Width           =   700
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Statistics.frx":0442
      Left            =   1605
      List            =   "Statistics.frx":0464
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   -15
      Width           =   615
   End
   Begin VB.CheckBox Option2 
      Caption         =   "Show all decimals"
      Height          =   225
      Left            =   2355
      TabIndex        =   26
      Top             =   45
      Width           =   1635
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   11
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2475
      Width           =   1710
   End
   Begin VB.CommandButton Button 
      Caption         =   "&SORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   2700
      Width           =   1980
   End
   Begin VB.CommandButton Button 
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   2280
      Width           =   1980
   End
   Begin VB.CommandButton Button 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "48"
      Top             =   3120
      Width           =   1980
   End
   Begin VB.PictureBox Display 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000015&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   6135
      ScaleHeight     =   375
      ScaleWidth      =   3300
      TabIndex        =   0
      Top             =   30
      Width           =   3330
   End
   Begin VB.Frame Frame1 
      Caption         =   "Method"
      Height          =   945
      Index           =   0
      Left            =   4320
      TabIndex        =   18
      Top             =   2640
      Width           =   1620
      Begin VB.OptionButton Option1 
         Caption         =   "Sample"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Population"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   19
         Top             =   555
         Width           =   1185
      End
   End
   Begin VB.ListBox ArrayList 
      Height          =   2595
      ItemData        =   "Statistics.frx":0487
      Left            =   4320
      List            =   "Statistics.frx":0489
      TabIndex        =   1
      Top             =   0
      Width           =   1650
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   10
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   675
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   9
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   315
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   8
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2115
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   7
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3210
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   6
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   5
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   4
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1710
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Index           =   3
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1710
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Tag"
      Height          =   255
      Left            =   6000
      TabIndex        =   101
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "Time"
      Height          =   255
      Left            =   2760
      TabIndex        =   100
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Caption         =   "Day"
      Height          =   255
      Left            =   0
      TabIndex        =   99
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label FieldLabel 
      Caption         =   "H"
      Height          =   195
      Index           =   8
      Left            =   5280
      TabIndex        =   82
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "A"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   81
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "B"
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   80
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "C"
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   79
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "D"
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   78
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "E"
      Height          =   195
      Index           =   5
      Left            =   3120
      TabIndex        =   77
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "F"
      Height          =   195
      Index           =   6
      Left            =   3840
      TabIndex        =   76
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "G"
      Height          =   195
      Index           =   7
      Left            =   4560
      TabIndex        =   75
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "I"
      Height          =   195
      Index           =   9
      Left            =   6000
      TabIndex        =   74
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "J"
      Height          =   195
      Index           =   10
      Left            =   6720
      TabIndex        =   73
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "K"
      Height          =   195
      Index           =   11
      Left            =   7440
      TabIndex        =   72
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "L"
      Height          =   195
      Index           =   12
      Left            =   8160
      TabIndex        =   71
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "M"
      Height          =   195
      Index           =   13
      Left            =   8880
      TabIndex        =   70
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "N"
      Height          =   195
      Index           =   14
      Left            =   240
      TabIndex        =   69
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "O"
      Height          =   195
      Index           =   15
      Left            =   960
      TabIndex        =   68
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "P"
      Height          =   195
      Index           =   16
      Left            =   1680
      TabIndex        =   67
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "Q"
      Height          =   195
      Index           =   17
      Left            =   2400
      TabIndex        =   66
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "R"
      Height          =   195
      Index           =   18
      Left            =   3120
      TabIndex        =   65
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "S"
      Height          =   195
      Index           =   19
      Left            =   3840
      TabIndex        =   64
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "T"
      Height          =   195
      Index           =   20
      Left            =   4560
      TabIndex        =   63
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "U"
      Height          =   195
      Index           =   21
      Left            =   5280
      TabIndex        =   62
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "V"
      Height          =   195
      Index           =   22
      Left            =   6000
      TabIndex        =   61
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "W"
      Height          =   195
      Index           =   23
      Left            =   6720
      TabIndex        =   60
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "X"
      Height          =   195
      Index           =   24
      Left            =   7440
      TabIndex        =   59
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "Y"
      Height          =   195
      Index           =   25
      Left            =   8160
      TabIndex        =   58
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label FieldLabel 
      Caption         =   "Z"
      Height          =   195
      Index           =   26
      Left            =   8880
      TabIndex        =   57
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label Label1 
      Caption         =   "Decimal places"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   420
      TabIndex        =   28
      Top             =   30
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "The skew is:"
      Height          =   195
      Index           =   12
      Left            =   90
      TabIndex        =   24
      Top             =   2175
      Width           =   2025
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   11
      Left            =   9030
      Picture         =   "Statistics.frx":048B
      Stretch         =   -1  'True
      Top             =   300
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   0
      Left            =   6285
      Picture         =   "Statistics.frx":0AA9
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   1
      Left            =   6540
      Picture         =   "Statistics.frx":0F63
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   2
      Left            =   6765
      Picture         =   "Statistics.frx":13C5
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   3
      Left            =   7020
      Picture         =   "Statistics.frx":187F
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   4
      Left            =   7275
      Picture         =   "Statistics.frx":1D39
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   5
      Left            =   7515
      Picture         =   "Statistics.frx":21F3
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   6
      Left            =   7770
      Picture         =   "Statistics.frx":26AD
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   7
      Left            =   8025
      Picture         =   "Statistics.frx":2B67
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   8
      Left            =   8280
      Picture         =   "Statistics.frx":3021
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   9
      Left            =   8535
      Picture         =   "Statistics.frx":34DB
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Digit 
      Height          =   330
      Index           =   10
      Left            =   8790
      Picture         =   "Statistics.frx":3995
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "The mode is:"
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   16
      Top             =   705
      Width           =   1920
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   75
      X2              =   3945
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   75
      X2              =   3975
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      Caption         =   "The number of entries is:"
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   390
      Width           =   2025
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   75
      X2              =   3930
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   75
      X2              =   3930
      Y1              =   1695
      Y2              =   1695
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   75
      X2              =   3975
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   75
      X2              =   3945
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   75
      X2              =   3930
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   75
      X2              =   3945
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Label Label1 
      Caption         =   "The average deviation is:"
      Height          =   255
      Index           =   9
      Left            =   105
      TabIndex        =   7
      Top             =   2520
      Width           =   2130
   End
   Begin VB.Label Label1 
      Caption         =   "The coefficient of variation is:"
      Height          =   255
      Index           =   8
      Left            =   90
      TabIndex        =   6
      Top             =   3255
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "The standard deviation is:"
      Height          =   255
      Index           =   7
      Left            =   90
      TabIndex        =   5
      Top             =   2895
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "The variance is:"
      Height          =   255
      Index           =   6
      Left            =   90
      TabIndex        =   4
      Top             =   1815
      Width           =   2115
   End
   Begin VB.Label Label1 
      Caption         =   "The mean is:"
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   3
      Top             =   1455
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "The median is:"
      Height          =   255
      Index           =   5
      Left            =   90
      TabIndex        =   2
      Top             =   1095
      Width           =   2025
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuSort 
         Caption         =   "&Sort"
      End
      Begin VB.Menu menuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "Edi&t"
      Begin VB.Menu menuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu menuCE 
         Caption         =   "Clear &Entry"
      End
      Begin VB.Menu menuDigit 
         Caption         =   "Clear &Digit"
      End
      Begin VB.Menu menuCP 
         Caption         =   "Clear &Previous"
      End
   End
   Begin VB.Menu menuData 
      Caption         =   "D&ata"
      Begin VB.Menu menuRecord 
         Caption         =   "Add &Record"
      End
      Begin VB.Menu menuTable 
         Caption         =   "&Add Table"
      End
      Begin VB.Menu menuDelRecord 
         Caption         =   "&Delete Record"
      End
      Begin VB.Menu menuDelTable 
         Caption         =   "Delete &Table"
      End
   End
   Begin VB.Menu menuTags 
      Caption         =   "&Tags"
      Begin VB.Menu menuInTag 
         Caption         =   "&Insert"
      End
      Begin VB.Menu menuOutTag 
         Caption         =   "&Remove"
      End
   End
   Begin VB.Menu menuExcel 
      Caption         =   "Exce&l"
      Begin VB.Menu menuSend 
         Caption         =   "&Send Table"
      End
   End
End
Attribute VB_Name = "Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DigitArray() As Integer
Dim Digits As Integer
Dim Sample As Boolean
Dim Calculator As Boolean
Dim DecimalPlaced As Boolean
Dim Rounding As Boolean
Dim Places As Integer
Dim myExcelFile As New ExcelFile
Dim TheIndex As Integer
Dim TheArray() As Variant
Dim SC As ExlCell
Dim LastValue As String

Private Type ExlCell
nRow As Long
nCol As Long
End Type

'tag database
Dim TDB As Database
Dim TRS As Recordset
Dim TTbl As TableDef
Dim TFld As Field

Private Sub LoadFields()
If CurrentCategory = "" Then Exit Sub
If DataName = "" Then Exit Sub
Set RS = DB.OpenRecordset("SELECT * " _
  & "FROM " & Categories & " ORDER BY Description")
With RS
If .RecordCount <> 0 Then
.Move DataName.ListIndex
    If .Fields("Description") = DataName Then
    CurrentRecord = DataName
        For i = 1 To 26
            If VarType(.Fields(CStr(i))) <> vbNull Then
            Value(i) = .Fields(CStr(i))
            End If
        Next
        If VarType(.Fields("Tag")) <> vbNull Then
            If Not ShowListItem(.Fields("Tag"), TagName) Then
            TagName.AddItem .Fields("Tag")
            End If
        Else
            TagName.ListIndex = 0
        End If
    End If
End If
.Close
End With
End Sub

Private Sub LoadTags()

On Error GoTo Out:

If IsFile(App.Path & "\Tag.MDB") Then GoTo SkipCreation
Set TDB = DBEngine.Workspaces(0).CreateDatabase(App.Path + "\Tag.MDB", dbLangGeneral)
SkipCreation:
Set TDB = OpenDatabase(App.Path & "\Tag.MDB")
On Error GoTo 0
Set Find = New TableDef
    For Each Find In TDB.TableDefs
        If Find.Name = "Tags" Then
        Set Find = Nothing
        GoTo Skip:
        End If
    Next
Set TTbl = TDB.CreateTableDef("Tags")
    For i = 1 To 26
    Set TFld = TTbl.CreateField(CStr(i), dbText, 25)
    TFld.AllowZeroLength = True
    TTbl.Fields.Append TFld
    Next
TDB.TableDefs.Append TTbl
'fill with default values
Set TRS = TDB.OpenRecordset("Tags")
With TRS
.AddNew
.Fields("1") = ""
.Fields("2") = "WBC"
.Fields("3") = "Hgb"
.Fields("4") = "Hct"
.Fields("5") = "Rbc"
.Fields("6") = "Plt"
.Fields("7") = "Lithium"
.Update
.Close
End With
Skip:
Set TRS = TDB.OpenRecordset("SELECT * " _
  & "FROM Tags")
With TRS
If .RecordCount <> 0 Then
    For i = 1 To 26
        If VarType(.Fields(CStr(i))) <> vbNull Then
        TagName.AddItem .Fields(CStr(i))
        End If
    Next
End If
.Close
End With

TDB.Close
Exit Sub
Out:
MsgBox "Unable to open or create the database."
End Sub
Private Sub SaveTag()

If Counter > 26 Then Exit Sub
If DataName = "" Then Exit Sub

If Categories = "" Then Exit Sub

Set RS = DB.OpenRecordset("SELECT * FROM " _
& Categories & " WHERE Description = '" & DataName & "'")
With RS
.Edit
    If TagName <> "" Then
    .Fields("Tag") = TagName
    Else
    .Fields("Tag") = Null
    End If
.Update
.Close
End With
End Sub

Private Sub ShowCurrentRecord()

For i = 0 To DataName.ListCount - 1
If CurrentRecord = DataName.List(i) Then
DataName.ListIndex = i
Exit Sub
End If
Next

End Sub

Private Sub ShowCurrentTable()
If CurrentCategory = "" Then CurrentCategory = CurrentTable
For i = 0 To Categories.ListCount - 1
If CurrentCategory = Categories.List(i) Then
Categories.ListIndex = i
Exit Sub
End If
Next
Categories.ListIndex = Categories.ListCount - 1
End Sub

Private Function ShowListItem(Item As String, TheBox As ComboBox) As Boolean
Dim i As Integer
For i = 1 To TheBox.ListCount - 1
    If TheBox.List(i) = Item Then
    TheBox.ListIndex = i
    ShowListItem = True
    Exit Function
    End If
Next
ShowListItem = False
End Function

Private Sub ShowSort()

If Sorted Then
SortListbox
Else
ListUnsorted
End If

End Sub

Private Function SqrTotal() As Double

Dim i As Long

For i = 1 To Counter
SqrTotal = (TheArray(i) * TheArray(i)) + SqrTotal
Next
End Function

Private Sub LoadArray()

Counter = 0
Erase TheArray

For i = 1 To 26
ReDim Preserve TheArray(i)
    If Value(i) <> vbNullString Then
    TheArray(i) = Value(i)
    Counter = Counter + 1
    Else
    Exit For
    End If
Next

End Sub
Private Function FindInList(CB As ComboBox, TheItem As String) As Integer
FindInList = SendMessageStr(CB.hwnd, CB_FINDSTRING, -1, TheItem)
End Function






Private Sub AddTable(TableName As String)

Dim RecordName As String

If TableName = "" Then Exit Sub
Set Find = New TableDef
   For Each Find In DB.TableDefs
      If Find.Name = TableName Then
      MsgBox "A category by that name already exists."
         For i = 0 To Categories.ListCount - 1
         Categories.ListIndex = i
            If TableName = Categories Then
            Set Find = Nothing
            Exit For
            End If
         Next
      Exit Sub
      End If
   Next
Set Tbl = DB.CreateTableDef(TableName)
Set Fld = Tbl.CreateField("Description", dbText, 50)
Tbl.Fields.Append Fld
Set Fld = Tbl.CreateField("Tag", dbText, 50)
Tbl.Fields.Append Fld
    For i = 1 To 26
    Set Fld = Tbl.CreateField(CStr(i), dbSingle, 8)
    Tbl.Fields.Append Fld
    Next
DB.TableDefs.Append Tbl

LoadTables
CurrentCategory = TableName

RecordName = Format(Time, "h:mm:ss")
BeginRecord CurrentTable, RecordName
CurrentRecord = RecordName
ShowCurrentRecord
ShowCurrentTable
End Sub



Private Sub DeleteTable(TableName As String)
Dim Action As VbMsgBoxResult
If Categories = "" Then Exit Sub
Action = MsgBox("Do you really want to delete the category " & """" & TableName & """" & " and all the data it contains.", vbYesNo, "DELETE TABLE")
If Action = vbYes Then
DB.TableDefs.Delete Categories
CurrentCategory = ""
Categories.Clear
Button_Click 11
LoadTables
End If
End Sub
Private Sub LoadTables()
Dim Temp As String

Categories.Clear
DataName.Clear
ClearBoxes

For i = 0 To DB.TableDefs.Count - 1
Temp = DB.TableDefs(i).Name
   If Left(Temp, 4) <> "MSys" Then
   Categories.AddItem Temp
   End If
Next


End Sub
Private Sub LoadRecords()
DataName.Clear
ClearBoxes

If Categories = "" Then Exit Sub


Set RS = DB.OpenRecordset("SELECT * " _
& "FROM " & Categories & " ORDER BY Description")

With RS
   If .RecordCount <> 0 Then
      For i = 1 To .RecordCount
      DataName.AddItem .Fields("Description")
      .MoveNext
      If .EOF Then Exit For
      Next
   End If
.Close
End With

End Sub
Private Sub ClearBoxes()
For i = 1 To 26
Value(i) = ""
Next
For i = 3 To 11
txtOutput(i) = ""
Next
ArrayList.Clear
TagName.Clear
LoadTags
End Sub











Private Sub WriteAll(Value As Variant, ByVal nRow As Long, ByVal nCol As Long)
With myExcelFile
Select Case VarType(Value)
Case 0 'empty
.WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, nRow, nCol, ""
Case 1 'null
.WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, nRow, nCol, ""
Case 2 'integer
.WriteValue xlsInteger, xlsFont0, xlsRightAlign, xlsNormal, nRow, nCol, Value
Case 3, 4, 5, 6 'long, single, double
.WriteValue xlsNumber, xlsFont0, xlsRightAlign, xlsNormal, nRow, nCol, Value
Case 7 'date
.WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, nRow, nCol, Value
Case 8 'string
.WriteValue xlsText, xlsFont0, xlsLeftAlign, xlsNormal, nRow, nCol, Value
Case 12 'variant
.WriteValue xlsText, xlsFont0, xlsCenterAlign, xlsNormal, nRow, nCol, Value
Case 14 'decimal
.WriteValue xlsNumber, xlsFont0, xlsRightAlign, xlsNormal, nRow, nCol, Value
Case Else
'9=object, 10=error, 11=boolean, 17=byte, 36=userdef, 8192=array
End Select
End With
End Sub

Private Sub WriteExcelFile()
On Error GoTo FileError
Dim FileName As String
Dim Row As Byte
Dim Col As Byte
Row = 3
Col = 3
With myExcelFile
FileName = App.Path & "\Temp.xls"
.CreateFile FileName
.PrintGridLines = False
.SetMargin xlsTopMargin, 1
.SetMargin xlsLeftMargin, 1
.SetMargin xlsRightMargin, 1
.SetMargin xlsBottomMargin, 1
.SetFont "Arial", 10, xlsNoFormat
.SetFont "Arial", 10, xlsBold
.SetFont "Arial", 10, xlsBold + xlsUnderline
.SetFont "Courier", 12, xlsItalic

    'Description
    .SetColumnWidth 1, 1, 20
    'Tag
    .SetColumnWidth 2, 2, 8
    For Col = 3 To 100
    'Values
    .SetColumnWidth Col, Col, 8
    Next
Set RS = DB.OpenRecordset("SELECT * " _
  & "FROM " & Categories & " ORDER BY Description")
If RS.RecordCount > 0 Then
Row = 1
WriteAll Categories, 1, 1
    For Row = 4 To (RS.RecordCount + 2 * 18) Step 18
    Col = 3
        WriteAll RS.Fields("Description"), Row - 1, 2
        WriteAll RS.Fields("Tag"), Row - 1, 1
        For i = 1 To 26
            If VarType(RS.Fields(CStr(i))) <> vbNull Then
            WriteAll RS.Fields(CStr(i)), Row, Col
            Else
            Exit For
            End If
        Col = Col + 1
        Next
    RS.MoveNext
    If RS.EOF Then Exit For
    Next
    RS.Close
    Set RS = Nothing
 End If
.CloseFile
End With
Exit Sub
FileError:
MsgBox "Error writing to Excel, " & Err.Description
End Sub

Private Sub Categories_Click()
CurrentCategory = Categories
CurrentRecord = ""
Button_Click 11
LoadRecords
End Sub


Private Sub DataName_KeyDown(KeyCode As Integer, Shift As Integer)
Dim BoxwHND As Long
Dim R As Long
    If KeyCode = 13 Then
        Const WM_USER = &H400
        Const CB_SHOWDROPDOWN = WM_USER + 15
        DataName.SetFocus
        BoxwHND = GetFocus()
        R = SendMessage(BoxwHND, CB_SHOWDROPDOWN, 0, 0)
        KeyCode = 0
    End If
End Sub

Private Sub DropList()
SendMessageStr Categories.hwnd, CB_SHOWDROPDOWN, True, 0&
End Sub

Private Sub Form_Paint()
Form_Loaded = True
End Sub

Private Sub menuData_Click()
menuDelRecord.Enabled = Trim(DataName) <> ""
menuDelTable.Enabled = Categories <> ""
End Sub

Private Sub menuDelRecord_Click()

Dim Action As VbMsgBoxResult

Action = MsgBox("Delete the record: """ & DataName & """ " & _
" and all its data?", vbYesNo, "DELETE " & DataName & "?")
If Action = vbYes Then

Set RS = DB.OpenRecordset("SELECT * " _
& "FROM " & Categories & " ORDER BY Description")

With RS
If .EOF Then Exit Sub
.Move DataName.ListIndex
    If .Fields("Description") = CurrentRecord Then
    CurrentRecord = ""
    End If
.Delete
.Close
End With
Digits = 0
Counter = 0
Erase DigitArray
Erase TheArray
Display.Cls
DecimalPlaced = False
ClearBoxes
LoadRecords
LoadFields
End If
End Sub




Private Sub menuDelTable_Click()
DeleteTable Categories
End Sub







Private Sub DataName_Click()
CurrentRecord = DataName
LoadFields
LoadArray
ShowResults
ShowSort
End Sub
Public Function IsFile(FileString As String) As Boolean
Dim FileNumber As Integer
On Error Resume Next
FileNumber = FreeFile()
Open FileString For Input As #FileNumber
    If Err Then
    IsFile = False
    Exit Function
    End If
IsFile = True
Close #FileNumber
End Function

Private Sub menuExcel_Click()
menuSend.Enabled = Categories <> "" And DataName <> ""
End Sub

Private Sub menuInTag_Click()
Dim TheTag As String
TheTag = InputBox("Enter new tag for default list.", "NEW TAG")
If TheTag = "" Then Exit Sub
Set TDB = OpenDatabase(App.Path & "\Tag.MDB")
'Find next open spot
Set TRS = TDB.OpenRecordset("SELECT * " _
  & "FROM Tags")
With TRS
If .RecordCount <> 0 Then
    For i = 2 To 26 'leave out first one, we know it's Null
        If VarType(.Fields(CStr(i))) = vbNull Then
        Exit For
        End If
    Next
End If
.Close
End With

Set TRS = TDB.OpenRecordset("SELECT * " _
  & "FROM Tags")
With TRS
.Edit
.Fields(CStr(i)) = TheTag
.Update
.Close
End With
TDB.Close
TagName.Clear
LoadTags
End Sub

Private Sub menuOutTag_Click()
Dim Action As VbMsgBoxResult
Action = MsgBox("Delete the tag """ & TagName & """?", vbOKCancel, "DELETE CURRENT TAG")
If Action = vbOK Then
Set TDB = OpenDatabase(App.Path & "\Tag.MDB")
'See if tag is there
Set TRS = TDB.OpenRecordset("SELECT * " _
  & "FROM Tags")
With TRS
    If .RecordCount <> 0 Then
        For i = 2 To 26 'leave out first one, we know it's Null
            If VarType(.Fields(CStr(i))) <> vbNull Then
                If .Fields(CStr(i)) = TagName Then
                .Edit
                .Fields(CStr(i)) = Null
                .Update
                Exit For
                End If
            End If
        If i = 26 Then MsgBox "Tag not found."
        Next
    Else
    MsgBox "No tags found."
    End If
    .Close
    End With
    TDB.Close
    TagName.Clear
    LoadTags
End If
End Sub


Private Sub menuRecord_Click()
Dim RecordName As String
RecordName = Format(Time, "h:mm:ss")

If CurrentCategory = "" Then
BeginRecord CurrentTable, RecordName
Else
BeginRecord CurrentCategory, RecordName
End If
CurrentRecord = RecordName
LoadTables
'ShowCurrentTable
ShowCurrentRecord
End Sub




Private Sub menuSend_Click()
WriteExcelFile
ParseAndShell "Excel """"" & App.Path & "\temp.xls"""""
End Sub

Private Sub menuTable_Click()
Dim TableName As String
TableName = InputBox("Table Name.", "Data Needed...")
If TableName <> "" Then
AddTable TableName
End If
End Sub

Private Sub menuTags_Click()
menuOutTag.Enabled = TagName.ListIndex > 0
End Sub

Private Sub TagName_Click()
If CurrentCategory = "" Or CurrentRecord = "" Then Exit Sub
SaveTag
End Sub


Private Sub Value_LostFocus(Index As Integer)
'If Not IsNumeric(Value(Index)) Then
'Value(Index) = ""
'End If
End Sub

Private Sub CopyRecords()
'alternate method of automating Excel, not used
Dim i As Single
Dim Recs As Integer
Dim nRow As Long
Dim nCol As Long
'Dim Fd As Field
'If RS.EOF And RS.BOF Then Exit Sub
'RS.MoveLast
'ReDim TheArray(RS.RecordCount + 1, RS.Fields.Count)
nCol = 0
    'For Each Fd In RS.Fields
    'TheArray(0, nCol) = Fd.Name
    nCol = nCol + 1
    'Next
   
'RS.MoveFirst
'Recs = RS.RecordCount
'For nRow = 1 To RS.RecordCount
    'For nCol = 0 To RS.Fields.Count - 1
    'TheArray(nRow, nCol) = RS.Fields(nCol).Value
        If IsNull(TheArray(nRow, nCol)) Then
        TheArray(nRow, nCol) = ""
        End If
    'Next
'RS.MoveNext
'Next
'Sht.Range(Sht.Cells(SC.nRow, SC.nCol), _
  'Sht.Cells(SC.nRow + RS.RecordCount + 1, _
  'SC.nCol + RS.Fields.Count)).Value = TheArray
End Sub



Private Sub AddToArray()

Dim Temp As String
If Digits = 0 Then Exit Sub
For i = 1 To Digits
Select Case DigitArray(i)
Case 0 To 9
Temp = Temp & CStr(DigitArray(i))
Case 10
Temp = Temp & "."
Case 11
Temp = "-"
Case Else
Exit Sub
End Select
Next
Counter = Counter + 1
ReDim Preserve TheArray(Counter)
TheArray(Counter) = Val(Temp)
ShowResults
    If Sorted Then
    SortListbox
    Else
    ListUnsorted
    End If
End Sub

Private Sub SaveToData()
Dim RecordName As String
Dim TableName As String

If Counter > 26 Then
MsgBox "Cannot save more than 26 entries."
Exit Sub
End If

If Categories = "" Then
TableName = CurrentTable
Else
TableName = CurrentCategory
End If

If CurrentRecord = "" Then

RecordName = Format(Time, "h:mm:ss")
BeginRecord TableName, RecordName
CurrentRecord = RecordName
Else
RecordName = CurrentRecord
End If

Set RS = DB.OpenRecordset("SELECT * FROM " _
& TableName & " WHERE Description = '" & RecordName & "'")
With RS
.Edit
    For i = 1 To Counter
    .Fields(CStr(i)) = TheArray(i)
    Next
    For i = Counter + 1 To 26
    .Fields(CStr(i)) = Null
    Next
    If TagName <> "" Then
    .Fields("Tag") = TagName
    End If
.Update
.Close
End With

End Sub
Private Function AveDev()

Dim i As Long
Dim Temp As Double

For i = 1 To Counter
Temp = Temp + Abs((TheArray(i) - Mean))
Next

AveDev = Temp / CDbl(Counter)
End Function


Private Function CoefDev() As Variant
On Error GoTo Err1
'coefficient of deviation in percent form
CoefDev = (StdDev / Mean * 100)
Exit Function
Err1:
CoefDev = "#Div/0!"
End Function

Private Function Mode() As Variant
'the most repeated value
On Error GoTo Err1:
Dim i As Long
Dim j As Long
Dim Temp As Variant
Dim Element As Long
ReDim n(Counter) As Long
'load number of repetitions into an array n()
For i = 1 To Counter
For j = 1 To Counter
If TheArray(i) = TheArray(j) Then
n(i) = n(i) + 1
End If
Next
Next
'compare elements of the repetition counting array
j = n(1)
Element = 1
For i = 2 To Counter
If n(i) > j Then
Element = i 'this element has higher value
j = n(i) 'update for next rep
End If
Next
'get results
If Element = 1 And n(1) = 1 Then 'no repetitions
Mode = "None"
Exit Function
End If
'look for the highest tying values
For i = 1 To Counter
If Element <> i Then 'skip same one
If n(Element) = n(i) Then 'if it is a match
If TheArray(Element) <> TheArray(i) Then 'but not same value
If InStr(1, Temp, TheArray(i)) = 0 Then 'if not already listed
If Temp = "" Then 'put in the first one
Temp = TheArray(Element)
End If
Temp = Temp & " or " & TheArray(i) 'add the matching reps
End If
End If
End If
End If
Next

'if no ties found show the highest repeated value
If Temp = "" Then
Temp = TheArray(Element)
End If

Mode = Temp
Exit Function
Err1:
Mode = "Error"
End Function

Private Function Mean() As Double
'average of all values
Mean = Total / CDbl(Counter)
End Function

Private Function Total() As Double

Dim i As Long

For i = 1 To Counter
Total = TheArray(i) + Total
Next

End Function


Public Sub ShowResults()

If Counter = 0 Then Exit Sub
Dim Temp As String
Dim Number As Double

SortArray

txtOutput(3) = Str(Median)

Temp = Str(Mean)
Temp = Format(Temp, "####.0000000000")
    If Not Rounding Then
    Number = CDbl(Temp)
    Else
    Number = Round(CDbl(Temp), Places)
    End If
txtOutput(4) = Str(Number)

Temp = CStr(Variance)
Temp = Format(Temp, "####.0000000000")
If IsNumeric(Temp) Then
    If Not Rounding Then
    Number = CDbl(Temp)
    Else
    Number = Round(CDbl(Temp), Places)
    End If
txtOutput(5) = Str(Number)
Else
txtOutput(5) = Temp
End If

Temp = CStr(StdDev) 'call function
Temp = Format(Temp, "####.0000000000") 'remove any scientific notation
If IsNumeric(Temp) Then 'see if it's a number or error message
    If Not Rounding Then
    Number = CDbl(Temp) ' convert back to double to drop trailing zeros
    Else
    Number = Round(CDbl(Temp), Places)
    End If
txtOutput(6) = Str(Number) 'display as string
Else
txtOutput(6) = Temp 'display error message
End If

Temp = CStr(CoefDev)
Temp = Format(Temp, "####.0000000000")
If IsNumeric(Temp) Then
    If Not Rounding Then
    Number = CDbl(Temp)
    Else
    Number = Round(CDbl(Temp), Places)
    End If
txtOutput(7) = Str(Number) & " %"
Else
txtOutput(7) = Temp
End If

Temp = CStr(Skew)
Temp = Format(Temp, "####.0000000000")
If IsNumeric(Temp) Then
    If Not Rounding Then
    Number = CDbl(Temp)
    Else
    Number = Round(CDbl(Temp), Places)
    End If
txtOutput(8) = Str(Number)
Else
txtOutput(8) = Temp
End If

txtOutput(9) = Str(Counter)
txtOutput(10) = CStr(Mode)

Temp = Str(AveDev)
Temp = Format(Temp, "####.0000000000")
If IsNumeric(Temp) Then
    If Not Rounding Then
    Number = CDbl(Temp)
    Else
    Number = Round(CDbl(Temp), Places)
    End If
txtOutput(11) = Str(Number)
Else
txtOutput(11) = Temp
End If
    
End Sub


Private Function Median() As Double
'the middle value of array or average of two if even number
Select Case Counter Mod 2
Case 0
Median = Str((SortedArray(Counter / 2) + SortedArray((Counter / 2) + 1)) / 2)
Case 1
Median = Str(SortedArray((Counter + 1) / 2))
End Select
End Function


Private Function Skew() As Variant
On Error GoTo Err1:
Dim i As Long
Dim Temp As Double
For i = 1 To Counter
Temp = Temp + ((TheArray(i) - Mean) / StdDev) ^ 3
Next
Skew = (Counter / ((Counter - 1) * (Counter - 2))) * Temp
Exit Function
Err1:
Skew = "#Div/0!"
End Function

Private Sub SortArray() 'only to get median value
Dim Temp As Double
Dim j As Long
Dim i As Long
'first copy the array so we can still remove the
'last element of original array in the undo feature
ReDim SortedArray(Counter)
For i = 1 To Counter
SortedArray(i) = TheArray(i)
Next
'then loop through swapping values through temp
For i = 1 To Counter
For j = 1 To Counter
Temp = SortedArray(i)
If Temp < SortedArray(j) Then
SortedArray(i) = SortedArray(j)
SortedArray(j) = Temp
End If
Next
Next
End Sub





Private Function StdDev() As Variant
On Error GoTo Err1: 'standard deviation
StdDev = Sqr(Abs(Variance))
Exit Function
Err1:
StdDev = "#Div/0!"
End Function



Private Function Variance() As Variant
On Error GoTo Err1

Dim Sum As Double
Dim i As Long

For i = 1 To Counter 'summation of squares for think formula
Sum = Sum + ((TheArray(i) - Mean) * (TheArray(i) - Mean))
Next

If Sample Then 'sample method
   If Calculator Then 'for using hand calculator
   Variance = ((SqrTotal - ((Total * Total) / Counter))) / (Counter - 1)
   Else 'using "think" formula...better
   Variance = Sum / (Counter - 1)
   End If
Else 'population method
   If Calculator Then
   Variance = ((SqrTotal - ((Total * Total) / Counter))) / Counter
   Else
   Variance = Sum / Counter
   End If
End If
Exit Function
Err1:
Variance = "#Div/0!"
End Function



Private Sub ArrayList_Click()
SortListbox
End Sub

Public Sub Button_Click(Index As Integer)

If Not Form_Loaded Then Exit Sub

Select Case Index
Case 0 To 9 'numbers 0 through 9
Form_KeyPress (Index + 48)

Case 10 'decimal point
Form_KeyPress 46

Case 11 'clear all data
Digits = 0
Counter = 0
Erase DigitArray
Erase TheArray
Display.Cls
DecimalPlaced = False
ClearBoxes
DataName.Clear
CurrentRecord = ""
LoadRecords


Case 12 'clear entry
Digits = 0
Erase DigitArray
Display.Cls
DecimalPlaced = False

Case 13 'backspace key
Form_KeyPress 8

Case 14 'enter button
Form_KeyPress 13

Case 15 'minus sign
Form_KeyPress 45

Case 16 'clear previous value (and current entry)
If Counter > 0 Then
Digits = 0
Erase DigitArray
Display.Cls
Counter = Counter - 1
DecimalPlaced = False
    If Counter > 0 Then
    ShowResults
    Else
    ClearBoxes
    End If
    
    If Sorted Then 'refresh listbox
    SortListbox
    Else
    ListUnsorted
    End If

SaveToData
LoadRecords
ShowCurrentTable
ShowCurrentRecord
End If


Case 17
Unload Me

Case 18 'show list box (un)sorted
If Sorted Then 'refresh listbox
ListUnsorted
Else
SortListbox
End If
End Select

If Index <> 17 Then
Me.SetFocus
End If

End Sub

Private Sub Combo1_Click()
Places = CInt(Combo1.Text)
If Counter > 0 Then
ShowResults
End If
End Sub


Private Sub Digit_Click(Index As Integer)
'image controls containing number bitmaps
End Sub

Private Sub Display_Click()
' calculator picture box
End Sub


Public Sub Form_KeyPress(KeyAscii As Integer)

Dim TheRecord As String
Dim TheDigit As Integer

Select Case KeyAscii
Case 8 'backspace
If Digits > 0 Then
    If DigitArray(Digits) = 10 Then
    DecimalPlaced = False
    End If
Digits = Digits - 1
End If

Case 13 'enter key

If Counter = 26 Then
MsgBox "Cannot hold more than 26 entries."
Button_Click 12
Exit Sub
End If

AddToArray
Digits = 0
Erase DigitArray
DecimalPlaced = False
SaveToData
    If Categories = "" Then
    TheRecord = CurrentRecord
    LoadTables
    ShowCurrentTable
    CurrentRecord = TheRecord
    End If
LoadRecords
ShowCurrentRecord
LoadFields
ShowResults
ShowSort


Case 45 'minus sign
If Digits > 0 Then Exit Sub
Digits = Digits + 1
ReDim Preserve DigitArray(Digits) As Integer
DigitArray(Digits) = 11

Case 46 'decimal point
If DecimalPlaced Then Exit Sub
Digits = Digits + 1
ReDim Preserve DigitArray(Digits) As Integer
DigitArray(Digits) = 10
DecimalPlaced = True

Case 48 To 57 'numbers 0 through 9
Digits = Digits + 1
ReDim Preserve DigitArray(Digits) As Integer
DigitArray(Digits) = CInt(Chr(KeyAscii))

Case 27
Button_Click 12

Case 101
Button_Click 12

Case 99
Button_Click 11

Case 100
Button_Click 13

Case 112
Button_Click 16

Case 115
Button_Click 18

Case 120
End

Case Else

Exit Sub
End Select

Display.Cls

    For i = 1 To Digits
    TheDigit = DigitArray(Digits - i + 1)
    Display.PaintPicture Digit(TheDigit).Picture, (Display.Width - 100) - Digit(TheDigit).Width * i, 0
    Next

Display.Refresh
End Sub

Private Sub Form_Load()
Sample = True
Rounding = True
Places = 3
Combo1.Text = "3"
OpenDB
LoadTags
CurrentCategory = CurrentTable
LoadTables
LoadRecords
CurrentRecord = DataName
ShowCurrentTable
ShowCurrentRecord
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
DB.Close

Set Fld = Nothing
Set RS = Nothing
Set Tbl = Nothing
Set DB = Nothing

Set TDB = Nothing
Set TRS = Nothing
Set TTbl = Nothing
Set TFld = Nothing
Erase TheArray
Unload Me
End Sub
Private Sub SortListbox()

ArrayList.Clear
For i = 1 To Counter
ArrayList.AddItem Str(SortedArray(i))
Next
Button(18).Caption = "UN&SORT"
Sorted = True

End Sub
Private Sub ListUnsorted()

ArrayList.Clear
For i = 1 To Counter
ArrayList.AddItem Str(TheArray(i))
Next
Button(18).Caption = "&SORT"
Sorted = False
End Sub

Private Sub Frame1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
' the two frames to separate option button pairs
End Sub

Private Sub menuCE_Click()
Button_Click 12
End Sub

Private Sub menuClear_Click()
Button_Click 11
End Sub

Private Sub menuCP_Click()
Button_Click 16
End Sub

Private Sub menuExit_Click()
Button_Click 17
End Sub

Private Sub menuFile_Click()
If Button(18).Caption = "&SORT" Then
menuSort.Caption = "&Sort"
Else
menuSort.Caption = "Un&sort"
End If
End Sub



Private Sub menuSort_Click()
Button_Click 18
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
Sample = False
Case 1
Sample = True
End Select
If Counter > 0 Then ShowResults
End Sub


Private Sub Option2_Click()
If Option2 Then
Rounding = False
Combo1.Enabled = False
Combo1.Text = 10
Else
Rounding = True
Combo1.Enabled = True
Combo1.Text = 3
End If
If Counter > 0 Then ShowResults
End Sub


