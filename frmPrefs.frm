VERSION 5.00
Object = "{BCE37951-37DF-4D69-A8A3-2CFABEE7B3CC}#1.0#0"; "CCRSlider.ocx"
Begin VB.Form widgetPrefs 
   Caption         =   "TenShillings Preferences"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   8880
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10693.53
   ScaleMode       =   0  'User
   ScaleWidth      =   8880
   Visible         =   0   'False
   Begin VB.Frame fraTimers 
      Caption         =   "Timers"
      Height          =   2175
      Left            =   90
      TabIndex        =   177
      Top             =   6270
      Visible         =   0   'False
      Width           =   2385
      Begin VB.Timer themeTimer 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   60
         Tag             =   "a timer to apply a theme automatically"
         Top             =   240
      End
      Begin VB.Timer tmrWritePositionAndSize 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   60
         Top             =   1200
      End
      Begin VB.Timer tmrPrefsScreenResolution 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   60
         Top             =   720
      End
      Begin VB.Label tmrLabel 
         Caption         =   "tmrWritePositionAndSize"
         Height          =   225
         Index           =   3
         Left            =   570
         TabIndex        =   180
         Top             =   1290
         Width           =   1785
      End
      Begin VB.Label tmrLabel 
         Caption         =   "tmrPrefsScreenResolution"
         Height          =   435
         Index           =   2
         Left            =   570
         TabIndex        =   179
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label tmrLabel 
         Caption         =   "themeTimer"
         Height          =   435
         Index           =   1
         Left            =   570
         TabIndex        =   178
         Top             =   330
         Width           =   1245
      End
   End
   Begin VB.CheckBox chkEnableResizing 
      Caption         =   "Enable Corner Resize"
      Height          =   210
      Left            =   3240
      TabIndex        =   103
      Top             =   10125
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Save the changes you have made to the preferences"
      Top             =   10035
      Width           =   1320
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Open the help utility"
      Top             =   10035
      Width           =   1320
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close the utility"
      Top             =   10035
      Width           =   1320
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuration"
      Height          =   8655
      Left            =   225
      TabIndex        =   2
      Top             =   1200
      Width           =   7605
      Begin VB.Frame fraConfigInner 
         BorderStyle     =   0  'None
         Height          =   7965
         Left            =   450
         TabIndex        =   25
         Top             =   435
         Width           =   6705
         Begin VB.Frame fraClockTooltips 
            BorderStyle     =   0  'None
            Height          =   1110
            Left            =   1785
            TabIndex        =   135
            Top             =   3615
            Width           =   3345
            Begin VB.OptionButton optWidgetTooltips 
               Caption         =   "Disable Gauge Tooltips *"
               Height          =   300
               Index           =   2
               Left            =   225
               TabIndex        =   138
               Top             =   795
               Width           =   2790
            End
            Begin VB.OptionButton optWidgetTooltips 
               Caption         =   "Gauge - Enable Square Tooltips"
               Height          =   300
               Index           =   1
               Left            =   225
               TabIndex        =   137
               Top             =   465
               Width           =   2790
            End
            Begin VB.OptionButton optWidgetTooltips 
               Caption         =   "Gauge - Enable Balloon Tooltips *"
               Height          =   315
               Index           =   0
               Left            =   225
               TabIndex        =   136
               Top             =   135
               Width           =   3060
            End
         End
         Begin vb6projectCCRSlider.Slider sliWidgetSize 
            Height          =   390
            Left            =   1920
            TabIndex        =   132
            Top             =   -45
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   5
            Max             =   200
            Value           =   5
            TickFrequency   =   3
            SelStart        =   5
         End
         Begin VB.Frame fraPrefsTooltips 
            BorderStyle     =   0  'None
            Height          =   1125
            Index           =   0
            Left            =   1860
            TabIndex        =   122
            Top             =   4770
            Width           =   3150
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Disable Prefs Tooltips *"
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   134
               Top             =   780
               Width           =   2970
            End
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Prefs - Enable Balloon Tooltips *"
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   124
               Top             =   120
               Width           =   2760
            End
            Begin VB.OptionButton optPrefsTooltips 
               Caption         =   "Prefs - Enable SquareTooltips *"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   123
               Top             =   450
               Width           =   2970
            End
         End
         Begin VB.ComboBox cmbScrollWheelDirection 
            Height          =   315
            Left            =   1995
            Style           =   2  'Dropdown List
            TabIndex        =   66
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   1410
            Width           =   2490
         End
         Begin vb6projectCCRSlider.Slider sliSkewDegrees 
            Height          =   450
            Left            =   1890
            TabIndex        =   146
            Top             =   2415
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   794
            Max             =   360
            Value           =   5
            TickFrequency   =   5
            SelStart        =   5
         End
         Begin VB.Frame fraCheckBoxHolder 
            BorderStyle     =   0  'None
            Height          =   2115
            Left            =   285
            TabIndex        =   151
            Top             =   5820
            Width           =   6180
            Begin VB.CheckBox chkShowTaskbar 
               Caption         =   "Show Widget in Taskbar"
               Height          =   225
               Left            =   1725
               TabIndex        =   154
               ToolTipText     =   "Check the box to show the widget in the taskbar"
               Top             =   195
               Width           =   3405
            End
            Begin VB.CheckBox chkDpiAwareness 
               Caption         =   "DPI Awareness Enable *"
               Height          =   285
               Left            =   1725
               TabIndex        =   153
               ToolTipText     =   "Check the box to make the program DPI aware. RESTART required."
               Top             =   855
               Width           =   3405
            End
            Begin VB.CheckBox chkShowHelp 
               Caption         =   "Show Help on Widget Start"
               Height          =   225
               Left            =   1725
               TabIndex        =   152
               ToolTipText     =   "Check the box to show the widget in the taskbar"
               Top             =   540
               Width           =   3405
            End
            Begin VB.Label lblConfiguration 
               Caption         =   $"frmPrefs.frx":0CCA
               Height          =   855
               Index           =   0
               Left            =   1740
               TabIndex        =   155
               Top             =   1215
               Width           =   4335
            End
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Rotate the widget. You can also use Mousewheel. Immediate. *"
            Height          =   555
            Index           =   4
            Left            =   2025
            TabIndex        =   156
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   3210
            Width           =   4515
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "180"
            Height          =   315
            Index           =   8
            Left            =   3660
            TabIndex        =   150
            Top             =   2895
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "360"
            Height          =   315
            Index           =   7
            Left            =   5550
            TabIndex        =   149
            Top             =   2895
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "0"
            Height          =   315
            Index           =   6
            Left            =   1995
            TabIndex        =   148
            Top             =   2910
            Width           =   345
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Widget Rotation :"
            Height          =   375
            Index           =   5
            Left            =   600
            TabIndex        =   147
            Top             =   2490
            Width           =   1365
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "The scroll-wheel resizing direction can be determined here. The direction chosen causes the gauge to grow. *"
            Height          =   675
            Index           =   6
            Left            =   2025
            TabIndex        =   91
            Top             =   1830
            Width           =   3990
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "160"
            Height          =   315
            Index           =   4
            Left            =   4740
            TabIndex        =   70
            Top             =   435
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "120"
            Height          =   315
            Index           =   3
            Left            =   3990
            TabIndex        =   69
            Top             =   435
            Width           =   345
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "50"
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   68
            Top             =   435
            Width           =   345
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Mouse Wheel Resize :"
            Height          =   345
            Index           =   3
            Left            =   255
            TabIndex        =   67
            ToolTipText     =   "To change the direction of the mouse scroll wheel when resiziing the globe gauge."
            Top             =   1455
            Width           =   2055
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel. Immediate. *"
            Height          =   555
            Index           =   2
            Left            =   2070
            TabIndex        =   65
            ToolTipText     =   "Adjust to a percentage of the original size. You can also use Ctrl+Mousewheel."
            Top             =   780
            Width           =   3810
         End
         Begin VB.Label lblConfiguration 
            Caption         =   "Widget Size :"
            Height          =   315
            Index           =   1
            Left            =   885
            TabIndex        =   64
            Top             =   30
            Width           =   975
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "80"
            Height          =   315
            Index           =   2
            Left            =   3360
            TabIndex        =   63
            Top             =   435
            Width           =   360
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "200 (%)"
            Height          =   315
            Index           =   5
            Left            =   5385
            TabIndex        =   62
            Top             =   435
            Width           =   735
         End
         Begin VB.Label lblGaugeSize 
            Caption         =   "5"
            Height          =   315
            Index           =   0
            Left            =   2085
            TabIndex        =   61
            Top             =   435
            Width           =   345
         End
      End
   End
   Begin VB.Frame fraIconGroup 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   -10
      TabIndex        =   160
      Top             =   0
      Width           =   8895
      Begin VB.Frame fraAboutButton 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   7695
         TabIndex        =   175
         Top             =   60
         Width           =   975
         Begin VB.Image imgAbout 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":0D8A
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
         Begin VB.Label lblAbout 
            Caption         =   "About"
            Height          =   240
            Index           =   0
            Left            =   255
            TabIndex        =   176
            Top             =   855
            Width           =   615
         End
         Begin VB.Image imgAboutClicked 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":1312
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame fraWindowButton 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   6615
         TabIndex        =   173
         Top             =   60
         Width           =   975
         Begin VB.Image imgWindow 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":17FD
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
         Begin VB.Label lblWindow 
            Caption         =   "Window"
            Height          =   240
            Left            =   180
            TabIndex        =   174
            Top             =   855
            Width           =   615
         End
         Begin VB.Image imgWindowClicked 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":1CC7
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame fraDevelopmentButton 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   5490
         TabIndex        =   171
         Top             =   60
         Width           =   1065
         Begin VB.Image imgDevelopment 
            Height          =   600
            Left            =   150
            Picture         =   "frmPrefs.frx":2073
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
         Begin VB.Label lblDevelopment 
            Caption         =   "Development"
            Height          =   240
            Left            =   45
            TabIndex        =   172
            Top             =   855
            Width           =   960
         End
         Begin VB.Image imgDevelopmentClicked 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":262B
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame fraPositionButton 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   4410
         TabIndex        =   169
         Top             =   60
         Width           =   930
         Begin VB.Image imgPosition 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":29B1
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
         Begin VB.Label lblPosition 
            Caption         =   "Position"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   170
            Top             =   855
            Width           =   615
         End
         Begin VB.Image imgPositionClicked 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":2F82
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame fraSoundsButton 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   3315
         TabIndex        =   167
         Top             =   60
         Width           =   930
         Begin VB.Image imgSounds 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":3320
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
         Begin VB.Label lblSounds 
            Caption         =   "Sounds"
            Height          =   240
            Left            =   210
            TabIndex        =   168
            Top             =   870
            Width           =   615
         End
         Begin VB.Image imgSoundsClicked 
            Height          =   600
            Left            =   160
            Picture         =   "frmPrefs.frx":38DF
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame fraFontsButton 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   2280
         TabIndex        =   165
         Top             =   90
         Width           =   930
         Begin VB.Image imgFonts 
            Height          =   600
            Left            =   180
            Picture         =   "frmPrefs.frx":3DAF
            Stretch         =   -1  'True
            Top             =   195
            Width           =   600
         End
         Begin VB.Label lblFonts 
            Caption         =   "Fonts"
            Height          =   240
            Left            =   270
            TabIndex        =   166
            Top             =   855
            Width           =   510
         End
         Begin VB.Image imgFontsClicked 
            Height          =   600
            Left            =   180
            Picture         =   "frmPrefs.frx":4305
            Stretch         =   -1  'True
            Top             =   195
            Width           =   600
         End
      End
      Begin VB.Frame fraConfigButton 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   1215
         TabIndex        =   163
         Top             =   60
         Width           =   930
         Begin VB.Label lblConfig 
            Caption         =   "Config."
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   164
            Top             =   855
            Width           =   510
         End
         Begin VB.Image imgConfig 
            Height          =   600
            Left            =   165
            Picture         =   "frmPrefs.frx":479E
            Stretch         =   -1  'True
            Top             =   195
            Width           =   600
         End
         Begin VB.Image imgConfigClicked 
            Height          =   600
            Left            =   165
            Picture         =   "frmPrefs.frx":4D7D
            Stretch         =   -1  'True
            Top             =   195
            Width           =   600
         End
      End
      Begin VB.Frame fraGeneralButton 
         Height          =   1140
         Left            =   240
         TabIndex        =   161
         Top             =   90
         Width           =   930
         Begin VB.Image imgGeneral 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   165
            Picture         =   "frmPrefs.frx":5282
            Stretch         =   -1  'True
            Top             =   225
            Width           =   600
         End
         Begin VB.Image imgGeneralClicked 
            Height          =   600
            Left            =   165
            Stretch         =   -1  'True
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lblGeneral 
            Caption         =   "General"
            Height          =   240
            Index           =   0
            Left            =   195
            TabIndex        =   162
            Top             =   855
            Width           =   705
         End
      End
   End
   Begin VB.Frame fraWindow 
      Caption         =   "Window"
      Height          =   8460
      Left            =   210
      TabIndex        =   4
      Top             =   1215
      Width           =   8415
      Begin VB.Frame fraWindowInner 
         BorderStyle     =   0  'None
         Height          =   7995
         Left            =   165
         TabIndex        =   6
         Top             =   345
         Width           =   7470
         Begin VB.CheckBox chkFormVisible 
            Caption         =   "Form Visible"
            Height          =   225
            Left            =   2250
            TabIndex        =   158
            Top             =   1230
            Width           =   2535
         End
         Begin vb6projectCCRSlider.Slider sliOpacity 
            Height          =   390
            Left            =   2115
            TabIndex        =   133
            Top             =   5205
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   688
            Min             =   20
            Max             =   100
            Value           =   20
            SmallChange     =   2
            SelStart        =   20
         End
         Begin VB.ComboBox cmbMultiMonitorResize 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   121
            Top             =   6435
            Width           =   3720
         End
         Begin VB.Frame fraHiding 
            BorderStyle     =   0  'None
            Height          =   2010
            Left            =   1395
            TabIndex        =   92
            Top             =   2970
            Width           =   5130
            Begin VB.ComboBox cmbHidingTime 
               Height          =   315
               Left            =   825
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Top             =   1575
               Width           =   3720
            End
            Begin VB.CheckBox chkWidgetHidden 
               Caption         =   "Hiding Widget *"
               Height          =   225
               Left            =   855
               TabIndex        =   93
               Top             =   315
               Width           =   2955
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   "Hiding :"
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   96
               Top             =   315
               Width           =   720
            End
            Begin VB.Label lblWindowLevel 
               Caption         =   $"frmPrefs.frx":5E71
               Height          =   975
               Index           =   1
               Left            =   855
               TabIndex        =   94
               Top             =   705
               Width           =   3900
            End
         End
         Begin VB.ComboBox cmbWindowLevel 
            Height          =   315
            Left            =   2220
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   0
            Width           =   3720
         End
         Begin VB.CheckBox chkIgnoreMouse 
            Caption         =   "Ignore Mouse *"
            Height          =   225
            Left            =   2250
            TabIndex        =   7
            ToolTipText     =   "Checking this box causes the program to ignore all mouse events."
            Top             =   2070
            Width           =   2535
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Checking this box causes the underlying form to show"
            Height          =   420
            Index           =   12
            Left            =   2250
            TabIndex        =   159
            Top             =   1620
            Width           =   4260
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Multi-Monitor Resizing :"
            Height          =   255
            Index           =   11
            Left            =   375
            TabIndex        =   120
            Top             =   6465
            Width           =   1830
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   $"frmPrefs.frx":5F14
            Height          =   1140
            Index           =   10
            Left            =   2235
            TabIndex        =   119
            Top             =   6885
            Width           =   4050
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "This setting controls the relative layering of this widget. You may use it to place it on top of other windows or underneath. "
            Height          =   660
            Index           =   3
            Left            =   2235
            TabIndex        =   101
            Top             =   570
            Width           =   4380
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Window Level :"
            Height          =   345
            Index           =   0
            Left            =   915
            TabIndex        =   15
            Top             =   60
            Width           =   1740
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "20%"
            Height          =   315
            Index           =   7
            Left            =   2205
            TabIndex        =   14
            Top             =   5700
            Width           =   345
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "100%"
            Height          =   315
            Index           =   9
            Left            =   5565
            TabIndex        =   13
            Top             =   5700
            Width           =   405
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "60%"
            Height          =   315
            Index           =   8
            Left            =   3975
            TabIndex        =   12
            Top             =   5700
            Width           =   840
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Opacity:"
            Height          =   315
            Index           =   6
            Left            =   1470
            TabIndex        =   11
            Top             =   5250
            Width           =   780
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Set the program transparency level."
            Height          =   330
            Index           =   5
            Left            =   2250
            TabIndex        =   10
            Top             =   6015
            Width           =   3810
         End
         Begin VB.Label lblWindowLevel 
            Caption         =   "Checking this box causes the program to ignore all mouse events except right click menu interactions."
            Height          =   660
            Index           =   4
            Left            =   2235
            TabIndex        =   9
            Top             =   2460
            Width           =   3810
         End
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   240
      TabIndex        =   140
      Top             =   1200
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Frame fraGeneralInner 
         BorderStyle     =   0  'None
         Height          =   1620
         Left            =   480
         TabIndex        =   141
         Top             =   270
         Width           =   6750
         Begin VB.CheckBox chkGenStartup 
            Caption         =   "Run the Ten Shillings Widget at Windows Startup "
            Height          =   465
            Left            =   1950
            TabIndex        =   142
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   1125
            Width           =   4020
         End
         Begin VB.CheckBox chkWidgetFunctions 
            Caption         =   "Double Click Enabled *"
            Height          =   465
            Left            =   1950
            TabIndex        =   144
            ToolTipText     =   "Check this box to enable the automatic start of the program when Windows is started."
            Top             =   165
            Width           =   4020
         End
         Begin VB.Label lblGeneral 
            Caption         =   "When checked this box enables the double click functionality. That's it! *"
            Height          =   375
            Index           =   2
            Left            =   1965
            TabIndex        =   157
            Tag             =   "lblRefreshInterval"
            Top             =   660
            Width           =   3570
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Auto Start :"
            Height          =   375
            Index           =   11
            Left            =   960
            TabIndex        =   143
            Tag             =   "lblRefreshInterval"
            Top             =   1245
            Width           =   1740
         End
         Begin VB.Label lblGeneral 
            Caption         =   "Widget Functions :"
            Height          =   375
            Index           =   1
            Left            =   420
            TabIndex        =   145
            Tag             =   "lblRefreshInterval"
            Top             =   285
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraAbout 
      Caption         =   "About"
      Height          =   8580
      Left            =   240
      TabIndex        =   72
      Top             =   1155
      Visible         =   0   'False
      Width           =   8520
      Begin VB.CommandButton btnGithubHome 
         Caption         =   "Github &Home"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   300
         Width           =   1470
      End
      Begin VB.CommandButton btnDonate 
         Caption         =   "&Donate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Opens a browser window and sends you to our donate page on Amazon"
         Top             =   1800
         Width           =   1470
      End
      Begin VB.Frame fraScrollbarCover 
         BorderStyle     =   0  'None
         Height          =   6225
         Left            =   7980
         TabIndex        =   86
         Top             =   2205
         Width           =   420
      End
      Begin VB.TextBox txtAboutText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Text            =   "frmPrefs.frx":602B
         Top             =   2205
         Width           =   7935
      End
      Begin VB.CommandButton btnAboutDebugInfo 
         Caption         =   "Debug &Info."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "This gives access to the debugging tool"
         Top             =   1425
         Width           =   1470
      End
      Begin VB.CommandButton btnFacebook 
         Caption         =   "&Facebook"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "This will link you to the Rocket/Steamy dock users Group"
         Top             =   1050
         Width           =   1470
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6765
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Here you can visit the update location where you can download new versions of the programs."
         Top             =   675
         Width           =   1470
      End
      Begin VB.Label lblMajorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2445
         TabIndex        =   88
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblMinorVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2790
         TabIndex        =   87
         Top             =   510
         Width           =   225
      End
      Begin VB.Label lblAbout 
         Caption         =   "Dean Beedell © 2025"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2430
         TabIndex        =   84
         Top             =   855
         Width           =   2175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Originator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   885
         TabIndex        =   83
         Top             =   855
         Width           =   795
      End
      Begin VB.Label lblAbout 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   900
         TabIndex        =   82
         Top             =   495
         Width           =   795
      End
      Begin VB.Label lblAbout 
         Caption         =   "Dean Beedell © 2025"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   2430
         TabIndex        =   81
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label lblAbout 
         Caption         =   "Current Developer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   885
         TabIndex        =   80
         Top             =   1215
         Width           =   1470
      End
      Begin VB.Label lblAbout 
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   885
         TabIndex        =   79
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label lblAbout 
         Caption         =   "Windows XP, ReactOS, Vista, 7, 8, 10  && 11+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2430
         TabIndex        =   78
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label lblAbout 
         Caption         =   "(32bit WoW64)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3525
         TabIndex        =   77
         Top             =   510
         Width           =   3000
      End
      Begin VB.Label lblDotDot 
         BackStyle       =   0  'Transparent
         Caption         =   ".        ."
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2610
         TabIndex        =   90
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblRevisionNum 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Centurion Light SF"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3090
         TabIndex        =   89
         Top             =   510
         Width           =   525
      End
   End
   Begin VB.Frame fraSounds 
      Caption         =   "Sounds"
      Height          =   2280
      Left            =   855
      TabIndex        =   5
      Top             =   1230
      Visible         =   0   'False
      Width           =   7965
      Begin VB.Frame fraSoundsInner 
         BorderStyle     =   0  'None
         Height          =   1875
         Left            =   765
         TabIndex        =   16
         Top             =   285
         Width           =   6420
         Begin VB.CheckBox chkEnableSounds 
            Caption         =   "Enable ALL sounds for the whole widget."
            Height          =   225
            Left            =   1485
            TabIndex        =   26
            ToolTipText     =   "Check this box to enable or disable all of the sounds used during any animation on the main screen."
            Top             =   285
            Width           =   4485
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Determine the sound of UI elements, clock tick and chiming volumes. Set the overall volume to loud or quiet."
            Height          =   540
            Index           =   4
            Left            =   885
            TabIndex        =   118
            Tag             =   "lblSharedInputFile"
            Top             =   750
            Width           =   4680
         End
         Begin VB.Label lblSoundsTab 
            Caption         =   "Audio :"
            Height          =   255
            Index           =   3
            Left            =   885
            TabIndex        =   71
            Tag             =   "lblSharedInputFile"
            Top             =   285
            Width           =   765
         End
      End
   End
   Begin VB.Frame fraFonts 
      Caption         =   "Fonts"
      Height          =   5565
      Left            =   255
      TabIndex        =   3
      Top             =   1230
      Width           =   8280
      Begin VB.Frame fraFontsInner 
         BorderStyle     =   0  'None
         Height          =   5010
         Left            =   690
         TabIndex        =   17
         Top             =   360
         Width           =   6105
         Begin VB.TextBox txtDisplayScreenFont 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   127
            Text            =   "Courier  New"
            Top             =   3075
            Width           =   3285
         End
         Begin VB.CommandButton btnDisplayScreenFont 
            Caption         =   "Font"
            Height          =   300
            Left            =   5010
            Style           =   1  'Graphical
            TabIndex        =   126
            Top             =   3075
            Width           =   585
         End
         Begin VB.TextBox txtDisplayScreenFontSize 
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   125
            Text            =   "8"
            Top             =   3615
            Width           =   510
         End
         Begin VB.CommandButton btnResetMessages 
            Caption         =   "Reset"
            Height          =   300
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   116
            Top             =   4230
            Width           =   885
         End
         Begin VB.TextBox txtPrefsFontCurrentSize 
            Height          =   315
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   104
            ToolTipText     =   "Disabled for manual input. Shows the current font size when form resizing is enabled."
            Top             =   1065
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtPrefsFontSize 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "8"
            ToolTipText     =   "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
            Top             =   1065
            Width           =   510
         End
         Begin VB.CommandButton btnPrefsFont 
            Caption         =   "Font"
            Height          =   300
            Left            =   5025
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   90
            Width           =   585
         End
         Begin VB.TextBox txtPrefsFont 
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "Times New Roman"
            Top             =   90
            Width           =   3285
         End
         Begin VB.Label lblFontsTab 
            Caption         =   $"frmPrefs.frx":6FE2
            Height          =   1710
            Index           =   0
            Left            =   1680
            TabIndex        =   139
            ToolTipText     =   "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
            Top             =   1545
            Width           =   4455
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text on the main gauge face *"
            Height          =   480
            Index           =   9
            Left            =   2415
            TabIndex        =   130
            Top             =   3600
            Width           =   4035
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Gauge Font :"
            Height          =   300
            Index           =   8
            Left            =   585
            TabIndex        =   129
            Tag             =   "lblPrefsFont"
            Top             =   3105
            Width           =   1230
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Console  Font Size :"
            Height          =   330
            Index           =   5
            Left            =   165
            TabIndex        =   128
            Tag             =   "lblPrefsFontSize"
            Top             =   3645
            Width           =   1590
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Hidden message boxes can be reactivated by pressing this reset button."
            Height          =   480
            Index           =   4
            Left            =   2670
            TabIndex        =   117
            Top             =   4170
            Width           =   3360
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Reset Pop ups :"
            Height          =   300
            Index           =   1
            Left            =   405
            TabIndex        =   115
            Tag             =   "lblPrefsFont"
            Top             =   4275
            Width           =   1470
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Resized Font"
            Height          =   315
            Index           =   10
            Left            =   4920
            TabIndex        =   105
            Top             =   1110
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "The chosen font size *"
            Height          =   480
            Index           =   7
            Left            =   2310
            TabIndex        =   24
            Top             =   1095
            Width           =   2400
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Base Font Size :"
            Height          =   330
            Index           =   3
            Left            =   435
            TabIndex        =   23
            Tag             =   "lblPrefsFontSize"
            Top             =   1095
            Width           =   1230
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Prefs Utility Font :"
            Height          =   300
            Index           =   2
            Left            =   360
            TabIndex        =   22
            Tag             =   "lblPrefsFont"
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblFontsTab 
            Caption         =   "Choose a font to be used for the text in this preferences window, gauge tooltips and message boxes *"
            Height          =   480
            Index           =   6
            Left            =   1695
            TabIndex        =   21
            Top             =   480
            Width           =   4035
         End
      End
   End
   Begin VB.Frame fraDevelopment 
      Caption         =   "Development"
      Height          =   6210
      Left            =   240
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraDevelopmentInner 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   870
         TabIndex        =   31
         Top             =   300
         Width           =   7455
         Begin VB.Frame fraDefaultEditor 
            BorderStyle     =   0  'None
            Height          =   2370
            Left            =   75
            TabIndex        =   106
            Top             =   3165
            Width           =   7290
            Begin VB.CommandButton btnDefaultEditor 
               Caption         =   "..."
               Height          =   300
               Left            =   5115
               Style           =   1  'Graphical
               TabIndex        =   108
               ToolTipText     =   "Click to select the .vbp file to edit the program - You need to have access to the source!"
               Top             =   210
               Width           =   315
            End
            Begin VB.TextBox txtDefaultEditor 
               Height          =   315
               Left            =   1440
               TabIndex        =   107
               Text            =   " eg. E:\vb6\fire call\FireCallWin.vbp"
               Top             =   195
               Width           =   3660
            End
            Begin VB.Label lblGitHub 
               Caption         =   $"frmPrefs.frx":7120
               ForeColor       =   &H8000000D&
               Height          =   915
               Left            =   1560
               TabIndex        =   111
               ToolTipText     =   "Double Click to visit github"
               Top             =   1440
               Width           =   4935
            End
            Begin VB.Label lblDebug 
               Caption         =   $"frmPrefs.frx":71E7
               Height          =   930
               Index           =   9
               Left            =   1545
               TabIndex        =   110
               Top             =   690
               Width           =   4785
            End
            Begin VB.Label lblDebug 
               Caption         =   "Default Editor :"
               Height          =   255
               Index           =   7
               Left            =   285
               TabIndex        =   109
               Tag             =   "lblSharedInputFile"
               Top             =   225
               Width           =   1350
            End
         End
         Begin VB.TextBox txtDblClickCommand 
            Height          =   315
            Left            =   1515
            TabIndex        =   39
            ToolTipText     =   "Enter a Windows command for the gauge to operate when double-clicked."
            Top             =   1095
            Width           =   3660
         End
         Begin VB.CommandButton btnOpenFile 
            Caption         =   "..."
            Height          =   300
            Left            =   5175
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Click to select a particular file for the gauge to run or open when double-clicked."
            Top             =   2250
            Width           =   315
         End
         Begin VB.TextBox txtOpenFile 
            Height          =   315
            Left            =   1515
            TabIndex        =   35
            ToolTipText     =   "Enter a particular file for the gauge to run or open when double-clicked."
            Top             =   2235
            Width           =   3660
         End
         Begin VB.ComboBox cmbDebug 
            Height          =   315
            ItemData        =   "frmPrefs.frx":728B
            Left            =   1530
            List            =   "frmPrefs.frx":728D
            Style           =   2  'Dropdown List
            TabIndex        =   32
            ToolTipText     =   "Choose to set debug mode."
            Top             =   -15
            Width           =   2160
         End
         Begin VB.Label lblDebug 
            Caption         =   "DblClick Command :"
            Height          =   510
            Index           =   1
            Left            =   -15
            TabIndex        =   41
            Tag             =   "lblPrefixString"
            Top             =   1155
            Width           =   1545
         End
         Begin VB.Label lblConfigurationTab 
            Caption         =   "Shift+double-clicking on the widget image will open this file. "
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   40
            Top             =   2730
            Width           =   3705
         End
         Begin VB.Label lblDebug 
            Caption         =   "Default command to run when the gauge receives a double-click eg.  mmsys.cpl to run the sounds utility."
            Height          =   570
            Index           =   5
            Left            =   1590
            TabIndex        =   38
            Tag             =   "lblSharedInputFileDesc"
            Top             =   1605
            Width           =   4410
         End
         Begin VB.Label lblDebug 
            Caption         =   "Open File :"
            Height          =   255
            Index           =   4
            Left            =   645
            TabIndex        =   37
            Tag             =   "lblSharedInputFile"
            Top             =   2280
            Width           =   1350
         End
         Begin VB.Label lblDebug 
            Caption         =   "Turning on the debugging will provide extra information in the debug window.  *"
            Height          =   495
            Index           =   2
            Left            =   1545
            TabIndex        =   34
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   4455
         End
         Begin VB.Label lblDebug 
            Caption         =   "Debug :"
            Height          =   375
            Index           =   0
            Left            =   855
            TabIndex        =   33
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   1740
         End
      End
   End
   Begin VB.Frame fraPosition 
      Caption         =   "Position"
      Height          =   7440
      Left            =   270
      TabIndex        =   28
      Top             =   1230
      Visible         =   0   'False
      Width           =   8520
      Begin VB.Frame fraPositionInner 
         BorderStyle     =   0  'None
         Height          =   6960
         Left            =   150
         TabIndex        =   29
         Top             =   300
         Width           =   7680
         Begin VB.TextBox txtLandscapeHoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   53
            Top             =   4425
            Width           =   2130
         End
         Begin VB.CheckBox chkPreventDragging 
            Caption         =   "Widget Position Locked. *"
            Height          =   225
            Left            =   2265
            TabIndex        =   99
            ToolTipText     =   "Checking this box turns off the ability to drag the program with the mouse, locking it in position."
            Top             =   3465
            Width           =   2505
         End
         Begin VB.TextBox txtPortraitYoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   59
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   6465
            Width           =   2130
         End
         Begin VB.TextBox txtPortraitHoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   57
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   6000
            Width           =   2130
         End
         Begin VB.TextBox txtLandscapeVoffset 
            Height          =   315
            Left            =   2250
            TabIndex        =   55
            ToolTipText     =   "Enter a prefix/nickname for outgoing messages."
            Top             =   4875
            Width           =   2130
         End
         Begin VB.ComboBox cmbWidgetLandscape 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   3930
            Width           =   2160
         End
         Begin VB.ComboBox cmbWidgetPortrait 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   48
            ToolTipText     =   "Choose the alarm sound."
            Top             =   5505
            Width           =   2160
         End
         Begin VB.ComboBox cmbWidgetPosition 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   45
            ToolTipText     =   "Choose the alarm sound."
            Top             =   2100
            Width           =   2160
         End
         Begin VB.ComboBox cmbAspectHidden 
            Height          =   315
            Left            =   2265
            Style           =   2  'Dropdown List
            TabIndex        =   42
            ToolTipText     =   "Choose the alarm sound."
            Top             =   0
            Width           =   2160
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   7
            Left            =   4530
            TabIndex        =   113
            Tag             =   "lblPrefixString"
            Top             =   6495
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   5
            Left            =   4530
            TabIndex        =   112
            Tag             =   "lblPrefixString"
            Top             =   6045
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "*"
            Height          =   255
            Index           =   1
            Left            =   4545
            TabIndex        =   102
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   345
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   4
            Left            =   4530
            TabIndex        =   98
            Tag             =   "lblPrefixString"
            Top             =   4905
            Width           =   435
         End
         Begin VB.Label lblPosition 
            Caption         =   "(px)"
            Height          =   300
            Index           =   2
            Left            =   4530
            TabIndex        =   97
            Tag             =   "lblPrefixString"
            Top             =   4500
            Width           =   390
         End
         Begin VB.Label lblPosition 
            Caption         =   "Portrait Top Y pos :"
            Height          =   510
            Index           =   17
            Left            =   645
            TabIndex        =   60
            Tag             =   "lblPrefixString"
            Top             =   6480
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Portrait Left X pos :"
            Height          =   510
            Index           =   16
            Left            =   660
            TabIndex        =   58
            Tag             =   "lblPrefixString"
            Top             =   6015
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Landscape Top Y pos :"
            Height          =   510
            Index           =   15
            Left            =   420
            TabIndex        =   56
            Tag             =   "lblPrefixString"
            Top             =   4905
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Landscape Left X pos :"
            Height          =   510
            Index           =   14
            Left            =   420
            TabIndex        =   54
            Tag             =   "lblPrefixString"
            Top             =   4455
            Width           =   2175
         End
         Begin VB.Label lblPosition 
            Caption         =   "Locked in Landscape :"
            Height          =   435
            Index           =   13
            Left            =   450
            TabIndex        =   52
            Tag             =   "lblAlarmSound"
            Top             =   3975
            Width           =   2115
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":728F
            Height          =   3435
            Index           =   12
            Left            =   5145
            TabIndex        =   50
            Tag             =   "lblAlarmSoundDesc"
            Top             =   3480
            Width           =   2520
         End
         Begin VB.Label lblPosition 
            Caption         =   "Locked in Portrait :"
            Height          =   375
            Index           =   11
            Left            =   690
            TabIndex        =   49
            Tag             =   "lblAlarmSound"
            Top             =   5550
            Width           =   2040
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":7461
            Height          =   705
            Index           =   10
            Left            =   2250
            TabIndex        =   47
            Tag             =   "lblAlarmSoundDesc"
            Top             =   2550
            Width           =   5325
         End
         Begin VB.Label lblPosition 
            Caption         =   "Widget Position by Percent:"
            Height          =   375
            Index           =   8
            Left            =   195
            TabIndex        =   46
            Tag             =   "lblAlarmSound"
            Top             =   2145
            Width           =   2355
         End
         Begin VB.Label lblPosition 
            Caption         =   $"frmPrefs.frx":7500
            Height          =   3045
            Index           =   6
            Left            =   2265
            TabIndex        =   44
            Tag             =   "lblAlarmSoundDesc"
            Top             =   450
            Width           =   5370
         End
         Begin VB.Label lblPosition 
            Caption         =   "Aspect Ratio Hidden Mode :"
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   43
            Tag             =   "lblAlarmSound"
            Top             =   45
            Width           =   2145
         End
      End
   End
   Begin VB.Frame fraCorner 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   8460
      TabIndex        =   181
      Top             =   9990
      Width           =   465
      Begin VB.Label lblDragCorner 
         Caption         =   "o"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   270
         TabIndex        =   182
         ToolTipText     =   "drag me"
         Top             =   390
         Visible         =   0   'False
         Width           =   345
      End
   End
   Begin VB.Label lblSize 
      Caption         =   "Size in twips"
      Height          =   285
      Left            =   1875
      TabIndex        =   114
      Top             =   9780
      Visible         =   0   'False
      Width           =   4170
   End
   Begin VB.Label lblAsterix 
      Caption         =   "All controls marked with a * take effect immediately."
      Height          =   300
      Left            =   1920
      TabIndex        =   100
      Top             =   10155
      Width           =   3870
   End
   Begin VB.Menu prefsMnuPopmenu 
      Caption         =   "The main menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "About Panzer Earth Widget"
      End
      Begin VB.Menu blank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCoffee 
         Caption         =   "Donate a coffee with KoFi"
      End
      Begin VB.Menu mnuSupport 
         Caption         =   "Contact Support"
      End
      Begin VB.Menu blank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuButton 
         Caption         =   "Theme Colours"
         Begin VB.Menu mnuLight 
            Caption         =   "Light Theme Enable"
         End
         Begin VB.Menu mnuDark 
            Caption         =   "High Contrast Theme Enable"
         End
         Begin VB.Menu mnuAuto 
            Caption         =   "Auto Theme Selection"
         End
      End
      Begin VB.Menu mnuLicenceA 
         Caption         =   "Display Licence Agreement"
      End
      Begin VB.Menu mnuClosePreferences 
         Caption         =   "Close Preferences"
      End
   End
End
Attribute VB_Name = "widgetPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder

'---------------------------------------------------------------------------------------
' Module    : widgetPrefs
' Author    : beededea
' Date      : 17/08/2022
' Purpose   : VB6 standard form to display the prefs
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
' Constants and APIs to create and subclass the dragCorner
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTBOTTOMRIGHT As Long = 17
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Constants defined for setting a theme to the prefs
Private Const COLOR_BTNFACE As Long = 15

' APIs declared for setting a theme to the prefs
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Private Types for determining prefs sizing
Private pPrefsDynamicSizingFlg As Boolean
Private pLastFormHeight As Long
Private Const pcPrefsFormHeight As Long = 10750
Private Const pcPrefsFormWidth  As Long = 8940

Private pPrefsFormResizedByDrag As Boolean
'------------------------------------------------------ ENDS

Private pPrefsStartupFlg As Boolean
Private pAllowSizeChangeFlg As Boolean
Private pAllowSkewChangeFlg As Boolean

' module level balloon tooltip variables for subclassed comboBoxes ONLY.
Private pCmbMultiMonitorResizeBalloonTooltip As String
Private pCmbScrollWheelDirectionBalloonTooltip As String
Private pCmbWindowLevelBalloonTooltip As String
Private pCmbHidingTimeBalloonTooltip As String
Private pCmbAspectHiddenBalloonTooltip As String
Private pCmbWidgetPositionBalloonTooltip As String
Private pCmbWidgetLandscapeBalloonTooltip As String
Private pCmbWidgetPortraitBalloonTooltip As String
Private pCmbDebugBalloonTooltip As String

Private mIsLoaded As Boolean ' property
Private mWidgetSize As Single   ' property
Private gdConstraintRatio As Double

Private Sub btnDefaultEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnDefaultEditor.hwnd, "Field to hold the path to a Visual Basic Project (VBP) file you would like to execute on a right click menu, edit option, if you select the adjacent button a file explorer will appear allowing you to select the VBP file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the Default Editor Field", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnGithubHome_Click
' Author    : beededea
' Date      : 22/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btnGithubHome_Click()
   On Error GoTo btnGithubHome_Click_Error

    Call menuForm.mnuGithubHome_Click

   On Error GoTo 0
   Exit Sub

btnGithubHome_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnGithubHome_Click of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkFormVisible_Click
' Author    : beededea
' Date      : 16/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkFormVisible_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo chkFormVisible_Click_Error
    
    If gbStartupFlg = True Then Exit Sub

    If chkFormVisible.Value = 0 Then
        gsFormVisible = "0"
    Else
        gsFormVisible = "1"
    End If

    btnSave.Enabled = True ' enable the save button
    
    If pPrefsStartupFlg = False Then ' don't run this on startup
    
        answer = vbYes
        answerMsg = "You must close this widget and soft reload it, in order to change the underlying form visibility, do you want me to close and restart this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "Check Form Visibility Confirmation", True, "chkFormVisible")
    
        sPutINISetting "Software\TenShillings", "formVisible", gsFormVisible, gsSettingsFile
        
        If answer = vbYes Then
            Call reloadProgram
        End If
    End If

    On Error GoTo 0
    Exit Sub

chkFormVisible_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkFormVisible_Click of Form widgetPrefs"
End Sub





Private Sub fraCorner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraCorner.hwnd, "The Drag Corner is enabled when DPI awareness is ON. Click and drag me to resize the whole window.", _
                  TTIconInfo, "Help on the Drag Corner", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optWidgetTooltips_Click
' Author    : beededea
' Date      : 19/08/2023
' Purpose   : three options radio buttons for selecting the widget/cal tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optWidgetTooltips_Click(Index As Integer)
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    On Error GoTo optWidgetTooltips_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

    If pPrefsStartupFlg = False Then
        gsWidgetTooltips = CStr(Index)
    
        optWidgetTooltips(0).Tag = CStr(Index)
        optWidgetTooltips(1).Tag = CStr(Index)
        optWidgetTooltips(2).Tag = CStr(Index)
        
        sPutINISetting "Software\TenShillings", "widgetTooltips", gsWidgetTooltips, gsSettingsFile

        answer = vbYes
        answerMsg = "You must soft reload this widget, in order to change the tooltip setting, do you want me to reload this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "Request to Enable Tooltips", True, "optWidgetTooltipsClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call reloadProgram
        End If
    End If


   On Error GoTo 0
   Exit Sub

optWidgetTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optWidgetTooltips_Click of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkWidgetFunctions_Click
' Author    : beededea
' Date      : 30/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub chkWidgetFunctions_Click()
    On Error GoTo chkWidgetFunctions_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

    On Error GoTo 0
    Exit Sub

chkWidgetFunctions_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkWidgetFunctions_Click of Form widgetPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : optPrefsTooltips_Click
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : three options radio buttons for selecting the VB6 preference form tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optPrefsTooltips_Click(Index As Integer)

   On Error GoTo optPrefsTooltips_Click_Error

    If pPrefsStartupFlg = False Then
    
        If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
        gsPrefsTooltips = CStr(Index)
        optPrefsTooltips(0).Tag = CStr(Index)
        optPrefsTooltips(1).Tag = CStr(Index)
        optPrefsTooltips(2).Tag = CStr(Index)
        
        sPutINISetting "Software\TenShillings", "prefsTooltips", gsPrefsTooltips, gsSettingsFile
        
        ' set the tooltips on the prefs screen
        Call setPrefsTooltips
    End If
     
   On Error GoTo 0
   Exit Sub

optPrefsTooltips_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optPrefsTooltips_Click of Form widgetPrefs"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmbMultiMonitorResize_Click
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : For monitors of different sizes, this allows you to resize the widget to suit the monitor it is currently sitting on.
'---------------------------------------------------------------------------------------
'
Private Sub cmbMultiMonitorResize_Click()
   On Error GoTo cmbMultiMonitorResize_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    
    If pPrefsStartupFlg = True Then Exit Sub
    
    gsMultiMonitorResize = CStr(cmbMultiMonitorResize.ListIndex)
    
    ' saves the current ratios for the RC form and the absolute sizes for the Prefs form
    Call saveMainFormsCurrentSizeAndRatios

   On Error GoTo 0
   Exit Sub

cmbMultiMonitorResize_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbMultiMonitorResize_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkShowHelp_Click
' Author    : beededea
' Date      : 03/07/2024
' Purpose   : show help on the program startup
'---------------------------------------------------------------------------------------
'
Private Sub chkShowHelp_Click()
   On Error GoTo chkShowHelp_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    If chkShowHelp.Value = 1 Then
        gsShowHelp = "1"
    Else
        gsShowHelp = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkShowHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowHelp_Click of Form widgetPrefs"
End Sub


' ----------------------------------------------------------------
' Procedure Name: Form_Initialize
' Purpose:
' Procedure Kind: Constructor (Initialize)
' Procedure Access: Private
' Author: beededea
' Date: 05/10/2023
' ----------------------------------------------------------------
Private Sub Form_Initialize()
    On Error GoTo Form_Initialize_Error
    
    ' initialise private variables
    Call initialisePrefsVars

    On Error GoTo 0
    Exit Sub

Form_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Initialize of Form widgetPrefs"
    
    End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load     WidgetPrefs
' Author    : beededea
' Date      : 25/04/2023
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    On Error GoTo Form_Load_Error
    
    Me.Visible = False
    Me.mnuAbout.Caption = "About TenShillings Widget Cairo" & gsRichClientEnvironment & " Cairo " & gsCodingEnvironment & " widget"

    pPrefsStartupFlg = True ' this is used to prevent some control initialisations from running code at startup
    IsLoaded = True
    gbWindowLevelWasChanged = False
    gdPrefsStartWidth = pcPrefsFormWidth
    gdPrefsStartHeight = pcPrefsFormHeight
    pPrefsFormResizedByDrag = False
            
    ' subclass ALL forms created by intercepting WM_Create messages, identifying dialog forms to centre them in the middle of the monitor - specifically the font form.
    If Not InIde Then subclassDialogForms
    
    ' subclass specific WidgetPrefs controls that need additional functionality that VB6 does not provide (scrollwheel/balloon tooltips)
    Call subClassControls
    
    ' set form resizing
    Call setFormResizingVars
    
    ' reverts TwinBasic form themeing to that of the earlier classic look and feel
    #If twinbasic Then
       Call setVisualStyles
    #End If
       
    ' read the last saved position from the settings.ini
    Call readPrefsPosition
    
    ' size and position the frames and buttons
    Call positionPrefsFramesButtons
        
    ' determine the frame heights in dynamic sizing or normal mode
    Call setframeHeights
    
    ' set the text in any labels that need a vbCrLf to space the text
    Call setPrefsLabels
    
    ' populate all the comboboxes in the prefs form
    Call populatePrefsComboBoxes
        
    ' adjust all the preferences and main program controls
    Call adjustPrefsControls
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips
    
    ' adjust the theme used by the prefs alone
    Call adjustPrefsTheme
    
    ' make the last used tab appear on startup
    Call showLastTab
    
    ' load the about text and load into prefs
    Call loadPrefsAboutText
    
    ' load the preference icons from a previously populated CC imageList
    Call loadHigherResPrefsImages
    
    ' set the height of the whole form not higher than the screen size, cause a form_resize event
    Call setPrefsHeight
    
    ' position the prefs on the current monitor
    Call positionPrefsMonitor
    
    ' start the timers
    Call startPrefsTimers
    
    glWidgetPrefsOldHeightTwips = widgetPrefs.height
    glWidgetPrefsOldWidthTwips = widgetPrefs.Width
    
    ' end the startup by un-setting the start global-ish flag
    pPrefsStartupFlg = False
    btnSave.Enabled = False ' disable the save button

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form widgetPrefs"

End Sub



' ---------------------------------------------------------------------------------------
' Procedure : initialisePrefsVars
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : initialise private variables
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrefsVars()

   On Error GoTo initialisePrefsVars_Error

    pPrefsDynamicSizingFlg = False
    pLastFormHeight = 0
    pPrefsStartupFlg = False
    pAllowSizeChangeFlg = False
    pAllowSkewChangeFlg = False
    pCmbMultiMonitorResizeBalloonTooltip = vbNullString
    pCmbScrollWheelDirectionBalloonTooltip = vbNullString
    pCmbWindowLevelBalloonTooltip = vbNullString
    pCmbHidingTimeBalloonTooltip = vbNullString
    pCmbAspectHiddenBalloonTooltip = vbNullString
    pCmbWidgetPositionBalloonTooltip = vbNullString
    pCmbWidgetLandscapeBalloonTooltip = vbNullString
    pCmbWidgetPortraitBalloonTooltip = vbNullString
    pCmbDebugBalloonTooltip = vbNullString
    pPrefsFormResizedByDrag = False
    mIsLoaded = False ' property

   On Error GoTo 0
   Exit Sub

initialisePrefsVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrefsVars of Form widgetPrefs"

End Sub

     
'
'---------------------------------------------------------------------------------------
' Procedure : setFormResizingVars
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : set form resizing characteristics
'---------------------------------------------------------------------------------------
'
Private Sub setFormResizingVars()

   On Error GoTo setFormResizingVars_Error

    With lblDragCorner
      .ForeColor = &H80000015
      .BackStyle = vbTransparent
      .AutoSize = True
      .Font.Size = 12
      .Font.Name = "Marlett"
      .Caption = "o"
      .Font.Bold = False
      .Visible = False
    End With
    
    If gsDpiAwareness = "1" Then
        pPrefsDynamicSizingFlg = True
        chkEnableResizing.Value = 1
        lblDragCorner.Visible = True
    End If
    
    glWidgetPrefsOldHeightTwips = widgetPrefs.height
    glWidgetPrefsOldWidthTwips = widgetPrefs.Width

   On Error GoTo 0
   Exit Sub

setFormResizingVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setFormResizingVars of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : identifyPrefsPrimaryMonitor
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : note the monitor primary at the preferences form_load and store as glOldPrefsFormMonitorPrimary - will be resampled regularly later and compared
'---------------------------------------------------------------------------------------
'
Private Sub identifyPrefsPrimaryMonitor()
    'Dim prefsFormHeight As Long: prefsFormHeight = 0
    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
    
    On Error GoTo identifyPrefsPrimaryMonitor_Error
    
    'prefsFormHeight = gdPrefsStartHeight

    gPrefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
    glOldPrefsFormMonitorPrimary = gPrefsMonitorStruct.IsPrimary ' -1 true

   On Error GoTo 0
   Exit Sub

identifyPrefsPrimaryMonitor_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure identifyPrefsPrimaryMonitor of Form widgetPrefs"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : setPrefsHeight
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : set the height of the whole form not higher than the screen size, cause a form_resize event
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsHeight()

   On Error GoTo setPrefsHeight_Error

    If gsDpiAwareness = "1" Then
        gbPrefsFormResizedInCode = True
        If CLng(gsPrefsPrimaryHeightTwips) < glPhysicalScreenHeightTwips Then
            widgetPrefs.height = CLng(gsPrefsPrimaryHeightTwips) ' on first run this also sets the prefs to one third of the screen height (value set in readPrefsPosition)
        Else
            widgetPrefs.height = glPhysicalScreenHeightTwips - 1000
        End If
    End If

   On Error GoTo 0
   Exit Sub

setPrefsHeight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsHeight of Form widgetPrefs"
End Sub
   
'---------------------------------------------------------------------------------------
' Procedure : startPrefsTimers
' Author    : beededea
' Date      : 20/02/2025
' Purpose   :  start the timers
'---------------------------------------------------------------------------------------
'
Private Sub startPrefsTimers()

    ' start the timer that records the prefs position every 5 seconds
   On Error GoTo startPrefsTimers_Error

    tmrWritePositionAndSize.Enabled = True
    
    ' start the timer that detects a MOVE event on the preferences form
    tmrPrefsScreenResolution.Enabled = True

   On Error GoTo 0
   Exit Sub

startPrefsTimers_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure startPrefsTimers of Form widgetPrefs"

End Sub
    

#If twinbasic Then
    '---------------------------------------------------------------------------------------
    ' Procedure : setVisualStyles
    ' Author    : beededea
    ' Date      : 13/01/2025
    ' Purpose   : loop through all the controls and identify the labels and text boxes and disable modern styles
    '             reverts TwinBasic form themeing to that of the earlier classic look and feel.
    '---------------------------------------------------------------------------------------
    '
        Private Sub setVisualStyles()
            Dim Ctrl As Control
          
            On Error GoTo setVisualStyles_Error

            For Each Ctrl In widgetPrefs.Controls
                If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is textBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is PictureBox) Or (TypeOf Ctrl Is Slider) Then
                
                    Ctrl.VisualStyles = False
                End If
            Next

       On Error GoTo 0
       Exit Sub

setVisualStyles_Error:

        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setVisualStyles of Form widgetPrefs"
        End Sub
#End If



'---------------------------------------------------------------------------------------
' Procedure : subClassControls
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : sub classing code to capture form movement and intercept messages to the comboboxes to provide missing balloon tooltips functionality
'---------------------------------------------------------------------------------------
'
Private Sub subClassControls()
    
   On Error GoTo subClassControls_Error

    If InIde And gbReload = False Then
        MsgBox "NOTE: Running in IDE so Sub classing is disabled" & vbCrLf & "Mousewheel will not scroll icon maps and balloon tooltips will not display on comboboxes" & vbCrLf & vbCrLf & _
            "In addition, the display screen will not show messages as it currently crashes when run within the IDE."
    Else
        ' sub classing code to intercept messages to the form itself in order to capture WM_EXITSIZEMOVE messages that occur AFTER the form has been resized
        
        Call SubclassForm(widgetPrefs.hwnd, ObjPtr(widgetPrefs))
        
        'now the comboboxes in order to capture the mouseOver and display the balloon tooltips
        
        Call SubclassComboBox(cmbMultiMonitorResize.hwnd, ObjPtr(cmbMultiMonitorResize))
        Call SubclassComboBox(cmbScrollWheelDirection.hwnd, ObjPtr(cmbScrollWheelDirection))
        Call SubclassComboBox(cmbWindowLevel.hwnd, ObjPtr(cmbWindowLevel))
        Call SubclassComboBox(cmbHidingTime.hwnd, ObjPtr(cmbHidingTime))
        
        Call SubclassComboBox(cmbWidgetLandscape.hwnd, ObjPtr(cmbWidgetLandscape))
        Call SubclassComboBox(cmbWidgetPortrait.hwnd, ObjPtr(cmbWidgetPortrait))
        Call SubclassComboBox(cmbWidgetPosition.hwnd, ObjPtr(cmbWidgetPosition))
        Call SubclassComboBox(cmbAspectHidden.hwnd, ObjPtr(cmbAspectHidden))
        Call SubclassComboBox(cmbDebug.hwnd, ObjPtr(cmbDebug))
        
    End If

    On Error GoTo 0
    Exit Sub

subClassControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure subClassControls of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MouseMoveOnComboText
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : Add a balloon tooltip dynamically to combo boxes using subclassing, called by combobox_proc
'             (VB6 will not allow Elroy's advanced tooltips to show on VB6 comboboxes, we must subclass the controls)
'             Note: Each control must also be added to the subClassControls routine
'---------------------------------------------------------------------------------------
'
Public Sub MouseMoveOnComboText(sComboName As String)
    Dim sTitle As String
    Dim sText As String

    On Error GoTo MouseMoveOnComboText_Error
    
    Select Case sComboName
        Case "cmbMultiMonitorResize"
            sTitle = "Help on the Drop Down Icon Filter"
            sText = pCmbMultiMonitorResizeBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbMultiMonitorResize.hwnd, sText, , sTitle, , , , True
        Case "cmbScrollWheelDirection"
            sTitle = "Help on the Scroll Wheel Direction"
            sText = pCmbScrollWheelDirectionBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbScrollWheelDirection.hwnd, sText, , sTitle, , , , True
        Case "cmbWindowLevel"
            sTitle = "Help on the Window Level"
            sText = pCmbWindowLevelBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbWindowLevel.hwnd, sText, , sTitle, , , , True
        Case "cmbHidingTime"
            sTitle = "Help on the Hiding Time"
            sText = pCmbHidingTimeBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbHidingTime.hwnd, sText, , sTitle, , , , True
            
        Case "cmbAspectHidden"
            sTitle = "Help on Hiding in Landscape/Portrait Mode"
            sText = pCmbAspectHiddenBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbAspectHidden.hwnd, sText, , sTitle, , , , True
        Case "cmbWidgetPosition"
            sTitle = "Help on Widget Position in Landscape/Portrait Modes"
            sText = pCmbWidgetPositionBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbWidgetPosition.hwnd, sText, , sTitle, , , , True
        Case "cmbWidgetLandscape"
            sTitle = "Help on Widget Locking in Landscape Mode"
            sText = pCmbWidgetLandscapeBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbWidgetLandscape.hwnd, sText, , sTitle, , , , True
        Case "cmbWidgetPortrait"
            sTitle = "Help on Widget Locking in Portrait Mode"
            sText = pCmbWidgetPortraitBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbWidgetPortrait.hwnd, sText, , sTitle, , , , True
        Case "cmbDebug"
            sTitle = "Help on Debug Mode"
            sText = pCmbDebugBalloonTooltip
            If gsPrefsTooltips = "0" Then CreateToolTip cmbDebug.hwnd, sText, , sTitle, , , , True
            
    End Select
    
   On Error GoTo 0
   Exit Sub

MouseMoveOnComboText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MouseMoveOnComboText of Form widgetPrefs"
End Sub


' ---------------------------------------------------------------------------------------
' Procedure : positionPrefsMonitor
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : position the prefs on the current monitor
'---------------------------------------------------------------------------------------
'
Public Sub positionPrefsMonitor()

    Dim formLeftTwips As Long: formLeftTwips = 0
    Dim formTopTwips As Long: formTopTwips = 0
    'Dim monitorCount As Long: monitorCount = 0
    
    On Error GoTo positionPrefsMonitor_Error
    
    If gsDpiAwareness = "1" Then
        formLeftTwips = Val(gsPrefsHighDpiXPosTwips)
        formTopTwips = Val(gsPrefsHighDpiYPosTwips)
    Else
        formLeftTwips = Val(gsPrefsLowDpiXPosTwips)
        formTopTwips = Val(gsPrefsLowDpiYPosTwips)
    End If
    
    If formLeftTwips = 0 Then
        If ((fMain.TenShillingsForm.Left + fMain.TenShillingsForm.Width) * glScreenTwipsPerPixelX) + 200 + widgetPrefs.Width > glPhysicalScreenWidthTwips Then
            widgetPrefs.Left = (fMain.TenShillingsForm.Left * glScreenTwipsPerPixelX) - (widgetPrefs.Width + 200)
        End If
    End If

    ' if a current location not stored then position to the middle of the screen
    
    If formLeftTwips <> 0 Then
        widgetPrefs.Left = formLeftTwips
    Else
        widgetPrefs.Left = glPhysicalScreenWidthTwips / 2 - widgetPrefs.Width / 2
    End If
    
    If formTopTwips <> 0 Then
        widgetPrefs.Top = formTopTwips
    Else
        widgetPrefs.Top = Screen.height / 2 - widgetPrefs.height / 2
    End If
    
    'monitorCount = fGetMonitorCount
    If glMonitorCount > 1 Then Call SetFormOnMonitor(Me.hwnd, formLeftTwips / fTwipsPerPixelX, formTopTwips / fTwipsPerPixelY)
    
    ' calculate the on-screen widget position
    If Me.Left < 0 Then
        widgetPrefs.Left = 10
    End If
    If Me.Top < 0 Then
        widgetPrefs.Top = 0
    End If
    If Me.Left > glVirtualScreenWidthTwips - 2500 Then
        widgetPrefs.Left = glVirtualScreenWidthTwips - 2500
    End If
    If Me.Top > glVirtualScreenHeightTwips - 2500 Then
        widgetPrefs.Top = glVirtualScreenHeightTwips - 2500
    End If
    
    ' if just one monitor or the global switch is off then exit
    If glMonitorCount > 1 And LTrim$(gsMultiMonitorResize) = "2" Then
    
        ' note the monitor primary at the preferences form_load and store as glOldwidgetFormMonitorPrimary
        Call identifyPrefsPrimaryMonitor

        If gPrefsMonitorStruct.IsPrimary = True Then
            gbPrefsFormResizedInCode = True
            gsPrefsPrimaryHeightTwips = fGetINISetting("Software\TenShillings", "prefsPrimaryHeightTwips", gsSettingsFile)
            If Val(gsPrefsPrimaryHeightTwips) <= 0 Then
                widgetPrefs.height = gdPrefsStartHeight
            Else
                widgetPrefs.height = CLng(gsPrefsPrimaryHeightTwips)
            End If
        Else
            gsPrefsSecondaryHeightTwips = fGetINISetting("Software\TenShillings", "prefsSecondaryHeightTwips", gsSettingsFile)
            gbPrefsFormResizedInCode = True
            If Val(gsPrefsSecondaryHeightTwips) <= 0 Then
                widgetPrefs.height = gdPrefsStartHeight
            Else
                widgetPrefs.height = CLng(gsPrefsSecondaryHeightTwips)
            End If
        End If
    End If

    '' tenShillingsOverlay.RotateBusyTimer = True
    
    On Error GoTo 0
    Exit Sub

positionPrefsMonitor_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsMonitor of Form widgetPrefs"
End Sub
    
    


'---------------------------------------------------------------------------------------
' Procedure : chkDpiAwareness_Click
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : toggle for setting the DPI awareness
'---------------------------------------------------------------------------------------
'
Private Sub chkDpiAwareness_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo chkDpiAwareness_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    If pPrefsStartupFlg = False Then ' don't run this on startup
                    
        answer = vbYes
        answerMsg = "You must close this widget and HARD restart it, in order to change the widget's DPI awareness (a simple soft reload just won't cut it), do you want me to close and restart this widget? I can do it now for you."
        answer = msgBoxA(answerMsg, vbYesNo, "DpiAwareness Confirmation", True, "chkDpiAwarenessRestart")
        
        If chkDpiAwareness.Value = 0 Then
            gsDpiAwareness = "0"
        Else
            gsDpiAwareness = "1"
        End If

        sPutINISetting "Software\TenShillings", "dpiAwareness", gsDpiAwareness, gsSettingsFile
        
        If answer = vbNo Then
            answer = vbYes
            answerMsg = "OK, the widget is still DPI aware until you restart. Some forms may show abnormally."
            answer = msgBoxA(answerMsg, vbOKOnly, "DpiAwareness Notification", True, "chkDpiAwarenessAbnormal")
        
            Exit Sub
        Else

            sPutINISetting "Software\TenShillings", "dpiAwareness", gsDpiAwareness, gsSettingsFile
            'Call reloadProgram ' this is insufficient, image controls still fail to resize and autoscale correctly
            Call hardRestart
        End If

    End If

   On Error GoTo 0
   Exit Sub

chkDpiAwareness_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkDpiAwareness_Click of Form widgetPrefs"
End Sub







'---------------------------------------------------------------------------------------
' Procedure : chkShowTaskbar_Click
' Author    : beededea
' Date      : 19/07/2023
' Purpose   : toggle for showing the program in the taskbar
'---------------------------------------------------------------------------------------
'
Private Sub chkShowTaskbar_Click()

   On Error GoTo chkShowTaskbar_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    If chkShowTaskbar.Value = 1 Then
        gsShowTaskbar = "1"
    Else
        gsShowTaskbar = "0"
    End If

   On Error GoTo 0
   Exit Sub

chkShowTaskbar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkShowTaskbar_Click of Form widgetPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : btnResetMessages_Click
' Author    : beededea
' Date      : 01/10/2023
' Purpose   : reset the improved message boxes so that any hidden boxes will reappear
'---------------------------------------------------------------------------------------
'
Private Sub btnResetMessages_Click()

    On Error GoTo btnResetMessages_Click_Error
        
    ' Clear all the message box "show again" entries in the registry
    Call clearAllMessageBoxRegistryEntries
    
    MsgBox "Message boxes fully reset, confirmation pop-ups will continue as normal."

    On Error GoTo 0
    Exit Sub

btnResetMessages_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnAboutDebugInfo_Click
' Author    : beededea
' Date      : 03/03/2020
' Purpose   : Enabling debug mode - not implemented
'---------------------------------------------------------------------------------------
'
Private Sub btnAboutDebugInfo_Click()

   On Error GoTo btnAboutDebugInfo_Click_Error
   'If giDebugFlg = 1 Then Debug.Print "%btnAboutDebugInfo_Click"

    'mnuDebug_Click
    MsgBox "The debug mode is not yet enabled."

   On Error GoTo 0
   Exit Sub

btnAboutDebugInfo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAboutDebugInfo_Click of form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDonate_Click
' Author    : beededea
' Date      : 13/06/2020
' Purpose   : Donate button
'---------------------------------------------------------------------------------------
'
Private Sub btnDonate_Click()
   On Error GoTo btnDonate_Click_Error

    Call mnuCoffee_ClickEvent

   On Error GoTo 0
   Exit Sub

btnDonate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDonate_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnFacebook_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : FB button
'---------------------------------------------------------------------------------------
'
Private Sub btnFacebook_Click()
   On Error GoTo btnFacebook_Click_Error
   'If giDebugFlg = 1 Then DebugPrint "%btnFacebook_Click"

    Call menuForm.mnuFacebook_Click
    

   On Error GoTo 0
   Exit Sub

btnFacebook_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnFacebook_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnOpenFile_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : button for opening a target file for dblClicking
'---------------------------------------------------------------------------------------
'
Private Sub btnOpenFile_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo btnOpenFile_Click_Error

    Call addTargetFile(txtOpenFile.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtOpenFile.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        'answer = MsgBox("The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?", vbYesNo)
        answer = vbYes
        answerMsg = "The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?"
        answer = msgBoxA(answerMsg, vbYesNo, "Create file confirmation", False)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If

    On Error GoTo 0
    Exit Sub

btnOpenFile_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnOpenFile_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnUpdate_Click
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : auto update button
'---------------------------------------------------------------------------------------
'
Private Sub btnUpdate_Click()
   On Error GoTo btnUpdate_Click_Error
   'If giDebugFlg = 1 Then DebugPrint "%btnUpdate_Click"

    'MsgBox "The update button is not yet enabled."
    menuForm.mnuLatest_Click

   On Error GoTo 0
   Exit Sub

btnUpdate_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnUpdate_Click of Form widgetPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : chkGenStartup_Click
' Author    : beededea
' Date      : 30/09/2023
' Purpose   : Toggle automatic startup by writing to the registry on save
'---------------------------------------------------------------------------------------
'
Private Sub chkGenStartup_Click()
    On Error GoTo chkGenStartup_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

    On Error GoTo 0
    Exit Sub

chkGenStartup_Click_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkGenStartup_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDefaultEditor_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : button for selecting the default VB6 editor VBP project file
'---------------------------------------------------------------------------------------
'
Private Sub btnDefaultEditor_Click()
    Dim retFileName As String: retFileName = vbNullString
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo btnDefaultEditor_Click_Error

    Call addTargetFile(txtDefaultEditor.Text, retFileName)
    
    If retFileName <> vbNullString Then
        txtDefaultEditor.Text = retFileName ' strips the buffered bit, leaving just the filename
    End If
    
    If retFileName = vbNullString Then
        Exit Sub
    End If
    
    If Not fFExists(retFileName) Then
        answer = vbYes
        answerMsg = "The file doesn't currently exist, do you want me to create the chosen file, " & "   -  are you sure?"
        answer = msgBoxA(answerMsg, vbYesNo, "Default Editor Confirmation", False)
        If answer = vbNo Then
            Exit Sub
        End If
    
        'create new
        Open retFileName For Output As #1
        Close #1
    End If

    On Error GoTo 0
    Exit Sub

btnDefaultEditor_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDefaultEditor_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
    
End Sub




'---------------------------------------------------------------------------------------
' Procedure : chkIgnoreMouse_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : toggle to ignore any mouse clicks
'---------------------------------------------------------------------------------------
'
Private Sub chkIgnoreMouse_Click()
   On Error GoTo chkIgnoreMouse_Click_Error

    If chkIgnoreMouse.Value = 0 Then
        gsIgnoreMouse = "0"
    Else
        gsIgnoreMouse = "1"
    End If

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkIgnoreMouse_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkIgnoreMouse_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkPreventDragging_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : lock the program in place, prevent dragging
'---------------------------------------------------------------------------------------
'
Private Sub chkPreventDragging_Click()
    On Error GoTo chkPreventDragging_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    ' immediately make the widget locked in place
    If chkPreventDragging.Value = 0 Then
        tenShillingsOverlay.Locked = False
        gsPreventDragging = "0"
        menuForm.mnuLockWidget.Checked = False
        If gsAspectRatio = "landscape" Then
            txtLandscapeHoffset.Text = vbNullString
            txtLandscapeVoffset.Text = vbNullString
        Else
            txtPortraitHoffset.Text = vbNullString
            txtPortraitYoffset.Text = vbNullString
        End If
    Else
        tenShillingsOverlay.Locked = True
        gsPreventDragging = "1"
        menuForm.mnuLockWidget.Checked = True
        If gsAspectRatio = "landscape" Then
            txtLandscapeHoffset.Text = fMain.TenShillingsForm.Left
            txtLandscapeVoffset.Text = fMain.TenShillingsForm.Top
        Else
            txtPortraitHoffset.Text = fMain.TenShillingsForm.Left
            txtPortraitYoffset.Text = fMain.TenShillingsForm.Top
        End If
    End If

    On Error GoTo 0
    Exit Sub

chkPreventDragging_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkPreventDragging_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkWidgetHidden_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : toggle to hide the program
'---------------------------------------------------------------------------------------
'
Private Sub chkWidgetHidden_Click()
   On Error GoTo chkWidgetHidden_Click_Error

    If chkWidgetHidden.Value = 0 Then
        'tenShillingsOverlay.Hidden = False
        fMain.TenShillingsForm.Visible = True

        frmTimer.revealWidgetTimer.Enabled = False
        gsWidgetHidden = "0"
    Else
        'tenShillingsOverlay.Hidden = True
        fMain.TenShillingsForm.Visible = False


        frmTimer.revealWidgetTimer.Enabled = True
        gsWidgetHidden = "1"
    End If
    
    sPutINISetting "Software\TenShillings", "widgetHidden", gsWidgetHidden, gsSettingsFile
    
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkWidgetHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkWidgetHidden_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbAspectHidden_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : selector for hiding in portrait/landscape mode
'---------------------------------------------------------------------------------------
'
Private Sub cmbAspectHidden_Click()

   On Error GoTo cmbAspectHidden_Click_Error

    If cmbAspectHidden.ListIndex = 1 And gsAspectRatio = "portrait" Then
        'tenShillingsOverlay.Hidden = True
        fMain.TenShillingsForm.Visible = False
    ElseIf cmbAspectHidden.ListIndex = 2 And gsAspectRatio = "landscape" Then
        'tenShillingsOverlay.Hidden = True
        fMain.TenShillingsForm.Visible = False
    Else
        'tenShillingsOverlay.Hidden = False
        fMain.TenShillingsForm.Visible = True
    End If

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbAspectHidden_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbAspectHidden_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbDebug_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : debug selector
'---------------------------------------------------------------------------------------
'
Private Sub cmbDebug_Click()
    On Error GoTo cmbDebug_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    If cmbDebug.ListIndex = 0 Then
        txtDefaultEditor.Text = "eg. E:\vb6\TenShillings-" & gsRichClientEnvironment & "-Widget-VB6MkII\TenShillings-" & gsRichClientEnvironment & "-Widget-" & gsCodingEnvironment & ".vbp"
        txtDefaultEditor.Enabled = False
        lblDebug(7).Enabled = False
        btnDefaultEditor.Enabled = False
        lblDebug(9).Enabled = False
    Else
        #If twinbasic Then
            txtDefaultEditor.Text = gsDefaultTBEditor
        #Else
            txtDefaultEditor.Text = gsDefaultVB6Editor
        #End If
        txtDefaultEditor.Enabled = True
        lblDebug(7).Enabled = True
        btnDefaultEditor.Enabled = True
        lblDebug(9).Enabled = True
    End If

    On Error GoTo 0
    Exit Sub

cmbDebug_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbDebug_Click of Form widgetPrefs"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : cmbHidingTime_Click
' Author    : beededea
' Date      : 17/02/2025
' Purpose   : enable the save button if a hiding time is selected
'---------------------------------------------------------------------------------------
'
Private Sub cmbHidingTime_Click()
   On Error GoTo cmbHidingTime_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbHidingTime_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbHidingTime_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbScrollWheelDirection_Click
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : selector for resizing using the mouse Scroll Wheel Direction
'---------------------------------------------------------------------------------------
'
Private Sub cmbScrollWheelDirection_Click()
   On Error GoTo cmbScrollWheelDirection_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    'tenShillingsOverlay.ZoomDirection = cmbScrollWheelDirection.List(cmbScrollWheelDirection.ListIndex)

   On Error GoTo 0
   Exit Sub

cmbScrollWheelDirection_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbScrollWheelDirection_Click of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : cmbWidgetLandscape_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   : option dropdown for locking in landscape mode after save
'---------------------------------------------------------------------------------------
'
Private Sub cmbWidgetLandscape_Click()
   On Error GoTo cmbWidgetLandscape_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbWidgetLandscape_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetLandscape_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWidgetPortrait_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   : option dropdown for locking in portrait mode after save
'---------------------------------------------------------------------------------------
'
Private Sub cmbWidgetPortrait_Click()
   On Error GoTo cmbWidgetPortrait_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

cmbWidgetPortrait_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetPortrait_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmbWidgetPosition_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : option dropdown to position by percent after save
'---------------------------------------------------------------------------------------
'
Private Sub cmbWidgetPosition_Click()
    On Error GoTo cmbWidgetPosition_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    If cmbWidgetPosition.ListIndex = 1 Then
        cmbWidgetLandscape.ListIndex = 0
        cmbWidgetPortrait.ListIndex = 0
        cmbWidgetLandscape.Enabled = False
        cmbWidgetPortrait.Enabled = False
        txtLandscapeHoffset.Enabled = False
        txtLandscapeVoffset.Enabled = False
        txtPortraitHoffset.Enabled = False
        txtPortraitYoffset.Enabled = False
        
    Else
        cmbWidgetLandscape.Enabled = True
        cmbWidgetPortrait.Enabled = True
        txtLandscapeHoffset.Enabled = True
        txtLandscapeVoffset.Enabled = True
        txtPortraitHoffset.Enabled = True
        txtPortraitYoffset.Enabled = True
    End If

    On Error GoTo 0
    Exit Sub

cmbWidgetPosition_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWidgetPosition_Click of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   : property by val to manually determine whether the preference form is loaded. It does this without
'             touching a VB6 intrinsic form property which would then load the form itself.
'---------------------------------------------------------------------------------------
'
Public Property Get IsLoaded() As Boolean
 
   On Error GoTo IsLoaded_Error

    IsLoaded = mIsLoaded

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property

'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : beededea
' Date      : 16/12/2024
' Purpose   : property by val to manually determine whether the preference form is loaded. It does this without
'             touching a VB6 intrinsic form property which would then load the form itself.
'---------------------------------------------------------------------------------------
'
Public Property Let IsLoaded(ByVal newValue As Boolean)
 
   On Error GoTo IsLoaded_Error

   If mIsLoaded <> newValue Then mIsLoaded = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

IsLoaded_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsLoaded of Form widgetPrefs"
 
End Property


'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : beededea
' Date      : 08/05/2023
' Purpose   : calling a manual property  by val to a form in this manual property allows external checks to the form to
'             determine whether it is loaded, without also activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    If IsLoaded = True Then
        If Me.WindowState = vbNormal Then
            IsVisible = Me.Visible
        Else
            IsVisible = False
        End If
    Else
        IsVisible = False
    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Form widgetPrefs"
            Resume Next
          End If
    End With
End Property

''---------------------------------------------------------------------------------------
'' Procedure : IsVisible
'' Author    : beededea
'' Date      : 16/12/2024
'' Purpose   : property by val to manually determine whether the preference form is visible. It does this without
''             touching any VB6 intrinsic form property which would then load the form itself.
''---------------------------------------------------------------------------------------
''
'Public Property Let IsVisible(ByVal newValue As Boolean)
'
'   On Error GoTo IsVisible_Error
'
'   If mIsVisible <> newValue Then mIsVisible = newValue Else Exit Property
'
'   On Error GoTo 0
'   Exit Property
'
'IsVisible_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Form widgetPrefs"
'
'End Property

'---------------------------------------------------------------------------------------
' Procedure : showLastTab
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : make the last used tab appear on startup
'---------------------------------------------------------------------------------------
'
Private Sub showLastTab()

   On Error GoTo showLastTab_Error

    If gsLastSelectedTab = "general" Then Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton)  ' was imgGeneralMouseUpEvent
    If gsLastSelectedTab = "config" Then Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton)     ' was imgConfigMouseUpEvent
    If gsLastSelectedTab = "position" Then Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    If gsLastSelectedTab = "development" Then Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    If gsLastSelectedTab = "fonts" Then Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    If gsLastSelectedTab = "sounds" Then Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    If gsLastSelectedTab = "window" Then Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    If gsLastSelectedTab = "about" Then Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)

    '' tenShillingsOverlay.RotateBusyTimer = True
    
  On Error GoTo 0
   Exit Sub

showLastTab_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showLastTab of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : positionPrefsFramesButtons
' Author    : beededea
' Date      : 01/05/2023
' Purpose   : size and position the frames and buttons. Note we are NOT using control
'             arrays so the form can be converted to Cairo forms later.
'---------------------------------------------------------------------------------------
'
Private Sub positionPrefsFramesButtons()
    On Error GoTo positionPrefsFramesButtons_Error

    Dim frameWidth As Integer: frameWidth = 0
    Dim frameTop As Integer: frameTop = 0
    Dim frameLeft As Integer: frameLeft = 0
    Dim buttonTop As Integer:    buttonTop = 0
    Dim rightHandAlignment As Long: rightHandAlignment = 0
    Dim leftHandGutterWidth As Long: leftHandGutterWidth = 0
       
    ' constrain the height/width ratio
    gdConstraintRatio = pcPrefsFormHeight / pcPrefsFormWidth
    
    ' align frames rightmost and leftmost to the buttons at the top
    buttonTop = -15
    frameTop = 1250
    leftHandGutterWidth = 240
    frameLeft = leftHandGutterWidth ' use the first frame leftmost as reference
    rightHandAlignment = fraAboutButton.Left + fraAboutButton.Width ' use final button rightmost as reference
    frameWidth = rightHandAlignment - frameLeft
    fraScrollbarCover.Left = rightHandAlignment - 690
    
    ' widgetPrefs.Width = rightHandAlignment + leftHandGutterWidth + 75 ' (not quite sure why we need the 75 twips padding) ' this triggers a resize '
     
    ' align the top buttons
    fraGeneralButton.Top = buttonTop
    fraConfigButton.Top = buttonTop
    fraFontsButton.Top = buttonTop
    fraSoundsButton.Top = buttonTop
    fraPositionButton.Top = buttonTop
    fraDevelopmentButton.Top = buttonTop
    fraWindowButton.Top = buttonTop
    fraAboutButton.Top = buttonTop
    
    ' align the frames
    fraGeneral.Top = frameTop
    fraConfig.Top = frameTop
    fraFonts.Top = frameTop
    fraSounds.Top = frameTop
    fraPosition.Top = frameTop
    fraDevelopment.Top = frameTop
    fraWindow.Top = frameTop
    fraAbout.Top = frameTop
    
    fraGeneral.Left = frameLeft
    fraConfig.Left = frameLeft
    fraSounds.Left = frameLeft
    fraPosition.Left = frameLeft
    fraFonts.Left = frameLeft
    fraDevelopment.Left = frameLeft
    fraWindow.Left = frameLeft
    fraAbout.Left = frameLeft
    
    fraGeneral.Width = frameWidth
    fraConfig.Width = frameWidth
    fraSounds.Width = frameWidth
    fraPosition.Width = frameWidth
    fraFonts.Width = frameWidth
    fraWindow.Width = frameWidth
    fraDevelopment.Width = frameWidth
    fraAbout.Width = frameWidth
    
    ' set the base visibility of the frames
    fraGeneral.Visible = True
    fraConfig.Visible = False
    fraSounds.Visible = False
    fraPosition.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraDevelopment.Visible = False
    fraAbout.Visible = False
            
    fraGeneralButton.BorderStyle = 1
    
    #If twinbasic Then
        fraGeneralButton.Refresh
    #End If

    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - 50
    btnHelp.Left = frameLeft
    
    '' tenShillingsOverlay.RotateBusyTimer = True

   On Error GoTo 0
   Exit Sub

positionPrefsFramesButtons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure positionPrefsFramesButtons of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : btnClose_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : VB6 button to close the prefs form
'---------------------------------------------------------------------------------------
'
Private Sub btnClose_Click()
   On Error GoTo btnClose_Click_Error

    btnSave.Enabled = False ' disable the save button
    Me.Hide
    Me.themeTimer.Enabled = False
    
    Call writePrefsPositionAndSize
    
    Call adjustPrefsControls(True)

   On Error GoTo 0
   Exit Sub

btnClose_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnClose_Click of Form widgetPrefs"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : btnHelp_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : display the help file
'---------------------------------------------------------------------------------------
'
Private Sub btnHelp_Click()
    
    On Error GoTo btnHelp_Click_Error
    
        If fFExists(App.Path & "\help\Help.chm") Then
            Call ShellExecute(Me.hwnd, "Open", App.Path & "\help\Help.chm", vbNullString, App.Path, 1)
        Else
            MsgBox ("%Err-I-ErrorNumber 11 - The help file - Help.chm - is missing from the help folder.")
        End If

   On Error GoTo 0
   Exit Sub

btnHelp_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnHelp_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : BCStr
' Author    : beededea
' Date      : 31/05/2025
' Purpose   : This replacement of CStr is ONLY for non-numeric boolean casts to a string.
'             NOTE: boolean values can be locale-sensitive when converted by CStr returning a local language result
'             It might be best to stick with LTrim$(Str$()) for the moment.
'             It is OK to convert checkboxes using cstr() as the boolean values are stored as 0 and -1
'---------------------------------------------------------------------------------------
'
Function BCStr(ByVal booleanValue As Boolean) As String
   On Error GoTo BCStr_Error

    If booleanValue Then
        BCStr = "True"
    Else
        BCStr = "False"
    End If
    
   On Error GoTo 0
   Exit Function

BCStr_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BCStr of Form widgetPrefs"
End Function
'
'---------------------------------------------------------------------------------------
' Procedure : btnSave_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : save the values from all the tabs
'             NOTE: boolean values can be locale-sensitive when converted by CStr returning a local language result
'             It might be best to stick with LTrim$(Str$()) for the moment.
'             It is OK to convert checkboxes using cstr() as the boolean values are stored as 0 and -1
'---------------------------------------------------------------------------------------
'
Private Sub btnSave_Click()
    
    On Error GoTo btnSave_Click_Error

    ' configuration
    gsWidgetTooltips = CStr(optWidgetTooltips(0).Tag)
    gsPrefsTooltips = CStr(optPrefsTooltips(0).Tag)
    
    gsShowTaskbar = CStr(chkShowTaskbar.Value)
    gsShowHelp = CStr(chkShowHelp.Value)
    
    gsDpiAwareness = CStr(chkDpiAwareness.Value)
    gsWidgetSize = CStr(sliWidgetSize.Value)
    gsSkewDegrees = CStr(sliSkewDegrees.Value)
    
    gsScrollWheelDirection = CStr(cmbScrollWheelDirection.ListIndex)
        
    ' general
    gsWidgetFunctions = CStr(chkWidgetFunctions.Value)
    gsStartup = CStr(chkGenStartup.Value)
    
    
    ' sounds
    gsEnableSounds = CStr(chkEnableSounds.Value)
'    gsEnableTicks = CStr(chkEnableTicks.Value)
'    gsEnableChimes = CStr(chkEnableChimes.Value)
'    gsEnableAlarms = CStr(chkEnableAlarms.Value)
'    gsVolumeBoost = CStr(chkVolumeBoost.Value)
    
    'development
    gsDebug = CStr(cmbDebug.ListIndex)
    gsDblClickCommand = txtDblClickCommand.Text
    gsOpenFile = txtOpenFile.Text
    #If twinbasic Then
        gsDefaultTBEditor = txtDefaultEditor.Text
    #Else
        gsDefaultVB6Editor = txtDefaultEditor.Text
    #End If
    
    ' position
    gsAspectHidden = CStr(cmbAspectHidden.ListIndex)
    gsWidgetPosition = CStr(cmbWidgetPosition.ListIndex)
    gsWidgetLandscape = CStr(cmbWidgetLandscape.ListIndex)
    gsWidgetPortrait = CStr(cmbWidgetPortrait.ListIndex)
    gsLandscapeFormHoffset = txtLandscapeHoffset.Text
    gsLandscapeFormVoffset = txtLandscapeVoffset.Text
    gsPortraitHoffset = txtPortraitHoffset.Text
    gsPortraitYoffset = txtPortraitYoffset.Text
    
'    gsVLocationPercPrefValue
'    gsHLocationPercPrefValue

    ' fonts
    gsPrefsFont = txtPrefsFont.Text
    gsWidgetFont = gsPrefsFont
    
    gsDisplayScreenFont = txtDisplayScreenFont.Text
    gsDisplayScreenFontSize = txtDisplayScreenFontSize.Text
    
'    gsDisplayScreenFontSize
'    gsDisplayScreenFontItalics
'    gsDisplayScreenFontColour

    ' the sizing is not saved here again as it saved during the setting phase.
    
'    If gsDpiAwareness = "1" Then
'        gsPrefsFontSizeHighDPI = txtPrefsFontSize.Text
'    Else
'        gsPrefsFontSizeLowDPI = txtPrefsFontSize.Text
'    End If
    'gsPrefsFontItalics = txtFontSize.Text

    ' Windows
    gsWindowLevel = CStr(cmbWindowLevel.ListIndex)
    gsPreventDragging = CStr(chkPreventDragging.Value)
    gsOpacity = CStr(sliOpacity.Value)
    gsWidgetHidden = CStr(chkWidgetHidden.Value)
    gsHidingTime = CStr(cmbHidingTime.ListIndex)
    gsIgnoreMouse = CStr(chkIgnoreMouse.Value)
    gsFormVisible = CStr(chkFormVisible.Value)
    
    gsMultiMonitorResize = CStr(cmbMultiMonitorResize.ListIndex)
     
            
    If gsStartup = "1" Then
        Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "TenShillings" & gsRichClientEnvironment & "Widget" & gsCodingEnvironment, """" & App.Path & "\" & "TenShillings-" & gsRichClientEnvironment & "-Widget-" & gsCodingEnvironment & ".exe""")
    Else
        Call writeRegistry(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "TenShillings" & gsRichClientEnvironment & "Widget" & gsCodingEnvironment, vbNullString)
    End If

    ' save the values from the general tab
    If fFExists(gsSettingsFile) Then
        sPutINISetting "Software\TenShillings", "widgetTooltips", gsWidgetTooltips, gsSettingsFile
        sPutINISetting "Software\TenShillings", "prefsTooltips", gsPrefsTooltips, gsSettingsFile

        sPutINISetting "Software\TenShillings", "showTaskbar", gsShowTaskbar, gsSettingsFile
        sPutINISetting "Software\TenShillings", "showHelp", gsShowHelp, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "dpiAwareness", gsDpiAwareness, gsSettingsFile
        
        
        sPutINISetting "Software\TenShillings", "widgetSize", gsWidgetSize, gsSettingsFile
        sPutINISetting "Software\TenShillings", "scrollWheelDirection", gsScrollWheelDirection, gsSettingsFile
                
        sPutINISetting "Software\TenShillings", "widgetFunctions", gsWidgetFunctions, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "pointerAnimate", gsPointerAnimate, gsSettingsFile
        sPutINISetting "Software\TenShillings", "samplingInterval", gsSamplingInterval, gsSettingsFile
        
              
        sPutINISetting "Software\TenShillings", "aspectHidden", gsAspectHidden, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetPosition", gsWidgetPosition, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetLandscape", gsWidgetLandscape, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetPortrait", gsWidgetPortrait, gsSettingsFile

        sPutINISetting "Software\TenShillings", "prefsFont", gsPrefsFont, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetFont", gsWidgetFont, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "prefsFontSizeHighDPI", gsPrefsFontSizeHighDPI, gsSettingsFile
        sPutINISetting "Software\TenShillings", "prefsFontSizeLowDPI", gsPrefsFontSizeLowDPI, gsSettingsFile
        sPutINISetting "Software\TenShillings", "prefsFontItalics", gsPrefsFontItalics, gsSettingsFile
        sPutINISetting "Software\TenShillings", "prefsFontColour", gsPrefsFontColour, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "displayScreenFont", gsDisplayScreenFont, gsSettingsFile
        sPutINISetting "Software\TenShillings", "displayScreenFontSize", gsDisplayScreenFontSize, gsSettingsFile
        sPutINISetting "Software\TenShillings", "displayScreenFontItalics", gsDisplayScreenFontItalics, gsSettingsFile
        sPutINISetting "Software\TenShillings", "displayScreenFontColour", gsDisplayScreenFontColour, gsSettingsFile

        'save the values from the Windows Config Items
        sPutINISetting "Software\TenShillings", "windowLevel", gsWindowLevel, gsSettingsFile
        sPutINISetting "Software\TenShillings", "preventDragging", gsPreventDragging, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "opacity", gsOpacity, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetHidden", gsWidgetHidden, gsSettingsFile
        sPutINISetting "Software\TenShillings", "hidingTime", gsHidingTime, gsSettingsFile
        sPutINISetting "Software\TenShillings", "ignoreMouse", gsIgnoreMouse, gsSettingsFile
        sPutINISetting "Software\TenShillings", "formVisible", gsFormVisible, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "multiMonitorResize", gsMultiMonitorResize, gsSettingsFile
        
        
        sPutINISetting "Software\TenShillings", "startup", gsStartup, gsSettingsFile

        sPutINISetting "Software\TenShillings", "enableSounds", gsEnableSounds, gsSettingsFile
'        sPutINISetting "Software\TenShillings", "enableTicks", gsEnableTicks, gsSettingsFile
'        sPutINISetting "Software\TenShillings", "enableChimes", gsEnableChimes, gsSettingsFile
'        sPutINISetting "Software\TenShillings", "enableAlarms", gsEnableAlarms, gsSettingsFile
'        sPutINISetting "Software\TenShillings", "volumeBoost", gsVolumeBoost, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "lastSelectedTab", gsLastSelectedTab, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "debug", gsDebug, gsSettingsFile
        sPutINISetting "Software\TenShillings", "dblClickCommand", gsDblClickCommand, gsSettingsFile
        sPutINISetting "Software\TenShillings", "openFile", gsOpenFile, gsSettingsFile
        sPutINISetting "Software\TenShillings", "defaultVB6Editor", gsDefaultVB6Editor, gsSettingsFile
        sPutINISetting "Software\TenShillings", "defaultTBEditor", gsDefaultTBEditor, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "widgetHighDpiXPos", gsWidgetHighDpiXPos, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetHighDpiYPos", gsWidgetHighDpiYPos, gsSettingsFile
        
        sPutINISetting "Software\TenShillings", "widgetLowDpiXPos", gsWidgetLowDpiXPos, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetLowDpiYPos", gsWidgetLowDpiYPos, gsSettingsFile
                       
    End If
    
    ' saves the current ratios for the RC form and the absolute sizes for the Prefs form
    Call saveMainFormsCurrentSizeAndRatios
    
    ' set the tooltips on the prefs screen
    Call setPrefsTooltips

    ' sets the characteristics of the widget and menus immediately after saving
    Call adjustMainControls(1)
    
    If widgetPrefs.IsVisible Then Me.SetFocus
    
    btnSave.Enabled = False ' disable the save button showing it has successfully saved
    
    ' reload here if the gsWindowLevel Was Changed
    If gbWindowLevelWasChanged = True Then
        gbWindowLevelWasChanged = False
        Call reloadProgram
    End If
    
   On Error GoTo 0
   Exit Sub

btnSave_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSave_Click of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkEnableSounds_Click
' Author    : beededea
' Date      : 18/02/2025
' Purpose   : toggle to enable/disable sounds on save
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableSounds_Click()
   On Error GoTo chkEnableSounds_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

chkEnableSounds_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableSounds_Click of Form widgetPrefs"
End Sub


' ----------------------------------------------------------------
' Procedure Name: cmbWindowLevel_Click
' Purpose: option to determine the windows Z order of the main program (not the prefs form)
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 28/05/2024
' ----------------------------------------------------------------
Private Sub cmbWindowLevel_Click()
    On Error GoTo cmbWindowLevel_Click_Error
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    If pPrefsStartupFlg = False Then gbWindowLevelWasChanged = True
    
    On Error GoTo 0
    Exit Sub

cmbWindowLevel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbWindowLevel_Click, line " & Erl & "."

End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnPrefsFont_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : VB6 button to select the font dialog
'---------------------------------------------------------------------------------------
'
Private Sub btnPrefsFont_Click()

    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    On Error GoTo btnPrefsFont_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    
    ' set the preliminary vars to feed and populate the changefont routine
    fntFont = gsPrefsFont
    ' gsWidgetFont
    
    If gsDpiAwareness = "1" Then
        fntSize = Val(gsPrefsFontSizeHighDPI)
    Else
        fntSize = Val(gsPrefsFontSizeLowDPI)
    End If
    
    If fntSize = 0 Then fntSize = 8
    fntItalics = CBool(gsPrefsFontItalics)
    fntColour = CLng(gsPrefsFontColour)
        
    Call changeFont(widgetPrefs, True, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult)
    
    gsPrefsFont = CStr(fntFont)
    gsWidgetFont = gsPrefsFont
    
    If gsDpiAwareness = "1" Then
        gsPrefsFontSizeHighDPI = CStr(fntSize)
        Call Form_Resize
    Else
        gsPrefsFontSizeLowDPI = CStr(fntSize)
    End If
    
    gsPrefsFontItalics = CStr(fntItalics)
    gsPrefsFontColour = CStr(fntColour)
    
    ' changes the displayed font to an adjusted base font size after a resize
    Call PrefsFormResizeEvent

    If fFExists(gsSettingsFile) Then ' does the tool's own settings.ini exist?
        sPutINISetting "Software\TenShillings", "prefsFont", gsPrefsFont, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetFont", gsWidgetFont, gsSettingsFile
        sPutINISetting "Software\TenShillings", "prefsFontSizeHighDPI", gsPrefsFontSizeHighDPI, gsSettingsFile
        sPutINISetting "Software\TenShillings", "prefsFontSizeLowDPI", gsPrefsFontSizeLowDPI, gsSettingsFile
        sPutINISetting "Software\TenShillings", "prefsFontItalics", gsPrefsFontItalics, gsSettingsFile
        sPutINISetting "Software\TenShillings", "PrefsFontColour", gsPrefsFontColour, gsSettingsFile
    End If
    
    If fntFont = vbNullString Then fntFont = "arial"
    txtPrefsFont.Text = fntFont
    txtPrefsFont.Font.Name = fntFont
    'txtPrefsFont.Font.Size = fntSize
    txtPrefsFont.Font.Italic = fntItalics
    txtPrefsFont.ForeColor = fntColour
    
    txtPrefsFontSize.Text = fntSize

   On Error GoTo 0
   Exit Sub

btnPrefsFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnPrefsFont_Click of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : btnDisplayScreenFont_Click
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : VB6 button to select the font dialog for the display console
'---------------------------------------------------------------------------------------
'
Private Sub btnDisplayScreenFont_Click()

    Dim fntFont As String: fntFont = vbNullString
    Dim fntSize As Integer: fntSize = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim fntColour As Long: fntColour = 0
    Dim fntItalics As Boolean: fntItalics = False
    Dim fntUnderline As Boolean: fntUnderline = False
    Dim fntFontResult As Boolean: fntFontResult = False
    
    On Error GoTo btnDisplayScreenFont_Click_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    
    ' set the preliminary vars to feed and populate the changefont routine
    fntFont = gsDisplayScreenFont
    
    fntSize = Val(gsDisplayScreenFontSize)
    If fntSize = 0 Then fntSize = 5
    fntItalics = CBool(gsDisplayScreenFontItalics)
    fntColour = CLng(gsDisplayScreenFontColour)
    
    displayFontSelector fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult
    If fntFontResult = False Then Exit Sub
            
    gsDisplayScreenFont = CStr(fntFont)
    gsDisplayScreenFontSize = CStr(fntSize)
    gsDisplayScreenFontItalics = CStr(fntItalics)
    gsDisplayScreenFontColour = CStr(fntColour)
    
    If gbThisWidgetAvailable = True Then
        fMain.TenShillingsForm.Widgets("lblTerminalText").Widget.FontSize = gsDisplayScreenFontSize
        fMain.TenShillingsForm.Widgets("lblTerminalText").Widget.FontName = gsDisplayScreenFont
    End If

    If fFExists(gsSettingsFile) Then ' does the tool's own settings.ini exist?
        sPutINISetting "Software\TenShillings", "displayScreenFont", gsDisplayScreenFont, gsSettingsFile
        sPutINISetting "Software\TenShillings", "displayScreenFontSize", gsDisplayScreenFontSize, gsSettingsFile
        sPutINISetting "Software\TenShillings", "displayScreenFontItalics", gsDisplayScreenFontItalics, gsSettingsFile
        sPutINISetting "Software\TenShillings", "displayScreenFontColour", gsDisplayScreenFontColour, gsSettingsFile
    End If
    
    If fntFont = vbNullString Then fntFont = "courier new"
    txtDisplayScreenFont.Text = fntFont
    txtDisplayScreenFont.Font.Name = fntFont
    txtDisplayScreenFont.Font.Italic = fntItalics
    txtDisplayScreenFont.ForeColor = fntColour
    txtDisplayScreenFontSize.Text = fntSize

   On Error GoTo 0
   Exit Sub

btnDisplayScreenFont_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDisplayScreenFont_Click of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsControls
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : adjust the controls so their startup position matches the last write of the config file
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsControls(Optional ByVal restartState As Boolean)
    ' Dim I As Integer: I = 0
    Dim fntWeight As Integer: fntWeight = 0
    Dim fntStyle As Boolean: fntStyle = False
    Dim sliWidgetSizeOldValue As Long: sliWidgetSizeOldValue = 0
    'Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
    
    On Error GoTo adjustPrefsControls_Error
    
    ' note the monitor ID at PrefsForm form_load and store as the prefsFormMonitorID
    'gPrefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
    
    'widgetPrefs.Height = CLng(gsPrefsPrimaryHeightTwips)
            
    ' general tab
    chkWidgetFunctions.Value = Val(gsWidgetFunctions)
    chkGenStartup.Value = Val(gsStartup)
            
     ' check whether the size has been previously altered via ctrl+mousewheel on the widget
    sliWidgetSizeOldValue = sliWidgetSize.Value
    sliWidgetSize.Value = Val(gsWidgetSize)
    If sliWidgetSize.Value <> sliWidgetSizeOldValue Then
        btnSave.Visible = True
    End If
    
    sliSkewDegrees.Value = Val(gsSkewDegrees)
        
    cmbScrollWheelDirection.ListIndex = Val(gsScrollWheelDirection)
    
    optWidgetTooltips(CStr(gsWidgetTooltips)).Value = True
    optWidgetTooltips(0).Tag = CStr(gsWidgetTooltips)
    optWidgetTooltips(1).Tag = CStr(gsWidgetTooltips)
    optWidgetTooltips(2).Tag = CStr(gsWidgetTooltips)
        
    optPrefsTooltips(CStr(gsPrefsTooltips)).Value = True
    optPrefsTooltips(0).Tag = CStr(gsPrefsTooltips)
    optPrefsTooltips(1).Tag = CStr(gsPrefsTooltips)
    optPrefsTooltips(2).Tag = CStr(gsPrefsTooltips)
    
    chkShowTaskbar.Value = Val(gsShowTaskbar)
    chkShowHelp.Value = Val(gsShowHelp)
    chkDpiAwareness.Value = Val(gsDpiAwareness)

    ' sounds tab
    chkEnableSounds.Value = Val(gsEnableSounds)
    
    ' development
    cmbDebug.ListIndex = Val(gsDebug)
    txtDblClickCommand.Text = gsDblClickCommand
    txtOpenFile.Text = gsOpenFile
    #If twinbasic Then
        txtDefaultEditor.Text = gsDefaultTBEditor
    #Else
        txtDefaultEditor.Text = gsDefaultVB6Editor
    #End If
    
    lblGitHub.Caption = "You can find the code for the TenShillings Widget on github, visit by double-clicking this link https://github.com/yereverluvinunclebert/TenShillings-" & gsRichClientEnvironment & "-Widget-" & gsCodingEnvironment
     
     If Not restartState = True Then
        ' fonts tab
        If gsPrefsFont <> vbNullString Then
            txtPrefsFont.Text = gsPrefsFont
            If gsDpiAwareness = "1" Then
                Call changeFormFont(widgetPrefs, gsPrefsFont, Val(gsPrefsFontSizeHighDPI), fntWeight, fntStyle, gsPrefsFontItalics, gsPrefsFontColour)
                txtPrefsFontSize.Text = gsPrefsFontSizeHighDPI
            Else
                Call changeFormFont(widgetPrefs, gsPrefsFont, Val(gsPrefsFontSizeLowDPI), fntWeight, fntStyle, gsPrefsFontItalics, gsPrefsFontColour)
                txtPrefsFontSize.Text = gsPrefsFontSizeLowDPI
            End If
        End If
        
        txtDisplayScreenFontSize.Text = gsDisplayScreenFontSize
    
        txtDisplayScreenFont.Font.Name = gsDisplayScreenFont
        'txtDisplayScreenFont.Font.Size = Val(gsDisplayScreenFont)
    End If
    
    ' position tab
    
    
    cmbAspectHidden.ListIndex = Val(gsAspectHidden)
    cmbWidgetPosition.ListIndex = Val(gsWidgetPosition)
        
    If gsPreventDragging = "1" Then
        If gsAspectRatio = "landscape" Then
'            txtLandscapeHoffset.Text = fMain.TenShillingsForm.Left
'            txtLandscapeVoffset.Text = fMain.TenShillingsForm.Top
            If gsDpiAwareness = "1" Then
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gsWidgetHighDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gsWidgetHighDpiYPos & "px"
            Else
                txtLandscapeHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gsWidgetLowDpiXPos & "px"
                txtLandscapeVoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gsWidgetLowDpiYPos & "px"
            End If
        Else
'            txtPortraitHoffset.Text = fMain.TenShillingsForm.Left
'            txtPortraitYoffset.Text = fMain.TenShillingsForm.Top
            If gsDpiAwareness = "1" Then
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gsWidgetHighDpiXPos & "px"
                txtPortraitYoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gsWidgetHighDpiYPos & "px"
            Else
                txtPortraitHoffset.ToolTipText = "Last Sampled Form X Horizontal Position : " & gsWidgetLowDpiXPos & "px"
                txtPortraitYoffset.ToolTipText = "Last Sampled Form Y Vertical Position : " & gsWidgetLowDpiYPos & "px"
            End If
        End If
    End If
    
    'cmbWidgetLandscape
    
    cmbWidgetLandscape.ListIndex = Val(gsWidgetLandscape)
    cmbWidgetPortrait.ListIndex = Val(gsWidgetPortrait)
    txtLandscapeHoffset.Text = gsLandscapeFormHoffset
    txtLandscapeVoffset.Text = gsLandscapeFormVoffset
    txtPortraitHoffset.Text = gsPortraitHoffset
    txtPortraitYoffset.Text = gsPortraitYoffset

    ' Windows tab
    
    cmbWindowLevel.ListIndex = Val(gsWindowLevel)
    chkIgnoreMouse.Value = Val(gsIgnoreMouse)
    chkFormVisible.Value = Val(gsFormVisible)
    
    chkPreventDragging.Value = Val(gsPreventDragging)
    sliOpacity.Value = Val(gsOpacity)
    chkWidgetHidden.Value = Val(gsWidgetHidden)
    cmbHidingTime.ListIndex = Val(gsHidingTime)
    cmbMultiMonitorResize.ListIndex = Val(gsMultiMonitorResize)
    
    If glMonitorCount > 1 Then
        cmbMultiMonitorResize.Visible = True
        lblWindowLevel(10).Visible = True
        lblWindowLevel(11).Visible = True
    Else
        cmbMultiMonitorResize.Visible = False
        lblWindowLevel(10).Visible = False
        lblWindowLevel(11).Visible = False
    End If
        
   On Error GoTo 0
   Exit Sub

adjustPrefsControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsControls of Form widgetPrefs on line " & Erl

End Sub

'---------------------------------------------------------------------------------------

'
'---------------------------------------------------------------------------------------
' Procedure : populatePrefsComboBoxes
' Author    : beededea
' Date      : 10/09/2022
' Purpose   : all combo boxes in the prefs are populated here with default values
'           : done by preference here rather than in the IDE
'---------------------------------------------------------------------------------------

Private Sub populatePrefsComboBoxes()

    'Dim I As Integer: I = 0
    
    On Error GoTo populatePrefsComboBoxes_Error
    
    cmbScrollWheelDirection.AddItem "up", 0
    cmbScrollWheelDirection.ItemData(0) = 0
    cmbScrollWheelDirection.AddItem "down", 1
    cmbScrollWheelDirection.ItemData(1) = 1
        
    cmbAspectHidden.AddItem "none", 0
    cmbAspectHidden.ItemData(0) = 0
    cmbAspectHidden.AddItem "portrait", 1
    cmbAspectHidden.ItemData(1) = 1
    cmbAspectHidden.AddItem "landscape", 2
    cmbAspectHidden.ItemData(2) = 2
    
    cmbWidgetPosition.AddItem "disabled", 0
    cmbWidgetPosition.ItemData(0) = 0
    cmbWidgetPosition.AddItem "enabled", 1
    cmbWidgetPosition.ItemData(1) = 1
    
    cmbWidgetLandscape.AddItem "disabled", 0
    cmbWidgetLandscape.ItemData(0) = 0
    cmbWidgetLandscape.AddItem "enabled", 1
    cmbWidgetLandscape.ItemData(1) = 1
    
    cmbWidgetPortrait.AddItem "disabled", 0
    cmbWidgetPortrait.ItemData(0) = 0
    cmbWidgetPortrait.AddItem "enabled", 1
    cmbWidgetPortrait.ItemData(1) = 1
    
    cmbDebug.AddItem "Debug OFF", 0
    cmbDebug.ItemData(0) = 0
    cmbDebug.AddItem "Debug ON", 1
    cmbDebug.ItemData(1) = 1
        
    ' populate comboboxes in the windows tab
    cmbWindowLevel.AddItem "Keep on top of other windows", 0
    cmbWindowLevel.ItemData(0) = 0
    cmbWindowLevel.AddItem "Normal", 0
    cmbWindowLevel.ItemData(1) = 1
    cmbWindowLevel.AddItem "Keep below all other windows", 0
    cmbWindowLevel.ItemData(2) = 2
    
    ' populate the hiding timer combobox
    cmbHidingTime.AddItem "1 minute", 0
    cmbHidingTime.ItemData(0) = 1
    cmbHidingTime.AddItem "5 minutes", 1
    cmbHidingTime.ItemData(1) = 5
    cmbHidingTime.AddItem "10 minutes", 2
    cmbHidingTime.ItemData(2) = 10
    cmbHidingTime.AddItem "20 minutes", 3
    cmbHidingTime.ItemData(3) = 20
    cmbHidingTime.AddItem "30 minutes", 4
    cmbHidingTime.ItemData(4) = 30
    cmbHidingTime.AddItem "I hour", 5
    cmbHidingTime.ItemData(5) = 60
    
    ' populate the multi monitor combobox
    cmbMultiMonitorResize.AddItem "Disabled", 0
    cmbMultiMonitorResize.ItemData(0) = 0
    cmbMultiMonitorResize.AddItem "Automatic Resizing Enabled", 1
    cmbMultiMonitorResize.ItemData(1) = 1
    cmbMultiMonitorResize.AddItem "Manual Sizing Stored Per Monitor", 2
    cmbMultiMonitorResize.ItemData(2) = 2

    On Error GoTo 0
    Exit Sub

populatePrefsComboBoxes_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure populatePrefsComboBoxes of Form widgetPrefs"
            Resume Next
          End If
    End With
                
End Sub

'
''---------------------------------------------------------------------------------------
'' Procedure : readFileWriteComboBox
'' Author    : beededea
'' Date      : 28/07/2023
'' Purpose   : Open and load the Array with the timezones text File
''---------------------------------------------------------------------------------------
''
'Private Sub readFileWriteComboBox(ByRef thisComboBox As Control, ByVal thisFileName As String)
'    Dim strArr() As String
'    Dim lngCount As Long: lngCount = 0
'    Dim lngIdx As Long: lngIdx = 0
'
'    On Error GoTo readFileWriteComboBox_Error
'
'    If fFExists(thisFileName) = True Then
'       ' the files must be DOS CRLF delineated
'       Open thisFileName For Input As #1
'           strArr() = Split(Input(LOF(1), 1), vbCrLf)
'       Close #1
'
'       lngCount = UBound(strArr)
'
'       thisComboBox.Clear
'       For lngIdx = 0 To lngCount
'           thisComboBox.AddItem strArr(lngIdx)
'       Next lngIdx
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'readFileWriteComboBox_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readFileWriteComboBox of Form widgetPrefs"

'End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : clearBorderStyle
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : removes all styling from the icon frames and makes the major frames below invisible too, not using control arrays.
'---------------------------------------------------------------------------------------
'
Private Sub clearBorderStyle()

   On Error GoTo clearBorderStyle_Error

    fraGeneral.Visible = False
    fraConfig.Visible = False
    fraFonts.Visible = False
    fraWindow.Visible = False
    fraPosition.Visible = False
    fraDevelopment.Visible = False
    fraSounds.Visible = False
    fraAbout.Visible = False

    fraGeneralButton.BorderStyle = 0
    fraConfigButton.BorderStyle = 0
    fraDevelopmentButton.BorderStyle = 0
    fraPositionButton.BorderStyle = 0
    fraFontsButton.BorderStyle = 0
    fraWindowButton.BorderStyle = 0
    fraSoundsButton.BorderStyle = 0
    fraAboutButton.BorderStyle = 0
    
    #If twinbasic Then
        fraGeneralButton.Refresh
        fraConfigButton.Refresh
        fraDevelopmentButton.Refresh
        fraPositionButton.Refresh
        fraFontsButton.Refresh
        fraWindowButton.Refresh
        fraSoundsButton.Refresh
        fraAboutButton.Refresh
    #End If

   On Error GoTo 0
   Exit Sub

clearBorderStyle_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clearBorderStyle of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Resize
' Author    : beededea
' Date      : 30/05/2023
' Purpose   : IMPORTANT: Called at every twip of resising, goodness knows what interval, we barely use this, instead we subclass and look for WM_EXITSIZEMOVE
'---------------------------------------------------------------------------------------
'
Private Sub Form_Resize()

    On Error GoTo Form_Resize_Error

    pPrefsFormResizedByDrag = True
         
    ' do not call the resizing function when the form is resized by dragging the border
    ' only call this if the resize is done in code
        
    If InIde Or gbPrefsFormResizedInCode = True Then
        Call PrefsFormResizeEvent
        If gbPrefsFormResizedInCode = True Then Exit Sub
    End If
    
    ' when resizing the form enable the save button to allow the recently set width/height to be saved.
    If glMonitorCount > 1 And Val(gsMultiMonitorResize) > 0 And widgetPrefs.IsVisible = True Then
        btnSave.Enabled = True
    End If
            
    On Error GoTo 0
    Exit Sub

Form_Resize_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form widgetPrefs"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : PrefsFormResizeEvent
' Author    : beededea
' Date      : 10/10/2024
' Purpose   : Called mostly by WM_EXITSIZEMOVE, the subclassed (intercepted) message that indicates that the window has just been moved.
'             (and on a mouseUp during a bottom-right drag of the additional corner indicator). Also, in code as specifcally required with an indicator flag.
'             This prevents a resize occurring during every twip movement and the controls resizing themselves continuously.
'             They now only resize when the form resize has completed.
'
'---------------------------------------------------------------------------------------
'
Public Sub PrefsFormResizeEvent()

    Dim currentFontSize As Long: currentFontSize = 0
    
    On Error GoTo PrefsFormResizeEvent_Error

    ' When minimised and a resize is called then simply exit.
    If Me.WindowState = vbMinimized Then Exit Sub
        
    If pPrefsDynamicSizingFlg = True And pPrefsFormResizedByDrag = True Then
    
        widgetPrefs.Width = widgetPrefs.height / gdConstraintRatio ' maintain the aspect ratio, note: this change calls this routine again...
        
        If gsDpiAwareness = "1" Then
            currentFontSize = gsPrefsFontSizeHighDPI
        Else
            currentFontSize = gsPrefsFontSizeLowDPI
        End If

        'make tab frames invisible so that the control resizing is not apparent to the user
        Call makeFramesInvisible
        Call resizeControls(widgetPrefs, gcPrefsControlPositions(), gdPrefsStartWidth, gdPrefsStartHeight, currentFontSize)

        Call tweakgcPrefsControlPositions(Me, gdPrefsStartWidth, gdPrefsStartHeight)
        'Call loadHigherResPrefsImages ' if you want higher res icons then load them here, current max. is 1010 twips or 67 pixels
        Call makeFramesVisible
        
    Else
        If Me.WindowState = 0 Then ' normal
            If widgetPrefs.Width > 9090 Then widgetPrefs.Width = 9090
            If widgetPrefs.Width < 9085 Then widgetPrefs.Width = 9090
            If pLastFormHeight <> 0 Then
               gbPrefsFormResizedInCode = True
               widgetPrefs.height = pLastFormHeight
            End If
        End If
    End If
    
    gbPrefsFormResizedInCode = False
    pPrefsFormResizedByDrag = False

   On Error GoTo 0
   Exit Sub

PrefsFormResizeEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrefsFormResizeEvent of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : makeFramesInvisible
' Author    : beededea
' Date      : 23/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub makeFramesInvisible()
    
   On Error GoTo makeFramesInvisible_Error

        If gsLastSelectedTab = "general" Then
            fraGeneral.Visible = False
            fraGeneralButton.Visible = False
        End If
        If gsLastSelectedTab = "config" Then
            fraConfig.Visible = False
            fraConfigButton.Visible = False
        End If
        If gsLastSelectedTab = "position" Then
            fraPosition.Visible = False
            fraPositionButton.Visible = False
        End If
            
        If gsLastSelectedTab = "development" Then
            fraDevelopment.Visible = False
            fraDevelopmentButton.Visible = False
        End If

        If gsLastSelectedTab = "fonts" Then
            fraFonts.Visible = False
            fraFontsButton.Visible = False
        End If

        If gsLastSelectedTab = "sounds" Then
            fraSounds.Visible = False
            fraSoundsButton.Visible = False
        End If

        If gsLastSelectedTab = "window" Then
            fraWindow.Visible = False
            fraWindowButton.Visible = False
        End If

        If gsLastSelectedTab = "about" Then
            fraAbout.Visible = False
            fraAboutButton.Visible = False
        End If
        
        fraIconGroup.Visible = False

   On Error GoTo 0
   Exit Sub

makeFramesInvisible_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeFramesInvisible of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : makeFramesVisible
' Author    : beededea
' Date      : 23/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub makeFramesVisible()
    
   On Error GoTo makeFramesVisible_Error

    If gsLastSelectedTab = "general" Then
        fraGeneral.Visible = True
        fraGeneralButton.Visible = True
    End If
    If gsLastSelectedTab = "config" Then
        fraConfig.Visible = True
        fraConfigButton.Visible = True
    End If
    If gsLastSelectedTab = "position" Then
        fraPosition.Visible = True
        fraPositionButton.Visible = True
    End If
        
    If gsLastSelectedTab = "development" Then
        fraDevelopment.Visible = True
        fraDevelopmentButton.Visible = True
    End If

    If gsLastSelectedTab = "fonts" Then
        fraFonts.Visible = True
        fraFontsButton.Visible = True
    End If

    If gsLastSelectedTab = "sounds" Then
        fraSounds.Visible = True
        fraSoundsButton.Visible = True
    End If

    If gsLastSelectedTab = "window" Then
        fraWindow.Visible = True
        fraWindowButton.Visible = True
    End If

    If gsLastSelectedTab = "about" Then
        fraAbout.Visible = True
        fraAboutButton.Visible = True
    End If
    
    fraIconGroup.Visible = True

   On Error GoTo 0
   Exit Sub

makeFramesVisible_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeFramesVisible of Form widgetPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : FormResizedOrMoved
' Author    : beededea
' Date      : 16/07/2024
' Purpose   : Non VB6-standard event caught by subclassing and intercepting the WM_EXITSIZEMOVE (WM_MOVED) event
'             If the user drags the corner then this routine is called
'             If the user drags the corner then this routine is called
'---------------------------------------------------------------------------------------
'
Public Sub FormResizedOrMoved(sForm As String)

    On Error GoTo FormResizedOrMoved_Error
        
    'passing a form name as it allows us to potentially subclass another form's movement
    Select Case sForm
        Case "widgetPrefs"
            ' call a resize of all controls only when the form resize (by dragging) has completed (mouseUP)
            If pPrefsFormResizedByDrag = True Then
            
                ' test the current form height and width, if the same then it is a Form Moved on the same monitor and not a form_resize.
                If widgetPrefs.height = glWidgetPrefsOldHeightTwips And widgetPrefs.Width = glWidgetPrefsOldWidthTwips Then
                    Exit Sub
                Else
                
                    glWidgetPrefsOldHeightTwips = widgetPrefs.height
                    glWidgetPrefsOldWidthTwips = widgetPrefs.Width
                    
                    Call PrefsFormResizeEvent
                    pPrefsFormResizedByDrag = False
                    
                End If
            Else
                ' call the procedure to resize the form automatically if it now resides on a different sized monitor
                Call positionPrefsByMonitorSize
                widgetPrefs.btnSave.Enabled = False
            End If
            
        Case "frmMessage"
            MsgBox " FORM RESIZED"
        Case Else
        
    End Select
    
   On Error GoTo 0
   Exit Sub

FormResizedOrMoved_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FormResizedOrMoved of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tweakgcPrefsControlPositions
' Author    : beededea
' Date      : 22/09/2023
' Purpose   : final tweak the bottom frame top and left positions
'---------------------------------------------------------------------------------------
'
Private Sub tweakgcPrefsControlPositions(ByVal thisForm As Form, ByVal m_FormWid As Single, ByVal m_FormHgt As Single)

    ' not sure why but the resizeControls routine can lead to incorrect positioning of frames and buttons
    Dim x_scale As Single: x_scale = 0
    Dim y_scale As Single: y_scale = 0
    
    On Error GoTo tweakgcPrefsControlPositions_Error

    ' Get the form's current scale factors.
    x_scale = thisForm.ScaleWidth / m_FormWid
    y_scale = thisForm.ScaleHeight / m_FormHgt

    fraGeneral.Left = fraGeneralButton.Left
    fraConfig.Left = fraGeneralButton.Left
    fraSounds.Left = fraGeneralButton.Left
    fraPosition.Left = fraGeneralButton.Left
    fraFonts.Left = fraGeneralButton.Left
    fraDevelopment.Left = fraGeneralButton.Left
    fraWindow.Left = fraGeneralButton.Left
    fraAbout.Left = fraGeneralButton.Left
         
    'fraGeneral.Top = fraGeneralButton.Top
    fraConfig.Top = fraGeneral.Top
    fraSounds.Top = fraGeneral.Top
    fraPosition.Top = fraGeneral.Top
    fraFonts.Top = fraGeneral.Top
    fraDevelopment.Top = fraGeneral.Top
    fraWindow.Top = fraGeneral.Top
    fraAbout.Top = fraGeneral.Top
    
    ' final tweak the bottom button positions
    
    btnHelp.Top = fraGeneral.Top + fraGeneral.height + (100 * y_scale)
    btnSave.Top = btnHelp.Top
    btnClose.Top = btnHelp.Top
    
    btnClose.Left = fraWindow.Left + fraWindow.Width - btnClose.Width
    btnSave.Left = btnClose.Left - btnSave.Width - (150 * x_scale)
    btnHelp.Left = fraGeneral.Left

    txtPrefsFontCurrentSize.Text = y_scale * txtPrefsFontCurrentSize.FontSize
    
    lblAsterix.Top = btnSave.Top + 50
    
   On Error GoTo 0
   Exit Sub

tweakgcPrefsControlPositions_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tweakgcPrefsControlPositions of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : beededea
' Date      : 27/05/2023
' Purpose   : standard form unload
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error
        
    ' Release the subclass hook for dialog forms
    If Not InIde Then ReleaseHook
    
    IsLoaded = False
    
    Call writePrefsPositionAndSize
    
    Call DestroyToolTip

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : various _MouseMove
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : setting the balloon tooltip text for several controls
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : optWidgetTooltips_MouseMove
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : setting the tooltip text for the specific radio button for selecting the widget/cal tooltip style
'---------------------------------------------------------------------------------------
'
Private Sub optWidgetTooltips_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim thisToolTip As String: thisToolTip = vbNullString
    On Error GoTo optWidgetTooltips_MouseMove_Error

    If gsPrefsTooltips = "0" Then
        If Index = 0 Then
            thisToolTip = "This setting enables the balloon tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips, note that their font size will match the Windows system font size."
            CreateToolTip optWidgetTooltips(Index).hwnd, thisToolTip, _
                  TTIconInfo, "Help on Balloon Tooltips on the GUI", , , , True
        ElseIf Index = 1 Then
            thisToolTip = "This setting enables the RichClient square tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips."
            CreateToolTip optWidgetTooltips(Index).hwnd, thisToolTip, _
                  TTIconInfo, "Help on RichClient Tooltips on the GUI", , , , True
        ElseIf Index = 2 Then
            thisToolTip = "This setting disables the balloon tooltips for elements within the Steampunk GUI."
            CreateToolTip optWidgetTooltips(Index).hwnd, thisToolTip, _
                  TTIconInfo, "Help on Disabling Tooltips on the GUI", , , , True
        End If
    
    End If

   On Error GoTo 0
   Exit Sub

optWidgetTooltips_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optWidgetTooltips_MouseMove of Form widgetPrefs"
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : btnResetMessages_MouseMove
' Author    : beededea
' Date      : 01/10/2023
' Purpose   : reset message boxes mouseOver
'---------------------------------------------------------------------------------------
'
Private Sub btnResetMessages_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo btnResetMessages_MouseMove_Error

    If gsPrefsTooltips = "0" Then CreateToolTip btnResetMessages.hwnd, "The various pop-up messages that this program generates can be manually hidden. This button restores them to their original visible state.", _
                  TTIconInfo, "Help on the message reset button", , , , True

    On Error GoTo 0
    Exit Sub

btnResetMessages_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnResetMessages_MouseMove of Form widgetPrefs"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = True
End Sub

Private Sub chkEnableResizing_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkEnableResizing.hwnd, "This allows you to resize the whole prefs window by dragging the bottom right corner of the window. It provides an alternative method of supporting high DPI screens.", _
                  TTIconInfo, "Help on Resizing", , , , True
End Sub





''---------------------------------------------------------------------------------------
'' Procedure : sliSkewDegrees_Change
'' Author    : beededea
'' Date      : 06/09/2025
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Private Sub sliSkewDegrees_Change()
'
'    On Error GoTo sliSkewDegrees_Change_Error
'
'    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
'
'    ' pAllowSkewChangeFlg prevents the slider altering the skew on the tenshillings widget unless the slider is the active user control.
'    ' It does this in order to avoid reciprocal back and forth between the main widget form and the prefs form,
'
'    ' ie. the slider slides but does not itself change the widget rotation, that is happening elsewhere
'    ' the widget rotation is modified in W_MouseWheel in cwTenShillingsOverlay
'
'    ' the pAllowSkewChangeFlg is only set when this control receives focus and the flag is unset when the user selects any another control.
'
'    If pAllowSkewChangeFlg = True Then
'        tenShillingsOverlay.SkewDegrees = sliSkewDegrees.Value
'    End If
'
'    Call saveMainRCFormSize
'
'    On Error GoTo 0
'    Exit Sub
'
'sliSkewDegrees_Change_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliSkewDegrees_Change of Form widgetPrefs"
'End Sub


Private Sub sliSkewDegrees_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip sliSkewDegrees.hwnd, "Adjust to rotate the whole widget. Any adjustment in skew made here takes place instantly (you can also use the Mousewheel when hovering over the widget itself).", _
                  TTIconInfo, "Help on the Size Rotate Slider", , , , True
End Sub





'---------------------------------------------------------------------------------------
' Procedure : sliSkewDegrees_Scroll
' Author    : beededea
' Date      : 10/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliSkewDegrees_Scroll()
    On Error GoTo sliSkewDegrees_Scroll_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

    ' pAllowSkewChangeFlg prevents the slider altering the skew on the tenshillings widget unless the slider is the active user control.
    ' It does this in order to avoid reciprocal back and forth between the main widget form and the prefs form,

    ' ie. the slider slides but does not itself change the widget rotation, that is happening elsewhere
    ' the widget rotation is modified in W_MouseWheel in cwTenShillingsOverlay

    ' the pAllowSkewChangeFlg is only set when this control receives focus and the flag is unset when the user selects any another control.

    If pAllowSkewChangeFlg = True Then
        tenShillingsOverlay.SkewDegrees = sliSkewDegrees.Value
    End If

    Call saveMainRCFormSize

    On Error GoTo 0
    Exit Sub

sliSkewDegrees_Scroll_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliSkewDegrees_Scroll of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliWidgetSize_Scroll
' Author    : beededea
' Date      : 10/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub sliWidgetSize_Scroll()
    On Error GoTo sliWidgetSize_Scroll_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
    
    ' pAllowSizeChangeFlg prevents the slider altering the skew on the tenshillings widget unless the slider is the active user control.
    ' It does this in order to avoid reciprocal back and forth between the main widget form and the prefs form
    
    ' ie. the slider slides but does not itself change the widget size, that is happening elsewhere
    ' the widget size is modified in W_MouseWheel in cwTenShillingsOverlay
    
    ' the pAllowSizeChangeFlg is only set when this control receives focus and the flag is unset when the user selects any another control.
    If pAllowSizeChangeFlg = True Then
        Me.WidgetSize = sliWidgetSize.Value / 100
    End If
    
    Call saveMainRCFormSize

    On Error GoTo 0
    Exit Sub

sliWidgetSize_Scroll_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliWidgetSize_Scroll of Form widgetPrefs"
End Sub

Private Sub txtPrefsFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtPrefsFont.hwnd, "This is a read-only text box. It displays the current font as set when you click the font selector button. This is in operation for informational purposes only. When resizing the form (drag bottom right) the font size will change in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change.  My preferred font for this utility is Centurion Light SF at 8pt size.", _
                  TTIconInfo, "Help on the Currently Selected Font", , , , True
End Sub

Private Sub txtPrefsFontCurrentSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtPrefsFontCurrentSize.hwnd, "This is a read-only text box. It displays the current font size as set when dynamic form resizing is enabled. Drag the right hand corner of the window downward and the form will auto-resize. This text box will display the resized font currently in operation for informational purposes only.", _
                  TTIconInfo, "Help on Setting the Font size Dynamically", , , , True
End Sub

Private Sub txtPrefsFontSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtPrefsFontSize.hwnd, "This is a read-only text box. It displays the current base font size as set when dynamic form resizing is enabled. The adjacent text box will display the automatically resized font currently in operation, for informational purposes only.", _
                  TTIconInfo, "Help on the Base Font Size", , , , True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseMove
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo lblDragCorner_MouseMove_Error

    lblDragCorner.MousePointer = 8
    
    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseMove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseMove of Form widgetPrefs"
   
End Sub


Private Sub btnAboutDebugInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnAboutDebugInfo.hwnd, "Here you can switch on Debug mode, not yet functional for this widget.", _
                  TTIconInfo, "Help on the Debug Info. Buttton", , , , True
End Sub



Private Sub btnClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnClose.hwnd, "Close the Preference Utility", _
                  TTIconInfo, "Help on the Close Buttton", , , , True
End Sub

'Private Sub btnDefaultVB6Editor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If gsPrefsTooltips = "0" Then CreateToolTip btnDefaultEditor.hWnd, "Clicking on this button will cause a file explorer window to appear allowing you to select a Visual Basic Project (VBP) file for opening via the right click menu edit option. Once selected the adjacent text field will be automatically filled with the chosen path and file.", _
'                  TTIconInfo, "Help on the VBP File Explorer Button", , , , True
'End Sub

Private Sub btnDisplayScreenFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnDisplayScreenFont.hwnd, "This is the font selector button, if you click it the font selection window will pop up for you to select your chosen font. When resizing the main widget the display screen font size will change in relation to widget size. The base font determines the initial size, the resulting resized font will dynamically change. ", _
                  TTIconInfo, "Help on the Font Selector Button", , , , True
End Sub

Private Sub btnDonate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnDonate.hwnd, "Here you can visit my KofI page and donate a Coffee if you like my creations.", _
                  TTIconInfo, "Help on the Donate Buttton", , , , True
End Sub

Private Sub btnFacebook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnFacebook.hwnd, "Here you can visit the Facebook page for the steampunk Widget community.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub btnGithubHome_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnGithubHome.hwnd, "Here you can visit the widget's home page on github, when you click the button it will open a browser window and take you to the github home page.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub

Private Sub btnHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnHelp.hwnd, "Opens the help document, this will open as a compiled HTML file.", _
                  TTIconInfo, "Help on the Help Buttton", , , , True
End Sub



Private Sub btnOpenFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnOpenFile.hwnd, "Clicking on this button will cause a file explorer window to appear allowing you to select any file you would like to execute on a shift+DBlClick. Once selected the adjacent text field will be automatically filled with the chosen path and file.", _
                  TTIconInfo, "Help on the shift+DBlClick File Explorer Button", , , , True
End Sub

Private Sub btnPrefsFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnPrefsFont.hwnd, "This is the font selector button, if you click it the font selection window will pop up for you to select your chosen font. Centurion Light SF is a good one and my personal favourite. When resizing the form (drag bottom right) the font size will change in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change. ", _
                  TTIconInfo, "Help on Setting the Font Selector Button", , , , True
End Sub

Private Sub btnSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnSave.hwnd, "Save the changes you have made to the preferences", _
                  TTIconInfo, "Help on the Save Buttton", , , , True
End Sub

Private Sub btnUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip btnUpdate.hwnd, "Here you can able to download a new version of the program from github, when you click the button it will open a browser window and take you to the github page.", _
                  TTIconInfo, "Help on the Update Buttton", , , , True
End Sub


Private Sub chkDpiAwareness_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkDpiAwareness.hwnd, "Tick here to make the program DPI aware. When you have Windows scaling switched on using a high res. large screen, application DPI scaling reduces blurriness. It also enables dragging the right-hand corner to resize. NOT required on small to medium screens that are less than 1366 bytes wide. Try it and see which suits your system. HARD restart required.", _
                  TTIconInfo, "Help on DPI Awareness Mode", , , , True
End Sub



Private Sub chkEnableSounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkEnableSounds.hwnd, "Check this box to enable or disable all of the sounds used during any animation on the main steampunk GUI, as well as all other chimes, tick sounds &c.", _
                  TTIconInfo, "Help on Enabling/Disabling Sounds", , , , True
End Sub



Private Sub chkGenStartup_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkGenStartup.hwnd, "Check this box to enable the automatic start of the program when Windows is started.", _
                  TTIconInfo, "Help on the Widget Automatic Start Toggle", , , , True
End Sub

Private Sub chkIgnoreMouse_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkIgnoreMouse.hwnd, "Checking this box causes the program to ignore all mouse events. A strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Ignore Mouse button", , , , True
End Sub


Private Sub chkFormVisible_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkFormVisible.hwnd, "Checking this box makes the underlying form visible. This only helps when developing/debugging. Requires a restart.", _
                  TTIconInfo, "Help on the Form Visible button", , , , True
End Sub


Private Sub chkPreventDragging_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkPreventDragging.hwnd, "Checking this box causes the program to lock in place and ignore all attempts to move it with the mouse. " & vbCrLf & vbCrLf & _
        "The widget can be locked into a certain position in either landscape/portrait mode, ensuring that the widget always appears exactly where you want it to.  " & vbCrLf & vbCrLf & _
        "Using the fields adjacent, you can assign a default x/y position for both Landscape or Portrait mode.  " & vbCrLf & vbCrLf & _
        "When the widget is locked in place (using the Widget Position Locked option in the Window Tab), this value is set automatically.  " & vbCrLf & vbCrLf & _
        "A strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Lock in Place option", , , , True
End Sub

Private Sub chkShowHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkShowHelp.hwnd, "Checking this box causes the rather attractive help canvas to appear every time the widget is started.", _
                  TTIconInfo, "Help on the Ignore Mouse option", , , , True
End Sub

Private Sub chkShowTaskbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkShowTaskbar.hwnd, "Check the box to show the widget in the Windows taskbar. A typical user may have multiple desktop widgets and it makes no sense to fill the taskbar with taskbar entries, this option allows you to enable a single one or two at your whim.", _
                  TTIconInfo, "Help on the Showing Entries in the Taskbar", , , , True
End Sub




Private Sub chkWidgetFunctions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkWidgetFunctions.hwnd, "When checked this box enables this widget's functionality. Any adjustment takes place instantly.", _
                  TTIconInfo, "Help on the Widget Function Toggle", , , , True
End Sub

Private Sub chkWidgetHidden_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip chkWidgetHidden.hwnd, "Checking this box causes the program to hide for a certain number of minutes. More useful from the widget's right click menu where you can hide the widget at will. Seemingly, a strange option, a left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Hidden option", , , , True
End Sub
Private Sub fraAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = True
    If gsPrefsTooltips = "0" Then CreateToolTip fraAbout.hwnd, "The About tab tells you all about this program and its creation using " & gsCodingEnvironment & ".", _
                  TTIconInfo, "Help on the About Tab", , , , True
End Sub
Private Sub fraConfigInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraConfigInner.hwnd, "The configuration panel is the location for optional configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True
End Sub
Private Sub fraConfig_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraConfig.hwnd, "The configuration panel is the location for important configuration items. These items change how the widget operates, configure them to suit your needs and your mode of operation.", _
                  TTIconInfo, "Help on Configuration", , , , True
End Sub
Private Sub fraDevelopment_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraDevelopment.hwnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True
End Sub
'Private Sub fraDefaultVB6Editor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    lblGitHub.ForeColor = &H80000012
'End Sub
Private Sub fraDevelopmentInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraDevelopmentInner.hwnd, "This tab contains elements that will assist in debugging and developing this program further. ", _
                  TTIconInfo, "Help on the Development Tab", , , , True

End Sub
Private Sub fraFonts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraFonts.hwnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gs program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True
End Sub

Private Sub fraFontsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraFontsInner.hwnd, "This tab allows you to set a specific font for the preferences only as there are no textual elements in the main program. We suggest Centurion Light SF at 8pt, which you will find bundled in the gs program folder. Choose a small 8pt font for each.", _
                  TTIconInfo, "Help on Setting the Fonts", , , , True
End Sub


Private Sub fraGeneral_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraGeneral.hwnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraGeneralInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraGeneralInner.hwnd, "The General Panel contains the most important user-configurable items required for the program to operate correctly.", _
                  TTIconInfo, "Help on Essential Configuration", , , , True
End Sub

Private Sub fraPosition_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gsPrefsTooltips = "0" Then CreateToolTip fraPosition.hwnd, "This tab allows you to determine the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet.", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub
Private Sub fraPositionInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraPositionInner.hwnd, "This tab allows you to determine the X and Y positioning of your widget in landscape and portrait screen modes. Best left well alone unless you use Windows on a tablet.", _
                  TTIconInfo, "Help on Tablet Positioning", , , , True
End Sub

Private Sub fraScrollbarCover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = False

End Sub
Private Sub fraSounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If gsPrefsTooltips = "0" Then CreateToolTip fraSounds.hwnd, "The sound panel allows you to configure the sounds that occur within gs. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub
Private Sub fraSoundsInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gsPrefsTooltips = "0" Then CreateToolTip fraSoundsInner.hwnd, "The sound panel allows you to configure the sounds that occur within gs. Some of the animations have associated sounds, you can control these here..", _
                  TTIconInfo, "Help on Configuring Sounds", , , , True
End Sub
Private Sub fraWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gsPrefsTooltips = "0" Then CreateToolTip fraWindow.hwnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub
Private Sub fraWindowInner_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     If gsPrefsTooltips = "0" Then CreateToolTip fraWindowInner.hwnd, "The Opacity and Window Level of the program are rather strange characteristics to change in a Windows program, however this widget is a copy of a Yahoo Widget of the same name. All widgets have similar window tab options including the capability to change the opacity and window level. Whether these options are useful to you or anyone is a moot point but as this tool aims to replicate the YWE version functionality it has been reproduced here. It is here as more of an experiment as to how to implement a feature, one carried over from the Yahoo Widget (javascript) version of this program.", _
                  TTIconInfo, "Help on YWE Quirk Mode Options", , , , True
End Sub


Private Sub fraGeneralButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraGeneralButton.hwnd, "Clicking on the General icon reveals the General Tab where the essential items can be configured, alarms, startup &c.", _
                  TTIconInfo, "Help on the General Tab Icon", , , , True
End Sub

Private Sub fraConfigButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraConfigButton.hwnd, "Clicking on the Config icon reveals the Configuration Tab where the optional items can be configured, DPI, tooltips &c.", _
                  TTIconInfo, "Help on the Configuration Tab Icon", , , , True
End Sub

Private Sub fraFontsButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraFontsButton.hwnd, "Clicking on the Fonts icon reveals the Fonts Tab where the font related items can be configured, size, type, popups &c.", _
                  TTIconInfo, "Help on the Font Tab Icon", , , , True
End Sub
Private Sub fraSoundsButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraSoundsButton.hwnd, "Clicking on the Sounds icon reveals the Sounds Tab where sound related items can be configured, volume, type &c.", _
                  TTIconInfo, "Help on the Sounds Tab Icon", , , , True
End Sub
Private Sub fraPositionButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraPositionButton.hwnd, "Clicking on the Position icon reveals the Position Tab where items related to Positioning can be configured, aspect ratios, landscape, &c.", _
                  TTIconInfo, "Help on the Position Tab Icon", , , , True
End Sub
Private Sub fraDevelopmentButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraDevelopmentButton.hwnd, "Clicking on the Development icon reveals the Development Tab where items relating to Development can be configured, debug, VBP location, &c.", _
                  TTIconInfo, "Help on the Development Tab Icon", , , , True
End Sub
Private Sub fraWindowButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraWindowButton.hwnd, "Clicking on the Window icon reveals the Window Tab where items relating to window sizing and layering can be configured &c.", _
                  TTIconInfo, "Help on the Window Tab Icon", , , , True
End Sub
Private Sub fraAboutButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip fraAboutButton.hwnd, "Clicking on the About icon reveals the About Tab where information about this desktop widget can be revealed.", _
                  TTIconInfo, "Help on the About Tab Icon", , , , True
End Sub
Private Sub lblGitHub_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblGitHub.ForeColor = &H8000000D
End Sub

'---------------------------------------------------------------------------------------
' Procedure : optPrefsTooltips_MouseMove
' Author    : beededea
' Date      : 10/01/2025
' Purpose   : series of radio buttons to set the tooltip type for the prefs utility
'---------------------------------------------------------------------------------------
'
Private Sub optPrefsTooltips_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim thisToolTip As String: thisToolTip = vbNullString

    On Error GoTo optPrefsTooltips_MouseMove_Error

    If gsPrefsTooltips = "0" Then
        If Index = 0 Then
            thisToolTip = "This setting enables the balloon tooltips for elements within the Steampunk GUI. These tooltips are multi-line and in general more attractive than standard windows style tooltips, note that their font size will match the Windows system font size."
            CreateToolTip optPrefsTooltips(Index).hwnd, thisToolTip, _
                  TTIconInfo, "Help on Balloon Tooltips on the Preference Utility", , , , True
        ElseIf Index = 1 Then
            thisToolTip = "This setting enables the standard Windows-style square tooltips for elements within the Steampunk GUI. These tooltips are single-line and the font size is limited to the Windows font size."
            CreateToolTip optPrefsTooltips(Index).hwnd, thisToolTip, _
                  TTIconInfo, "Help on " & gsCodingEnvironment & " Native Tooltips on the Preference Utility", , , , True
        ElseIf Index = 2 Then
            thisToolTip = "This setting disables the balloon tooltips for elements within the Steampunk GUI."
            CreateToolTip optPrefsTooltips(Index).hwnd, thisToolTip, _
                  TTIconInfo, "Help on Disabling Tooltips on the Preference Utility", , , , True
        End If
    End If

   On Error GoTo 0
   Exit Sub

optPrefsTooltips_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure optPrefsTooltips_MouseMove of Form widgetPrefs"
End Sub

Private Sub sliWidgetSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip sliWidgetSize.hwnd, "Adjust to a percentage of the original size. Any adjustment in size made here takes place instantly (you can also use Ctrl+Mousewheel when hovering over the widget itself).", _
                  TTIconInfo, "Help on the Size sliSkewDegrees", , , , True
End Sub

Private Sub sliOpacity_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip sliOpacity.hwnd, "Sliding this causes the program's opacity to change from solidly opaque to fully transparent or some way in-between. Seemingly, a strange option for a windows program, a useful left-over from the Yahoo Widgets days that offered this additional option. Replicated here as a homage to the old widget platform.", _
                  TTIconInfo, "Help on the Opacity sliSkewDegrees", , , , True

End Sub
Private Sub txtAboutText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraScrollbarCover.Visible = False
End Sub



Private Sub txtDblClickCommand_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtDblClickCommand.hwnd, "Field to hold the any double click command that you have assigned to this widget. For example: taskmgr or %systemroot%\syswow64\ncpa.cpl", _
                  TTIconInfo, "Help on the Double Click Command", , , , True
End Sub

Private Sub txtDefaultEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtDefaultEditor.hwnd, "Field to hold the path to a Visual Basic Project (VBP) file you would like to execute on a right click menu, edit option, if you select the adjacent button a file explorer will appear allowing you to select the VBP file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the Default Editor Field", , , , True
End Sub

Private Sub txtDisplayScreenFont_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtDisplayScreenFont.hwnd, "This is a read-only text box. It displays the current font - as set when you click the font selector button. This field is in operation for informational purposes only. When resizing the main widget (CTRL+ mouse scroll wheel) the font size will change in relation to widget size. The base font determines the initial size, the resulting resized font will dynamically change. My preferred font for the display screen is Courier New at 6pt size.", _
                  TTIconInfo, "Help on the Display Screen Font", , , , True
End Sub

Private Sub txtDisplayScreenFontSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtDisplayScreenFontSize.hwnd, "This is a read-only text box. It displays the current base font size as set when dynamic form resizing is enabled. The adjacent text box will display the automatically resized font currently in operation, for informational purposes only.", _
                  TTIconInfo, "Help on the Base Font Size for Display Screen", , , , True
End Sub

Private Sub txtLandscapeHoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtLandscapeHoffset.hwnd, "Field to hold the horizontal offset for the widget position in landscape mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Landscape X Horizontal Field", , , , True
End Sub

Private Sub txtLandscapeVoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtLandscapeVoffset.hwnd, "Field to hold the vertical offset for the widget position in landscape mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Landscape Y Vertical Field", , , , True
End Sub

Private Sub txtOpenFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtOpenFile.hwnd, "Field to hold the path to a file you would like to execute on a shift+DBlClick, if you select the adjacent button a file explorer will appear allowing you to select any file, this field is automatically filled with the chosen file.", _
                  TTIconInfo, "Help on the shift+DBlClick Field", , , , True
End Sub

Private Sub txtPortraitHoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtPortraitHoffset.hwnd, "Field to hold the horizontal offset for the widget position in Portrait mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Portrait X Horizontal Field", , , , True
End Sub

Private Sub txtPortraitYoffset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If gsPrefsTooltips = "0" Then CreateToolTip txtPortraitYoffset.hwnd, "Field to hold the vertical offset for the widget position in Portrait mode. When you lock the widget using the lock button above, this field is automatically filled.", _
                  TTIconInfo, "Help on the Portrait Y Vertical Field", , , , True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : General _MouseDown events to generate menu pop-ups across the form
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : due to a bug/difference with TwinBasic versus VB6
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : Form_MouseDown
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : standard form down event to generate the menu across the board
'---------------------------------------------------------------------------------------
'
Private Sub Form_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo Form_MouseDown_Error

    If Button = 2 Then

        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
        
    End If

   On Error GoTo 0
   Exit Sub

Form_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_MouseDown of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lblDragCorner_MouseDown
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : the label corner mouse down
'---------------------------------------------------------------------------------------
'
Private Sub lblDragCorner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo lblDragCorner_MouseDown_Error
    
    If Button = vbLeftButton Then
        pPrefsFormResizedByDrag = True
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
    End If
    
    On Error GoTo 0
    Exit Sub

lblDragCorner_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblDragCorner_MouseDown of Form widgetPrefs"

End Sub

Private Sub fraFonts_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
'
Private Sub fraFontsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraGeneral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraConfigInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraDevelopment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraDevelopmentInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraGeneralInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraPositionInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraSounds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraSoundsInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub

Private Sub fraWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub fraWindowInner_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If
End Sub
Private Sub imgGeneral_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgGeneral.Visible = False
    imgGeneralClicked.Visible = True
End Sub
Private Sub imgAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgAbout.Visible = False
    imgAboutClicked.Visible = True
End Sub
Private Sub imgDevelopment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgDevelopment.Visible = False
    imgDevelopmentClicked.Visible = True
End Sub
Private Sub imgFonts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgFonts.Visible = False
    imgFontsClicked.Visible = True
End Sub
Private Sub imgConfig_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgConfig.Visible = False
    imgConfigClicked.Visible = True
End Sub
Private Sub imgPosition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgPosition.Visible = False
    imgPositionClicked.Visible = True
End Sub
Private Sub imgSounds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgSounds.Visible = False
    imgSoundsClicked.Visible = True
End Sub
Private Sub imgWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgWindow.Visible = False
    imgWindowClicked.Visible = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtAboutText_MouseDown
' Author    : beededea
' Date      : 30/09/2023
' Purpose   : make a pop up menu appear on the text box by being tricky and clever
'---------------------------------------------------------------------------------------
'
Private Sub txtAboutText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo txtAboutText_MouseDown_Error

    If Button = vbRightButton Then
        txtAboutText.Enabled = False
        txtAboutText.Enabled = True
        Me.PopupMenu prefsMnuPopmenu, vbPopupMenuRightButton
    End If

    On Error GoTo 0
    Exit Sub

txtAboutText_MouseDown_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtAboutText_MouseDown of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : lblGitHub_dblClick
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : label to allow a link to github to be clicked
'---------------------------------------------------------------------------------------
'
Private Sub lblGitHub_dblClick()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo lblGitHub_dblClick_Error
    
    If gsWidgetFunctions = "0" Or gsIgnoreMouse = "1" Then Exit Sub

    answer = vbYes
    answerMsg = "This option opens a browser window and take you straight to Github. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Proceed to Github? ", True, "lblGitHubDblClick")
    If answer = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", "https://github.com/yereverluvinunclebert/TenShillings-" & gsRichClientEnvironment & "-Widget-" & gsCodingEnvironment, vbNullString, App.Path, 1)
    End If

   On Error GoTo 0
   Exit Sub

lblGitHub_dblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lblGitHub_dblClick of Form widgetPrefs"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : tmrPrefsScreenResolution_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : for handling rotation of the screen in tablet mode or a resolution change
'             possibly due to an old game in full screen mode.
            ' when the timer frequency is reduced caused some weird async. effects, 500 ms seems fine, disable the timer first
'---------------------------------------------------------------------------------------
'
Private Sub tmrPrefsScreenResolution_Timer()

'    Static oldWidgetPrefsLeft As Long
'    Static oldWidgetPrefsTop As Long
'    Static beenMovingFlg As Boolean
'
'    Static oldPrefsFormMonitorID As Long
''    Static oldPrefsFormMonitorPrimary As Long
'    Static oldgPrefsMonitorStructWidthTwips As Long
'    Static oldgPrefsMonitorStructHeightTwips As Long
'    Static oldPrefsWidgetLeftPixels As Long
'
'    Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
'    'Dim prefsFormMonitorPrimary As Long: prefsFormMonitorPrimary = 0
'    Dim monitorStructWidthTwips As Long: monitorStructWidthTwips = 0
'    Dim monitorStructHeightTwips As Long: monitorStructHeightTwips = 0
'    Dim resizeProportion As Double: resizeProportion = 0
'
'    Dim answer As VbMsgBoxResult: answer = vbNo
'    Dim answerMsg As String: answerMsg = vbNullString
'
    On Error GoTo tmrPrefsScreenResolution_Timer_Error
'
'    ' calls a routine that tests for a change in the monitor upon which the form sits, if so, resizes
'    If widgetPrefs.IsVisible = False Then Exit Sub
'
'    ' prefs hasn't moved at all
'    If widgetPrefs.Left = oldWidgetPrefsLeft Then Exit Sub  ' this can only work if the reposition is being performed by the timer
'    ' we are also hopefully calling this routine on a mouseUP event after a form move, where the above line will not apply.
'
'    ' if just one monitor or the global switch is off then exit
'    If monitorCount > 1 And (LTrim$(gsMultiMonitorResize) = "1" Or LTrim$(gsMultiMonitorResize) = "2") Then
'
'        ' turn off the timer that saves the prefs height and position
'        tmrPrefsMonitorSaveHeight.Enabled = False
'        tmrWritePositionAndSize.Enabled = False
'        tmrPrefsScreenResolution.Enabled = False ' turn off this very timer here
'
'        ' populate the OLD vars if empty, to allow valid comparison next run
'        If oldWidgetPrefsLeft <= 0 Then oldWidgetPrefsLeft = widgetPrefs.Left
'        If oldWidgetPrefsTop <= 0 Then oldWidgetPrefsTop = widgetPrefs.Top
'
'        ' test whether the form has been moved (VB6 has no FORM_MOVING nor FormResizedOrMoved EVENTS)
'        If widgetPrefs.Left <> oldWidgetPrefsLeft Or widgetPrefs.Top <> oldWidgetPrefsTop Then
'
'               ' note the monitor ID at PrefsForm form_load and store as the prefsFormMonitorID
'                gPrefsMonitorStruct = formScreenProperties(widgetPrefs, prefsFormMonitorID)
'
'                'prefsFormMonitorPrimary = gPrefsMonitorStruct.IsPrimary ' -1 true
'
'                ' sample the physical monitor resolution
'                monitorStructWidthTwips = gPrefsMonitorStruct.Width
'                monitorStructHeightTwips = gPrefsMonitorStruct.Height
'
'                'if the old monitor ID has not been stored already (form load) then do so now
'                If oldPrefsFormMonitorID = 0 Then oldPrefsFormMonitorID = prefsFormMonitorID
'
'                ' same with other 'old' vars
'                If oldgPrefsMonitorStructWidthTwips = 0 Then oldgPrefsMonitorStructWidthTwips = monitorStructWidthTwips
'                If oldgPrefsMonitorStructHeightTwips = 0 Then oldgPrefsMonitorStructHeightTwips = monitorStructHeightTwips
'                If oldPrefsWidgetLeftPixels = 0 Then oldPrefsWidgetLeftPixels = widgetPrefs.Left
'
'                ' if the monitor ID has changed
'                If oldPrefsFormMonitorID <> prefsFormMonitorID Then
'                'If oldPrefsFormMonitorPrimary <> prefsFormMonitorPrimary Then
'
''                    ' screenWrite ("Prefs Stored monitor primary status = " & CBool(oldPrefsFormMonitorPrimary))
''                    ' screenWrite ("Prefs Current monitor primary status = " & CBool(prefsFormMonitorPrimary))
'
'                    If LTrim$(gsMultiMonitorResize) = "1" Then
'                        'if the resolution is different then calculate new size proportion
'                        If monitorStructWidthTwips <> oldgPrefsMonitorStructWidthTwips Or monitorStructHeightTwips <> oldgPrefsMonitorStructHeightTwips Then
'                            'now calculate the size of the widget according to the screen HeightTwips.
'                            resizeProportion = gPrefsMonitorStruct.Height / oldgPrefsMonitorStructHeightTwips
'                            newPrefsHeight = widgetPrefs.Height * resizeProportion
'                            widgetPrefs.Height = newPrefsHeight
'                        End If
'                    ElseIf LTrim$(gsMultiMonitorResize) = "2" Then
'                        ' set the size according to saved values
'                        If gPrefsMonitorStruct.IsPrimary = True Then
'                            widgetPrefs.Height = Val(gsPrefsPrimaryHeightTwips)
'                        Else
'                            'gsPrefsSecondaryHeightTwips = "15000"
'                            widgetPrefs.Height = Val(gsPrefsSecondaryHeightTwips)
'                        End If
'                    End If
'
'                End If
'
'                ' set the current values as 'old' for comparison on next run
'                'oldPrefsFormMonitorPrimary = prefsFormMonitorPrimary
'                oldPrefsFormMonitorID = prefsFormMonitorID
'                oldgPrefsMonitorStructWidthTwips = monitorStructWidthTwips
'                oldgPrefsMonitorStructHeightTwips = monitorStructHeightTwips
'                oldPrefsWidgetLeftPixels = widgetPrefs.Left
'            End If
'
'    End If
'
'    oldWidgetPrefsLeft = widgetPrefs.Left
'    oldWidgetPrefsTop = widgetPrefs.Top
'
'    tmrPrefsScreenResolution.Enabled = True
'    tmrPrefsMonitorSaveHeight.Enabled = True
'    tmrWritePositionAndSize.Enabled = True
    
    On Error GoTo 0
    Exit Sub

tmrPrefsScreenResolution_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPrefsScreenResolution_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub





'---------------------------------------------------------------------------------------
' Procedure : General _MouseUp events to generate menu pop-ups across the form
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : due to a bug/difference with TwinBasic versus VB6
'---------------------------------------------------------------------------------------
'#If TWINBASIC Then
'    Private Sub imgAboutClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
'    End Sub
'#Else
    Private Sub imgAbout_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("about", imgAbout, imgAboutClicked, fraAbout, fraAboutButton)
    End Sub
'#End If

'#If TWINBASIC Then
'    Private Sub imgDevelopmentClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
'    End Sub
'#Else
    Private Sub imgDevelopment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("development", imgDevelopment, imgDevelopmentClicked, fraDevelopment, fraDevelopmentButton)
    End Sub
'#End If

'#If TWINBASIC Then
'    Private Sub imgFontsClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
'    End Sub
'#Else
    Private Sub imgFonts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("fonts", imgFonts, imgFontsClicked, fraFonts, fraFontsButton)
    End Sub
'#End If

'#If TWINBASIC Then
'    Private Sub imgConfigClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
'    End Sub
'#Else
    Private Sub imgConfig_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("config", imgConfig, imgConfigClicked, fraConfig, fraConfigButton) ' was imgConfigMouseUpEvent
    End Sub
'#End If

'#If TWINBASIC Then
'    Private Sub imgPositionClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
'    End Sub
'#Else
    Private Sub imgPosition_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("position", imgPosition, imgPositionClicked, fraPosition, fraPositionButton)
    End Sub
'#End If

'#If TWINBASIC Then
'    Private Sub imgSoundsClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
'    End Sub
'#Else
    Private Sub imgSounds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("sounds", imgSounds, imgSoundsClicked, fraSounds, fraSoundsButton)
    End Sub
'#End If

'#If TWINBASIC Then
'    Private Sub imgWindowClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
'    End Sub
'#Else
    Private Sub imgWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("window", imgWindow, imgWindowClicked, fraWindow, fraWindowButton)
    End Sub
'#End If

'#If TWINBASIC Then
'    Private Sub imgGeneralClicked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'        Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton) ' was imgGeneralMouseUpEvent
'    End Sub
'#Else
    Private Sub imgGeneral_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call picButtonMouseUpEvent("general", imgGeneral, imgGeneralClicked, fraGeneral, fraGeneralButton) ' was imgGeneralMouseUpEvent
    End Sub
'#End If




Private Sub sliSkewDegrees_GotFocus()
    pAllowSkewChangeFlg = True
End Sub

Private Sub sliSkewDegrees_LostFocus()
    pAllowSkewChangeFlg = False
End Sub

Private Sub sliWidgetSize_GotFocus()
    pAllowSizeChangeFlg = True
End Sub

Private Sub sliWidgetSize_LostFocus()
    pAllowSizeChangeFlg = False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Various Change events below
' Author    : beededea
' Date      : 15/08/2023
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Procedure : sliOpacity_Change
' Author    : beededea
' Date      : 15/08/2023
' Purpose   : save the sliSkewDegrees opacity values as they change
'---------------------------------------------------------------------------------------
'
Private Sub sliOpacity_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    
    On Error GoTo sliOpacity_Change_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

    If pPrefsStartupFlg = False Then
        gsOpacity = CStr(sliOpacity.Value)
    
        sPutINISetting "Software\TenShillings", "opacity", gsOpacity, gsSettingsFile
        
        'Call setOpacity(sliOpacity.Value) ' this works but reveals the background form itself
        
        answer = vbYes
        answerMsg = "You must perform a hard reload on this widget in order to change the widget's opacity, do you want me to do it for you now?"
        answer = msgBoxA(answerMsg, vbYesNo, "Hard Reload Request", True, "sliOpacityClick")
        If answer = vbNo Then
            Exit Sub
        Else
            Call hardRestart
        End If
    End If

   On Error GoTo 0
   Exit Sub

sliOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliOpacity_Change of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : sliOpacity_Change
' Author    : beededea
' Date      : 18/02/2025
' Purpose   : sliSkewDegrees to change opacity of the whole widget.
'---------------------------------------------------------------------------------------
'
Private Sub sliOpacity_Change()
   On Error GoTo sliOpacity_Change_Error

    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

   On Error GoTo 0
   Exit Sub

sliOpacity_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliOpacity_Change of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : sliWidgetSize_Change
' Author    : beededea
' Date      : 30/09/2023
' Purpose   : sliSkewDegrees to change the size of the whole widget.
'---------------------------------------------------------------------------------------
'
'Public Sub sliWidgetSize_Change()
'    On Error GoTo sliWidgetSize_Change_Error
'
'
'
'    On Error GoTo 0
'    Exit Sub
'
'sliWidgetSize_Change_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliWidgetSize_Change of Form widgetPrefs"
'
'End Sub

'---------------------------------------------------------------------------------------
' Property  : WidgetSize
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the WidgetSize of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Get WidgetSize() As Single
   On Error GoTo widgetSizeGet_Error

   WidgetSize = mWidgetSize

   On Error GoTo 0
   Exit Property

widgetSizeGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WidgetSize of Class Module cwTenShillingsOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property  : WidgetSize
' Author    : beededea
' Date      : 10/05/2023
' Purpose   : property to determine (by value) the WidgetSize value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Let WidgetSize(ByVal newValue As Single)
   On Error GoTo widgetSizeLet_Error

    If mWidgetSize <> newValue Then mWidgetSize = newValue Else Exit Property
        
    tenShillingsOverlay.Zoom = (mWidgetSize)

   On Error GoTo 0
   Exit Property

widgetSizeLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property WidgetSize of Class Module cwTenShillingsOverlay"
End Property


Private Sub txtDblClickCommand_Change()
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtDefaultEditor_Change()
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeHoffset_Change()
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtLandscapeVoffset_Change()
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
End Sub
Private Sub txtOpenFile_Change()
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
End Sub

Private Sub txtPortraitHoffset_Change()
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
End Sub

Private Sub txtPortraitYoffset_Change()
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button

End Sub

Private Sub txtPrefsFont_Change()
    If gbStartupFlg = False Then btnSave.Enabled = True ' enable the save button
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuAbout_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : right click about option from the pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuAbout_Click()
    
    On Error GoTo mnuAbout_Click_Error

    Call aboutClickEvent

    On Error GoTo 0
    Exit Sub

mnuAbout_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAbout_Click of form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsTooltips
' Author    : beededea
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setPrefsTooltips()

   On Error GoTo setPrefsTooltips_Error
    
    ' here we set the variables used for the comboboxes, each combobox has to be sub classed and these variables are used during that process
    
    If optPrefsTooltips(0).Value = True Then
        ' module level balloon tooltip variables for subclassed comboBoxes ONLY.
        pCmbMultiMonitorResizeBalloonTooltip = "This option will only appear on multi-monitor systems. This dropdown has three choices that affect the automatic sizing of both the main widget and the preference utility. " & vbCrLf & vbCrLf & _
            "For monitors of different sizes, this allows you to resize the widget to suit the monitor it is currently sitting on. The automatic option resizes according to the relative proportions of the two screens.  " & vbCrLf & vbCrLf & _
            "The manual option resizes according to sizes that you set manually. Just resize the widget on the monitor of your choice and the program will store it. This option only works for no more than TWO monitors."
   
        pCmbScrollWheelDirectionBalloonTooltip = "This option will allow you to change the direction of the mouse scroll wheel when resizing the widget. IF you want to resize the widget on your desktop, hold the CTRL key along with moving the scroll wheel UP/DOWN. Some prefer scrolling UP rather than DOWN. You configure that here."
        pCmbWindowLevelBalloonTooltip = "You can determine the window level here. You can keep it above all other windows or you can set it to bottom to keep the widget below all other windows."
        pCmbHidingTimeBalloonTooltip = "The hiding time that you can set here determines how long the widget will disappear when you click the menu option to hide the widget."
        
        pCmbWidgetLandscapeBalloonTooltip = "The widget can be locked into landscape mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for Landscape mode. "
        pCmbWidgetPortraitBalloonTooltip = "The widget can be locked into portrait mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for portrait mode. "
        pCmbWidgetPositionBalloonTooltip = "Tablets only. The widget can be positioned proportionally when switching between portrait/landscape. If you want to enable this, disable the options below."
        pCmbAspectHiddenBalloonTooltip = "Here you can choose whether the widget is hidden by default in either landscape or portrait mode or not at all. This allows you to have certain widgets that do not obscure the screen in one mode or another. If you accidentally set it so you can't find it on screen then change the setting here to none."
        
        pCmbDebugBalloonTooltip = "Here you can set debug mode. This will enable the editor field and allow you to assign a VBP/TwinProj file for the " & gsCodingEnvironment & " IDE editor"
        
    Else
        ' module level balloon tooltip variables for subclassed comboBoxes ONLY.
        
        pCmbMultiMonitorResizeBalloonTooltip = vbNullString
        pCmbScrollWheelDirectionBalloonTooltip = vbNullString
        pCmbWindowLevelBalloonTooltip = vbNullString
        pCmbHidingTimeBalloonTooltip = vbNullString
        
        pCmbWidgetLandscapeBalloonTooltip = vbNullString
        pCmbWidgetPortraitBalloonTooltip = vbNullString
        pCmbWidgetPositionBalloonTooltip = vbNullString
        pCmbAspectHiddenBalloonTooltip = vbNullString
        pCmbDebugBalloonTooltip = vbNullString
        
        ' for some reason, the balloon tooltip on the checkbox used to dismiss the balloon tooltips does not disappear, this forces it go away.
        CreateToolTip optPrefsTooltips(0).hwnd, "", _
                  TTIconInfo, "Help", , , , True
        CreateToolTip optPrefsTooltips(1).hwnd, "", _
                  TTIconInfo, "Help", , , , True
        CreateToolTip optPrefsTooltips(2).hwnd, "", _
                  TTIconInfo, "Help", , , , True
                  

        
    End If
    
    
    ' next we just do the native VB6 tooltips
     If optPrefsTooltips(1).Value = True Then
        imgConfig.ToolTipText = "Opens the configuration tab"
        imgConfigClicked.ToolTipText = "Opens the configuration tab"
        imgDevelopment.ToolTipText = "Opens the Development tab"
        imgDevelopmentClicked.ToolTipText = "Opens the Development tab"
        imgPosition.ToolTipText = "Opens the Position tab"
        imgPositionClicked.ToolTipText = "Opens the Position tab"
        btnSave.ToolTipText = "Save the changes you have made to the preferences"
        btnHelp.ToolTipText = "Open the help utility"
        imgSounds.ToolTipText = "Opens the Sounds tab"
        imgSoundsClicked.ToolTipText = "Opens the Sounds tab"
        btnClose.ToolTipText = "Close the utility"
        imgWindow.ToolTipText = "Opens the Window tab"
        imgWindowClicked.ToolTipText = "Opens the Window tab"
        lblWindow.ToolTipText = "Opens the Window tab"
        imgFonts.ToolTipText = "Opens the Fonts tab"
        imgFonts.ToolTipText = "Opens the Fonts tab"
        imgFontsClicked.ToolTipText = "Opens the Fonts tab"
        imgGeneral.ToolTipText = "Opens the general tab"
        imgGeneralClicked.ToolTipText = "Opens the general tab"
        lblPosition(6).ToolTipText = "Tablets only. Don't fiddle with this unless you really know what you are doing. Here you can choose whether this the widget widget is hidden by default in either landscape or portrait mode or not at all. This option allows you to have certain widgets that do not obscure the screen in either landscape or portrait. If you accidentally set it so you can't find your widget on screen then change the setting here to NONE."
        chkGenStartup.ToolTipText = "Check this box to enable the automatic start of the program when Windows is started."
        chkWidgetFunctions.ToolTipText = "When checked this box enables this widget's functionality. Any adjustment takes place instantly. "
        
        txtPortraitYoffset.ToolTipText = "Field to hold the vertical offset for the widget position in portrait mode."
        txtPortraitHoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in portrait mode."
        txtLandscapeVoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in landscape mode."
        txtLandscapeHoffset.ToolTipText = "Field to hold the horizontal offset for the widget position in landscape mode."
        cmbWidgetLandscape.ToolTipText = "The widget can be locked into landscape mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for Landscape mode. "
        cmbWidgetPortrait.ToolTipText = "The widget can be locked into portrait mode, it ensures that the widget always appears where you want it to. Using the fields below, you can assign a default x/y position for portrait mode. "
        cmbWidgetPosition.ToolTipText = "Tablets only. The widget can be positioned proportionally when switching between portrait/landscape. If you want to enable this, disable the options below."
        cmbAspectHidden.ToolTipText = " Here you can choose whether the widget is hidden by default in either landscape or portrait mode or not at all. This allows you to have certain widgets that do not obscure the screen in one mode or another. If you accidentally set it so you can't find it on screen then change the setting here to none."
        chkEnableSounds.ToolTipText = "Check this box to enable or disable all of the sounds used during any animation on the main steampunk GUI as well as all other chimes, tick sounds."
'        chkEnableTicks.ToolTipText = "Enables or disables just the sound of the widget ticking."
'        chkEnableChimes.ToolTipText = "Enables or disables just the widget chimes."
'        chkEnableAlarms.ToolTipText = "Enables or disables the widget alarm chimes. Please note disabling this means your alarms will not alert you audibly!"
        
'        chkVolumeBoost.ToolTipText = "Sets the volume of the various sound elements, you can boost from quiet to loud."
        btnDefaultEditor.ToolTipText = "Click to select the .vbp file to edit the program - You need to have access to the source!"
        txtDblClickCommand.ToolTipText = "Enter a Windows command for the widget to operate when double-clicked."
        btnOpenFile.ToolTipText = "Click to select a particular file for the widget to run or open when double-clicked."
        txtOpenFile.ToolTipText = "Enter a particular file for the widget to run or open when double-clicked."
        cmbDebug.ToolTipText = "Choose to set debug mode."
        
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within this preferences window only"
        btnPrefsFont.ToolTipText = "The Font Selector."
        txtPrefsFont.ToolTipText = "Disabled for manual input. Choose a font via the font selector to be used only for this preferences window"
        txtPrefsFontSize.ToolTipText = "Disabled for manual input. Choose a font size via the font selector that fits the text boxes"
        
        lblFontsTab(0).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
        lblFontsTab(6).ToolTipText = "We suggest Centurion Light SF at 8pt - which you will find in the FCW program folder"
        lblFontsTab(7).ToolTipText = "Choose a font size that fits the text boxes"
        
        txtDisplayScreenFontSize.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within the widget display screen only"
        btnDisplayScreenFont.ToolTipText = "The Font Selector."
        txtDisplayScreenFont.ToolTipText = "Disabled for manual input. Choose a font size using the font selector to be used within the widget display screen only"
        
        cmbWindowLevel.ToolTipText = "You can determine the window position here. Set to bottom to keep the widget below other windows."
        cmbHidingTime.ToolTipText = "The hiding time that you can set here determines how long the widget will disappear when you click the menu option to hide the widget."
        cmbMultiMonitorResize.ToolTipText = "When you have a multi-monitor set-up, the widget can auto-resize on a smaller secondary monitor. Here you determine the proportion of the resize."
        
        chkEnableResizing.ToolTipText = "Provides an alternative method of supporting high DPI screens."
        chkPreventDragging.ToolTipText = "Checking this box turns off the ability to drag the program with the mouse. The locking in position effect takes place instantly."
        chkIgnoreMouse.ToolTipText = "Checking this box causes the program to ignore all mouse events."
        chkFormVisible.ToolTipText = "Checking this box causes the underlying form to show itself, useful only when debugging."
        
        sliOpacity.ToolTipText = "Set the transparency of the program. Any change in opacity takes place instantly."
        cmbScrollWheelDirection.ToolTipText = "To change the direction of the mouse scroll wheel when resizing the widget."
        
        optWidgetTooltips(0).ToolTipText = "Check the box to enable larger balloon tooltips for all controls on the main program"
        optWidgetTooltips(1).ToolTipText = "Check the box to enable RichClient square tooltips for all controls on the main program"
        optWidgetTooltips(2).ToolTipText = "Check the box to disable tooltips for all controls on the main program"
        
        chkShowTaskbar.ToolTipText = "Check the box to show the widget in the taskbar"
        chkShowHelp.ToolTipText = "Check the box to show the help page on startup"
        
        sliWidgetSize.ToolTipText = "Adjust to a percentage of the original size. Any adjustment in size takes place instantly (you can also use Ctrl+Mousewheel hovering over the widget itself)."
        btnFacebook.ToolTipText = "This will link you to the our Steampunk/Dieselpunk program users Group."
        imgAbout.ToolTipText = "Opens the About tab"
        btnAboutDebugInfo.ToolTipText = "This gives access to the debugging tool"
        btnDonate.ToolTipText = "Buy me a Kofi! This button opens a browser window and connects to Kofi donation page"
        btnUpdate.ToolTipText = "Here you can visit the update location where you can download new versions of the programs."
        
        btnGithubHome.ToolTipText = "Here you can visit the widget's home page on github, when you click the button it will open a browser window and take you to the github home page."

        txtPrefsFontCurrentSize.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        'lblCurrentFontsTab.ToolTipText = "Disabled for manual input. Shows the current font size when form resizing is enabled."
        
        chkDpiAwareness.ToolTipText = " Check the box to make the program DPI aware. RESTART required."
        'optEnablePrefsTooltips.ToolTipText = "Check the box to enable tooltips for all controls in the preferences utility"
        
        optPrefsTooltips(0).ToolTipText = "Check the box to enable larger balloon tooltips for all controls within this Preference Utility. These tooltips are multi-line and in general more attractive, note that their font size will match the Windows system font size."
        optPrefsTooltips(1).ToolTipText = "Check the box to enable Windows-style square tooltips for all controls within this Preference Utility. Note that their font size will match the Windows system font size."
        optPrefsTooltips(2).ToolTipText = "This setting enables/disables the tooltips for all elements within this Preference Utility."

        btnResetMessages.ToolTipText = "This button restores the pop-up messages to their original visible state."

        sliSkewDegrees.ToolTipText = "Adjust to rotate the whole widget. Any adjustment in skew made here takes place instantly (you can also use the Mousewheel when hovering over the widget itself)."
    Else
    
        imgConfig.ToolTipText = vbNullString
        imgConfigClicked.ToolTipText = vbNullString
        imgDevelopment.ToolTipText = vbNullString
        imgDevelopmentClicked.ToolTipText = vbNullString
        imgPosition.ToolTipText = vbNullString
        imgPositionClicked.ToolTipText = vbNullString
        btnSave.ToolTipText = vbNullString
        btnHelp.ToolTipText = vbNullString
        imgSounds.ToolTipText = vbNullString
        imgSoundsClicked.ToolTipText = vbNullString
        btnClose.ToolTipText = vbNullString
        imgWindow.ToolTipText = vbNullString
        imgWindowClicked.ToolTipText = vbNullString
        imgFonts.ToolTipText = vbNullString
        imgFontsClicked.ToolTipText = vbNullString
        imgGeneral.ToolTipText = vbNullString
        imgGeneralClicked.ToolTipText = vbNullString
        chkGenStartup.ToolTipText = vbNullString
        chkWidgetFunctions.ToolTipText = vbNullString
                
        txtPortraitYoffset.ToolTipText = vbNullString
        txtPortraitHoffset.ToolTipText = vbNullString
        txtLandscapeVoffset.ToolTipText = vbNullString
        txtLandscapeHoffset.ToolTipText = vbNullString
        cmbWidgetLandscape.ToolTipText = vbNullString
        cmbWidgetPortrait.ToolTipText = vbNullString
        cmbWidgetPosition.ToolTipText = vbNullString
        cmbAspectHidden.ToolTipText = vbNullString
        chkEnableSounds.ToolTipText = vbNullString
'        chkEnableTicks.ToolTipText = vbNullString
'        chkEnableChimes.ToolTipText = vbNullString
'        chkEnableAlarms.ToolTipText = vbNullString
        'chkVolumeBoost.ToolTipText = vbNullString
        
        btnDefaultEditor.ToolTipText = vbNullString
        txtDblClickCommand.ToolTipText = vbNullString
        btnOpenFile.ToolTipText = vbNullString
        txtOpenFile.ToolTipText = vbNullString
        cmbDebug.ToolTipText = vbNullString
        txtPrefsFontSize.ToolTipText = vbNullString
        btnPrefsFont.ToolTipText = vbNullString
        txtPrefsFont.ToolTipText = vbNullString
        txtPrefsFontCurrentSize.ToolTipText = vbNullString
        
        
        txtDisplayScreenFontSize.ToolTipText = vbNullString
        btnDisplayScreenFont.ToolTipText = vbNullString
        txtDisplayScreenFont.ToolTipText = vbNullString
        
        cmbWindowLevel.ToolTipText = vbNullString
        cmbHidingTime.ToolTipText = vbNullString
        cmbMultiMonitorResize.ToolTipText = vbNullString
        
        chkEnableResizing.ToolTipText = vbNullString
        chkPreventDragging.ToolTipText = vbNullString
        chkIgnoreMouse.ToolTipText = vbNullString
        chkFormVisible.ToolTipText = vbNullString
        sliOpacity.ToolTipText = vbNullString
        cmbScrollWheelDirection.ToolTipText = vbNullString
        
        optWidgetTooltips(0).ToolTipText = vbNullString
        optWidgetTooltips(1).ToolTipText = vbNullString
        optWidgetTooltips(2).ToolTipText = vbNullString
        
        chkShowTaskbar.ToolTipText = vbNullString
        chkShowHelp.ToolTipText = vbNullString

        sliWidgetSize.ToolTipText = vbNullString
        btnFacebook.ToolTipText = vbNullString
        imgAbout.ToolTipText = vbNullString
        btnAboutDebugInfo.ToolTipText = vbNullString
        btnDonate.ToolTipText = vbNullString
        btnUpdate.ToolTipText = vbNullString
        btnGithubHome.ToolTipText = vbNullString
        
        chkDpiAwareness.ToolTipText = vbNullString
        'optEnablePrefsTooltips.ToolTipText = vbNullString
        
        optPrefsTooltips(0).ToolTipText = vbNullString
        optPrefsTooltips(1).ToolTipText = vbNullString
        optPrefsTooltips(2).ToolTipText = vbNullString
        
        btnResetMessages.ToolTipText = vbNullString
        
        sliSkewDegrees.ToolTipText = vbNullString
    
    End If

   On Error GoTo 0
   Exit Sub

setPrefsTooltips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsTooltips of Form widgetPrefs"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setPrefsLabels
' Author    : beededea
' Date      : 27/09/2023
' Purpose   : set the text in any labels that need a vbCrLf to space the text
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsLabels()

    On Error GoTo setPrefsLabels_Error

    lblFontsTab(0).Caption = "When resizing the form (drag bottom right) the font size will in relation to form height. The base font determines the initial size, the resulting resized font will dynamically change." & vbCrLf & vbCrLf & _
        "My preferred font for this utility is Centurion Light SF at 8pt size."

    On Error GoTo 0
    Exit Sub

setPrefsLabels_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsLabels of Form widgetPrefs"
        
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DestroyToolTip
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : It's not a bad idea to put this in the Form_Unload event just to make sure.
'---------------------------------------------------------------------------------------
'
Public Sub DestroyToolTip()
    
   On Error GoTo DestroyToolTip_Error

    If hwndTT <> 0& Then DestroyWindow hwndTT
    hwndTT = 0&

   On Error GoTo 0
   Exit Sub

DestroyToolTip_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DestroyToolTip of Form widgetPrefs"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : loadPrefsAboutText
' Author    : beededea
' Date      : 12/03/2020
' Purpose   : The text for the about page is stored here
'---------------------------------------------------------------------------------------
'
Private Sub loadPrefsAboutText()
    On Error GoTo loadPrefsAboutText_Error
    'If giDebugFlg = 1 Then Debug.Print "%loadPrefsAboutText"
    
    lblMajorVersion.Caption = App.Major
    lblMinorVersion.Caption = App.Minor
    lblRevisionNum.Caption = App.Revision
    
    lblAbout(1).Caption = "(32bit WoW64 using " & gsCodingEnvironment & " + " & gsRichClientEnvironment & ")"
    
    Call LoadFileToTB(txtAboutText, App.Path & "\resources\txt\about.txt", False)
    
    
    
   On Error GoTo 0
   Exit Sub

loadPrefsAboutText_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPrefsAboutText of Form widgetPrefs"
    
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : picButtonMouseUpEvent
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : capture the icon button clicks avoiding creating a control array
'---------------------------------------------------------------------------------------
'
Private Sub picButtonMouseUpEvent(ByVal thisTabName As String, ByRef thisPicName As image, ByRef thisPicNameClicked As image, ByRef thisFraName As Frame, Optional ByRef thisFraButtonName As Frame)
    
    On Error GoTo picButtonMouseUpEvent_Error
    
    Dim padding As Long: padding = 0
    Dim BorderWidth As Long: BorderWidth = 0
    Dim captionHeight As Long: captionHeight = 0
    Dim y_scale As Single: y_scale = 0
    
    thisPicNameClicked.Visible = False
    thisPicName.Visible = True
      
    btnSave.Visible = False
    btnClose.Visible = False
    btnHelp.Visible = False
    
    Call clearBorderStyle

    gsLastSelectedTab = thisTabName
    sPutINISetting "Software\TenShillings", "lastSelectedTab", gsLastSelectedTab, gsSettingsFile

    thisFraName.Visible = True
    
    thisFraButtonName.BorderStyle = 1

    #If twinbasic Then
        thisFraButtonName.Refresh
    #End If

    ' Get the form's current scale factors.
    y_scale = Me.ScaleHeight / gdPrefsStartHeight
    
    If gsDpiAwareness = "1" Then
        btnHelp.Top = fraGeneral.Top + fraGeneral.height + (100 * y_scale)
    Else
        btnHelp.Top = thisFraName.Top + thisFraName.height + (200 * y_scale)
    End If
    
    btnSave.Top = btnHelp.Top
    btnClose.Top = btnSave.Top
    
    btnSave.Visible = True
    btnClose.Visible = True
    btnHelp.Visible = True
    
    lblAsterix.Top = btnSave.Top + 50
    lblSize.Top = lblAsterix.Top - 300
    
    chkEnableResizing.Top = btnSave.Top + 50
    'chkEnableResizing.Left = lblAsterix.Left
    
    BorderWidth = (widgetPrefs.Width - Me.ScaleWidth) / 2
    captionHeight = widgetPrefs.height - Me.ScaleHeight - BorderWidth
        
    ' under windows 10+ the internal window calcs are all wrong due to the bigger title bars
    If pPrefsDynamicSizingFlg = False Then
        padding = 200 ' add normal padding below the help button to position the bottom of the form

        pLastFormHeight = btnHelp.Top + btnHelp.height + captionHeight + BorderWidth + padding
        gbPrefsFormResizedInCode = True
        widgetPrefs.height = pLastFormHeight
    End If
    
    If gsDpiAwareness = "0" Then
        If thisTabName = "about" Then
            lblAsterix.Visible = False
            chkEnableResizing.Visible = True
        Else
            lblAsterix.Visible = True
            chkEnableResizing.Visible = False
        End If
    End If
    
   On Error GoTo 0
   Exit Sub

picButtonMouseUpEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure picButtonMouseUpEvent of Form widgetPrefs"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : themeTimer_Timer
' Author    : beededea
' Date      : 13/06/2020
' Purpose   : a timer to apply a theme automatically
'---------------------------------------------------------------------------------------
'
Private Sub themeTimer_Timer()
        
    Dim SysClr As Long: SysClr = 0

    On Error GoTo themeTimer_Timer_Error
    
    If widgetPrefs.IsVisible = False Then Exit Sub

    SysClr = GetSysColor(COLOR_BTNFACE)

    If SysClr <> glStoreThemeColour Then
        Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

themeTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure themeTimer_Timer of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : main menu item to buy the developer a coffee from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuCoffee_Click()
    On Error GoTo mnuCoffee_Click_Error
    
    Call mnuCoffee_ClickEvent

    On Error GoTo 0
    Exit Sub
mnuCoffee_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_Click of Form widgetPrefs"
End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : mnuLicenceA_Click
' Author    : beededea
' Date      : 17/08/2022
' Purpose   : menu option to show licence from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuLicenceA_Click()
    On Error GoTo mnuLicenceA_Click_Error

    Call mnuLicence_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuLicenceA_Click_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicenceA_Click of Form widgetPrefs"
            Resume Next
          End If
    End With

End Sub



'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_Click
' Author    : beededea
' Date      : 13/02/2019
' Purpose   : menu option to open support page from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuSupport_Click()
    
    On Error GoTo mnuSupport_Click_Error

    Call mnuSupport_ClickEvent

    On Error GoTo 0
    Exit Sub

mnuSupport_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_Click of Form widgetPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : mnuClosePreferences_Click
' Author    : beededea
' Date      : 06/09/2024
' Purpose   : right click close option from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuClosePreferences_Click()
   On Error GoTo mnuClosePreferences_Click_Error

    Call btnClose_Click

   On Error GoTo 0
   Exit Sub

mnuClosePreferences_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuClosePreferences_Click of Form widgetPrefs"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuAuto_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   : right click auto theme option from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuAuto_Click()
    
   On Error GoTo mnuAuto_Click_Error

    If themeTimer.Enabled = True Then
        MsgBox "Automatic Theme Selection is now Disabled"
        mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
        mnuAuto.Checked = False
        
        themeTimer.Enabled = False
    Else
        MsgBox "Auto Theme Selection Enabled. If the o/s theme changes the utility should automatically skin the utility to suit the theme."
        mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
        mnuAuto.Checked = True
        
        themeTimer.Enabled = True
        Call setThemeColour
    End If

   On Error GoTo 0
   Exit Sub

mnuAuto_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuAuto_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuDark_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   : right click dark theme option from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuDark_Click()
   On Error GoTo mnuDark_Click_Error

    mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    mnuAuto.Checked = False
    mnuDark.Caption = "Dark Theme Enabled"
    mnuLight.Caption = "Light Theme Enable"
    themeTimer.Enabled = False
    
    gsSkinTheme = "dark"

    Call setThemeShade(212, 208, 199)

   On Error GoTo 0
   Exit Sub

mnuDark_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuDark_Click of Form widgetPrefs"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuLight_Click
' Author    : beededea
' Date      : 19/05/2020
' Purpose   : right click light theme option from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub mnuLight_Click()
    'MsgBox "Auto Theme Selection Manually Disabled"
   On Error GoTo mnuLight_Click_Error
    
    mnuAuto.Caption = "Auto Theme Disabled - Click to Enable"
    mnuAuto.Checked = False
    mnuDark.Caption = "Dark Theme Enable"
    mnuLight.Caption = "Light Theme Enabled"
    themeTimer.Enabled = False
    
    gsSkinTheme = "light"

    Call setThemeShade(240, 240, 240)

   On Error GoTo 0
   Exit Sub

mnuLight_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLight_Click of Form widgetPrefs"
End Sub




'
'---------------------------------------------------------------------------------------
' Procedure : setThemeShade
' Author    : beededea
' Date      : 06/05/2023
' Purpose   : set the theme shade, Windows classic dark/new lighter theme colours from the prefs-specific pop-up menu
'---------------------------------------------------------------------------------------
'
Private Sub setThemeShade(ByVal redC As Integer, ByVal greenC As Integer, ByVal blueC As Integer)
    
    Dim Ctrl As Control
    
    On Error GoTo setThemeShade_Error

    ' RGB(redC, greenC, blueC) is the background colour used by the lighter themes
    
    Me.BackColor = RGB(redC, greenC, blueC)
    
    ' all buttons must be set to graphical
    For Each Ctrl In Me.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          Ctrl.BackColor = RGB(redC, greenC, blueC)
        End If
    Next
    
    If redC = 212 Then
        'classicTheme = True
        mnuLight.Checked = False
        mnuDark.Checked = True
        
        Call setPrefsIconImagesDark
        
    Else
        'classicTheme = False
        mnuLight.Checked = True
        mnuDark.Checked = False
        
        Call setPrefsIconImagesLight
                
    End If
    
    'now change the color of the sliders.
'    widgetPrefs.sliAnimationInterval.BackColor = RGB(redC, greenC, blueC)
    'widgetPrefs.'sliWidgetSkew.BackColor = RGB(redC, greenC, blueC)
    sliWidgetSize.BackColor = RGB(redC, greenC, blueC)
    sliOpacity.BackColor = RGB(redC, greenC, blueC)
    sliSkewDegrees.BackColor = RGB(redC, greenC, blueC)
    txtAboutText.BackColor = RGB(redC, greenC, blueC)
    
    sPutINISetting "Software\TenShillings", "skinTheme", gsSkinTheme, gsSettingsFile

    On Error GoTo 0
    Exit Sub

setThemeShade_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeShade of Module Module1"
            Resume Next
          End If
    End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setThemeColour
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : if the o/s is capable of supporting the classic theme it tests every 10 secs
'             to see if a theme has been switched
'
'---------------------------------------------------------------------------------------
'
Private Sub setThemeColour()
    
    Dim SysClr As Long: SysClr = 0
    
   On Error GoTo setThemeColour_Error
   'If giDebugFlg = 1  Then Debug.Print "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeShade(212, 208, 199)
        SysClr = GetSysColor(COLOR_BTNFACE)
        gsSkinTheme = "dark"
        
        mnuDark.Caption = "Dark Theme Enabled"
        mnuLight.Caption = "Light Theme Enable"

    Else
        Call setModernThemeColours
        mnuDark.Caption = "Dark Theme Enable"
        mnuLight.Caption = "Light Theme Enabled"
    End If

    glStoreThemeColour = SysClr

   On Error GoTo 0
   Exit Sub

setThemeColour_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeColour of module module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : adjustPrefsTheme
' Author    : beededea
' Date      : 25/04/2023
' Purpose   : adjust the theme used by the prefs alone
'---------------------------------------------------------------------------------------
'
Private Sub adjustPrefsTheme()
   On Error GoTo adjustPrefsTheme_Error

    If gsSkinTheme <> vbNullString Then
        If gsSkinTheme = "dark" Then
            Call setThemeShade(212, 208, 199)
        Else
            Call setThemeShade(240, 240, 240)
        End If
    Else
        If gbClassicThemeCapable = True Then
            mnuAuto.Caption = "Auto Theme Enabled - Click to Disable"
            themeTimer.Enabled = True
        Else
            gsSkinTheme = "light"
            Call setModernThemeColours
        End If
    End If
    
    

   On Error GoTo 0
   Exit Sub

adjustPrefsTheme_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustPrefsTheme of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setModernThemeColours
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : by 'modern theme' we mean the very light indeed almost white background to standard Windows forms...
'---------------------------------------------------------------------------------------
'
Private Sub setModernThemeColours()
         
    Dim SysClr As Long: SysClr = 0
    
    On Error GoTo setModernThemeColours_Error
    
    'the widgetPrefs.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"

    'MsgBox "Windows Alternate Theme detected"
    SysClr = GetSysColor(COLOR_BTNFACE)
    If SysClr = 13160660 Then
        Call setThemeShade(212, 208, 199)
        gsSkinTheme = "dark"
    Else ' 15790320
        Call setThemeShade(240, 240, 240)
        gsSkinTheme = "light"
    End If

   On Error GoTo 0
   Exit Sub

setModernThemeColours_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setModernThemeColours of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : loadHigherResPrefsImages
' Author    : beededea
' Date      : 18/06/2023
' Purpose   : load the images for the classic or high brightness themes
'---------------------------------------------------------------------------------------
'
Private Sub loadHigherResPrefsImages()
    
    On Error GoTo loadHigherResPrefsImages_Error
      
    If Me.WindowState = vbMinimized Then Exit Sub
    
'
'    If dynamicSizingFlg = False Then
'        Exit Sub
'    End If
'
'    If Me.Width < 10500 Then
'        topIconWidth = 600
'    End If
'
'    If Me.Width >= 10500 And Me.Width < 12000 Then 'Me.Height / ratio ' maintain the aspect ratio
'        topIconWidth = 730
'    End If
'
'    If Me.Width >= 12000 And Me.Width < 13500 Then 'Me.Height / ratio ' maintain the aspect ratio
'        topIconWidth = 834
'    End If
'
'    If Me.Width >= 13500 And Me.Width < 15000 Then 'Me.Height / ratio ' maintain the aspect ratio
'        topIconWidth = 940
'    End If
'
'    If Me.Width >= 15000 Then 'Me.Height / ratio ' maintain the aspect ratio
'        topIconWidth = 1010
'    End If
        
    If mnuDark.Checked = True Then
        Call setPrefsIconImagesDark
    Else
        Call setPrefsIconImagesLight
    End If
    
    
    
   On Error GoTo 0
   Exit Sub

loadHigherResPrefsImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadHigherResPrefsImages of Form widgetPrefs"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : tmrWritePositionAndSize_Timer
' Author    : beededea
' Date      : 27/05/2023
' Purpose   : periodically read the prefs form position/height and store
'---------------------------------------------------------------------------------------
'
Private Sub tmrWritePositionAndSize_Timer()
    
    On Error GoTo tmrWritePositionAndSize_Timer_Error
   
    ' save the current X and y position of this form to allow repositioning when restarting
    If widgetPrefs.IsVisible = True Then Call writePrefsPositionAndSize

   On Error GoTo 0
   Exit Sub

tmrWritePositionAndSize_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrWritePositionAndSize_Timer of Form widgetPrefs"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : chkEnableResizing_Click
' Author    : beededea
' Date      : 27/05/2023
' Purpose   : toggle to enable sizing when in low DPI aware mode
'---------------------------------------------------------------------------------------
'
Private Sub chkEnableResizing_Click()
   On Error GoTo chkEnableResizing_Click_Error

    If chkEnableResizing.Value = 1 Then
        pPrefsDynamicSizingFlg = True
        txtPrefsFontCurrentSize.Visible = True
        'lblCurrentFontsTab.Visible = True
        'Call writePrefsPositionAndSize
        chkEnableResizing.Caption = "Disable Corner Resizing"
    Else
        pPrefsDynamicSizingFlg = False
        txtPrefsFontCurrentSize.Visible = False
        'lblCurrentFontsTab.Visible = False
        Unload widgetPrefs
        Me.Show
        Call readPrefsPosition
        chkEnableResizing.Caption = "Enable Corner Resizing"
    End If
    
    Call setframeHeights

   On Error GoTo 0
   Exit Sub

chkEnableResizing_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkEnableResizing_Click of Form widgetPrefs"

End Sub


 



'---------------------------------------------------------------------------------------
' Procedure : setframeHeights
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : set the frame heights to manual sizes for the low DPI mode as per the YWE prefs
'---------------------------------------------------------------------------------------
'
Private Sub setframeHeights()
   On Error GoTo setframeHeights_Error

    If pPrefsDynamicSizingFlg = True Then
        fraGeneral.height = fraAbout.height
        fraFonts.height = fraAbout.height
        fraConfig.height = fraAbout.height
        fraSounds.height = fraAbout.height
        fraPosition.height = fraAbout.height
        fraDevelopment.height = fraAbout.height
        fraWindow.height = fraAbout.height
        
        fraGeneral.Width = fraAbout.Width
        fraFonts.Width = fraAbout.Width
        fraConfig.Width = fraAbout.Width
        fraSounds.Width = fraAbout.Width
        fraPosition.Width = fraAbout.Width
        fraDevelopment.Width = fraAbout.Width
        fraWindow.Width = fraAbout.Width
    
        'If gsDpiAwareness = "1" Then
            ' save the initial positions of ALL the controls on the prefs form
            Call SaveSizes(widgetPrefs, gcPrefsControlPositions(), gdPrefsStartWidth, gdPrefsStartHeight)
        'End If
    Else
        fraGeneral.height = 2205
        fraConfig.height = 8777
        fraSounds.height = 3985
        fraPosition.height = 7544
        fraFonts.height = 5643
        
'        fraWindow.Height = 8700
'        fraWindowInner.Height = 8085
        
'        ' the lowest window controls are not displayed on a single monitor
        If glMonitorCount > 1 Then
            fraWindow.height = 8700
            fraWindowInner.height = 8085
        Else
            fraWindow.height = 7225
            fraWindowInner.height = 6345
        End If

        fraDevelopment.height = 6297
        fraAbout.height = 8700
    End If
    
   On Error GoTo 0
   Exit Sub

setframeHeights_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setframeHeights of Form widgetPrefs"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesDark
' Author    : beededea
' Date      : 22/06/2023
' Purpose   : set the bright images for the grey classic theme
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsIconImagesDark()

Dim count As Long
    
    On Error GoTo setPrefsIconImagesDark_Error
    
        ' setting the Prefs tab Jpeg icon images to the GDIP imageList dictionary, previously used Cairo.ImageList("about-icon-dark-clicked").Picture

'        thisImageList.ImageHeight = 0
'        thisImageList.ImageWidth = 0

        ' normal images
        Set imgGeneral.Picture = thisImageList.Picture("general-icon-dark")
        Set imgConfig.Picture = thisImageList.Picture("config-icon-dark")
        Set imgFonts.Picture = thisImageList.Picture("font-icon-dark")
        Set imgSounds.Picture = thisImageList.Picture("sounds-icon-dark")
        Set imgPosition.Picture = thisImageList.Picture("position-icon-dark")
        Set imgDevelopment.Picture = thisImageList.Picture("development-icon-dark")
        Set imgWindow.Picture = thisImageList.Picture("windows-icon-dark")
        Set imgAbout.Picture = thisImageList.Picture("about-icon-dark")
        
        ' clicked images
        Set imgGeneralClicked.Picture = thisImageList.Picture("general-icon-dark-clicked")
        Set imgConfigClicked.Picture = thisImageList.Picture("config-icon-dark-clicked")
        Set imgFontsClicked.Picture = thisImageList.Picture("font-icon-dark-clicked")
        Set imgSoundsClicked.Picture = thisImageList.Picture("sounds-icon-dark-clicked")
        Set imgPositionClicked.Picture = thisImageList.Picture("position-icon-dark-clicked")
        Set imgDevelopmentClicked.Picture = thisImageList.Picture("development-icon-dark-clicked")
        Set imgWindowClicked.Picture = thisImageList.Picture("windows-icon-dark-clicked")
        Set imgAboutClicked.Picture = thisImageList.Picture("about-icon-dark-clicked")

   On Error GoTo 0
   Exit Sub

setPrefsIconImagesDark_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesDark of Form widgetPrefs"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : setPrefsIconImagesLight
' Author    : beededea
' Date      : 22/06/2023
' Purpose   : set the bright images for the bright 'modern' theme
'---------------------------------------------------------------------------------------
'
Private Sub setPrefsIconImagesLight()
    
    On Error GoTo setPrefsIconImagesLight_Error
    
        ' setting the Prefs tab Jpeg icon images to the GDIP imageList dictionary, previously used Cairo.ImageList("about-icon-dark-clicked").Picture

        ' normal images
        Set imgGeneral.Picture = thisImageList.Picture("general-icon-light")
        Set imgConfig.Picture = thisImageList.Picture("config-icon-light")
        Set imgFonts.Picture = thisImageList.Picture("font-icon-light")
        Set imgSounds.Picture = thisImageList.Picture("sounds-icon-light")
        Set imgPosition.Picture = thisImageList.Picture("position-icon-light")
        Set imgDevelopment.Picture = thisImageList.Picture("development-icon-light")
        Set imgWindow.Picture = thisImageList.Picture("windows-icon-light")
        Set imgAbout.Picture = thisImageList.Picture("about-icon-light")
        
        ' clicked images
        Set imgGeneralClicked.Picture = thisImageList.Picture("general-icon-light-clicked")
        Set imgConfigClicked.Picture = thisImageList.Picture("config-icon-light-clicked")
        Set imgFontsClicked.Picture = thisImageList.Picture("font-icon-light-clicked")
        Set imgSoundsClicked.Picture = thisImageList.Picture("sounds-icon-light-clicked")
        Set imgPositionClicked.Picture = thisImageList.Picture("position-icon-light-clicked")
        Set imgDevelopmentClicked.Picture = thisImageList.Picture("development-icon-light-clicked")
        Set imgWindowClicked.Picture = thisImageList.Picture("windows-icon-light-clicked")
        Set imgAboutClicked.Picture = thisImageList.Picture("about-icon-light-clicked")
        
   On Error GoTo 0
   Exit Sub

setPrefsIconImagesLight_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setPrefsIconImagesLight of Form widgetPrefs"

End Sub





'---------------------------------------------------------------------------------------
' Procedure : tmrPrefsMonitorSaveHeight_Timer
' Author    : beededea
' Date      : 26/08/2024
' Purpose   : save the current height of this form to allow resizing when restarting or placing on another monitor
'---------------------------------------------------------------------------------------
'
'Private Sub tmrPrefsMonitorSaveHeight_Timer()
'
'    'Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
'
'    On Error GoTo tmrPrefsMonitorSaveHeight_Timer_Error
'
'    If widgetPrefs.IsVisible = False Then Exit Sub
'
'    If LTrim$(gsMultiMonitorResize) <> "2" Then Exit Sub
'
'    If gPrefsMonitorStruct.IsPrimary = True Then
'        gsPrefsPrimaryHeightTwips = CStr(widgetPrefs.Height)
'        sPutINISetting "Software\TenShillings", "prefsPrimaryHeightTwips", gsPrefsPrimaryHeightTwips, gsSettingsFile
'    Else
'        gsPrefsSecondaryHeightTwips = CStr(widgetPrefs.Height)
'        sPutINISetting "Software\TenShillings", "prefsSecondaryHeightTwips", gsPrefsSecondaryHeightTwips, gsSettingsFile
'    End If
'
'   On Error GoTo 0
'   Exit Sub
'
'tmrPrefsMonitorSaveHeight_Timer_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrPrefsMonitorSaveHeight_Timer of Form widgetPrefs"
'
'End Sub





'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'  --- All folded content will be temporary put under this lines ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'CODEFOLD STORAGE:
'CODEFOLD STORAGE END:
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'--- If you're Subclassing: Move the CODEFOLD STORAGE up as needed ---
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\




