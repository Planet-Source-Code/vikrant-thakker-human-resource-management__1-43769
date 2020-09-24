VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPayDeduct 
   Appearance      =   0  'Flat
   Caption         =   "Pay Deduction/Cash Exchange"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid Dg 
      Height          =   7080
      Left            =   1755
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   12488
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   0
      ForeColor       =   12648447
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Employee List"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameDed_Exch 
      Height          =   6975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12135
      Begin VB.Frame FrameDed 
         Appearance      =   0  'Flat
         Caption         =   "Deduction Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3435
         Left            =   120
         TabIndex        =   42
         Top             =   3480
         Width           =   5925
         Begin VB.TextBox txtExcess 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            ToolTipText     =   "if Leave days > 2 then  No. of leave days - 2 "
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtTotAmtDed 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            ToolTipText     =   "Excess Leave * DedPerDay"
            Top             =   2175
            Width           =   1935
         End
         Begin VB.TextBox txtDedPerDay 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Section/ 30"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtNOD 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "EXCESS LEAVE DAYS"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   540
            TabIndex        =   50
            Top             =   1080
            Width           =   1665
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL AMT. DEDUCTED"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   300
            TabIndex        =   49
            Top             =   2280
            Width           =   1905
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "DEDUCTION PER DAY"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   480
            TabIndex        =   48
            Top             =   1680
            Width           =   1710
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NO. OF LEAVES TAKEN"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   420
            TabIndex        =   47
            Top             =   480
            Width           =   1785
         End
      End
      Begin VB.Frame FrameExch 
         Appearance      =   0  'Flat
         Caption         =   "Exchange Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3435
         Left            =   6000
         TabIndex        =   31
         Top             =   3480
         Width           =   5985
         Begin VB.TextBox txtLeaveTaken 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   960
            Width           =   1155
         End
         Begin VB.TextBox txtTotLeavesAllowed 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "2"
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox txtBalDays 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtAmtPerDay 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txtTotAmtAllowed 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BALANCE DAYS"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   1065
            TabIndex        =   41
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL LEAVES ALLOWED"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   300
            TabIndex        =   40
            Top             =   480
            Width           =   2025
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "LEAVES TAKEN"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   1095
            TabIndex        =   39
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "AMOUNT PER DAY"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   795
            TabIndex        =   38
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL AMOUNT ALLOWED"
            ForeColor       =   &H000080FF&
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   2760
            Width           =   2115
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Employee Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3420
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   11865
         Begin VB.ComboBox cmbDummyYear 
            Height          =   315
            Left            =   9240
            TabIndex        =   51
            Text            =   "DummyYear"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.ComboBox txtMnt 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7605
            TabIndex        =   22
            Top             =   2760
            Width           =   1635
         End
         Begin VB.TextBox txtSection 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "BASIC"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7605
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   405
            Width           =   1575
         End
         Begin VB.ComboBox cmbYear 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "frmPayded.frx":0000
            Left            =   7605
            List            =   "frmPayded.frx":0002
            TabIndex        =   21
            Top             =   1935
            Width           =   1590
         End
         Begin VB.CommandButton cmdHelp 
            BackColor       =   &H0041E9D8&
            Caption         =   "&List"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3150
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   405
            Width           =   780
         End
         Begin VB.TextBox txtDesig 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "DESIG"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1965
            Width           =   3255
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "NAME"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1185
            Width           =   3495
         End
         Begin VB.TextBox txtCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "SNO"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1710
            TabIndex        =   17
            Top             =   405
            Width           =   1335
         End
         Begin VB.TextBox txtBasicPay 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "BASIC"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7590
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1215
            Width           =   1575
         End
         Begin VB.OptionButton optDeduction 
            Caption         =   "DEDUCTION"
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   225
            TabIndex        =   15
            Top             =   2715
            Width           =   1455
         End
         Begin VB.OptionButton optExchange 
            Caption         =   "EXCHANGE"
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   2025
            TabIndex        =   14
            Top             =   2715
            Width           =   1545
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SECTION"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   6660
            TabIndex        =   30
            Top             =   405
            Width           =   705
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "YEAR"
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   6870
            TabIndex        =   29
            Top             =   1980
            Width           =   555
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "DESIGNATION"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   315
            TabIndex        =   28
            Top             =   1965
            Width           =   1110
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NAME"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   960
            TabIndex        =   27
            Top             =   1245
            Width           =   465
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CODE"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   975
            TabIndex        =   26
            Top             =   525
            Width           =   450
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BASIC PAY"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   6510
            TabIndex        =   25
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "MONTH"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   6690
            TabIndex        =   24
            Top             =   2790
            Width           =   600
         End
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0041E9D8&
      Caption         =   "E&xit"
      Height          =   420
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7890
      Width           =   870
   End
   Begin VB.Frame Framebutton 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   2700
      TabIndex        =   3
      Top             =   7215
      Width           =   6495
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cancel"
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modify"
         Height          =   420
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   915
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H0080C0FF&
         Caption         =   "Save"
         Height          =   420
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H0080C0FF&
         Caption         =   "Next"
         Height          =   420
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdPrev 
         BackColor       =   &H0080C0FF&
         Caption         =   "Previous"
         Height          =   420
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H0080C0FF&
         Caption         =   "Remove"
         Height          =   420
         Left            =   3690
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   870
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add"
         Height          =   420
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0041E9D8&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7185
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pay  Deduction /  Cash exchange for earned leave"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   1710
      TabIndex        =   0
      Top             =   405
      Width           =   7320
   End
End
Attribute VB_Name = "frmPayDeduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsEmpList As Recordset
Dim Modify, Add, ValidCode As Boolean

' BOOLEANS DECLARED AND THE FUNCTIONS

' Modify = True if Modify button is clicked
' Add = True if Add button is clicked
' ValidCode = True if the employee code entered is valid (does exist)

' USER DEFINED FUNCTIONS AND THEIR PURPOSE

' EnableAll  : enable all the text boxes
' DisableAll : disable all the text boxes
' ClearAll   : Clear (Blank) all the text boxes
' Modi       : If modify = True then Enables all the text boxes
'              If modify = False then Disable all the text boxes
' TotalLeaveDays : This function calculates the Total No. Of Leaves by an employee in a particular month

Private Sub Search()
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
    For i = 0 To rsEmp.RecordCount - 1 Step 1
        If (txtCode = rsEmp!Code) Then ' If the employee code is found then
        'Enter the data of employee in respected text boxes
            If Not IsNull(rsEmp!Name) Then txtName.Text = rsEmp!Name
            If Not IsNull(rsEmp!Desig) Then txtDesig.Text = rsEmp!Desig
            If Not IsNull(rsEmp!Sect) Then txtSection.Text = rsEmp!Sect
            If Not IsNull(rsEmp!BasicPay) Then txtBasicPay.Text = rsEmp!BasicPay
            SendKeys "{TAB}"
            KeyAscii = 0

   If Add = False Then 'If add Button is not pressed
        cmdModify.Enabled = True
        cmdRemove.Enabled = True
   End If
            Exit Sub
        End If
    rsEmp.MoveNext
    If rsEmp.EOF = True Then
        MsgBox "Invalid Employee Code !", vbOKOnly, "OASYS"
        Call ClearAll
        txtCode.SetFocus
        Exit Sub
    End If
    Next
End Sub

'This function is to Check whether the entered Empoyee Code is Valid (Exists or not)
Private Sub ValidateCode()
If rsEmp.RecordCount > 0 Then rsEmp.MoveFirst
    For i = 0 To rsEmp.RecordCount - 1 Step 1
        If (txtCode = rsEmp!Code) Then
        ValidCode = True 'ValidCode = True if Employee Code does exist
            Exit Sub
        End If
    rsEmp.MoveNext
    If rsEmp.EOF = True Then
    ValidCode = False
        MsgBox "Invalid Employee Code !", vbOKOnly, "OASYS"
        Exit Sub
    End If
    Next
End Sub

Private Sub cmbYear_GotFocus()
Call txt_GotFocus
End Sub

Private Sub cmbYear_KeyPress(KeyAscii As Integer)
txtMnt.Clear
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
Call ValidateYears
End If
End Sub
Private Sub AddMonths()
If Trim(cmbYear.Text) <> "" Then
    If rsWD.RecordCount > 0 Then rsWD.MoveFirst
    For i = 0 To rsWD.RecordCount - 1 Step 1
        If cmbYear.Text = rsWD!Year Then
            txtMnt.AddItem (rsWD!Month)
        End If
        If rsWD.EOF = False Then rsWD.MoveNext
    Next
End If
End Sub
Private Sub cmbYear_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo aerr
Modify = False
Add = True   ' As Add button is clicked, Boolean Add = True

Call Modi
Call EnableAll  ' Enable all the text boxes
Call ClearAll   ' Blank all the text boxes

'rsPD.AddNew

cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdModify.Enabled = False
cmdRemove.Enabled = False
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdClose.Enabled = False

txtCode.SetFocus
cmdSave.Enabled = False
Exit Sub
aerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo cerr
Modify = False
Add = False
Call Modi
'rsPD.CancelUpdate
Call DisableAll   ' Disable all the text boxes

cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdRemove.Enabled = True
cmdCancel.Enabled = False
cmdSave.Enabled = False
cmdClose.Enabled = True

Call cmdPrev_Click
Exit Sub
cerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdClose_Click()
On Error GoTo eerr
frmMain.Show
Unload Me
Exit Sub
eerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdHelp_Click()  ' This is to show the Data Grid
Framebutton.Visible = False
FrameDed.Visible = False
FrameExch.Visible = False
Frame3.Visible = False
cmdClose.Visible = False
dg.Refresh
dg.Visible = True
cmdOK.Visible = True
End Sub

Private Sub cmdModify_Click()
On Error GoTo merr
MsgBox "You cannot modify this data !", vbOKOnly, "OASYS"
Exit Sub
Modify = True  ' Modify button is clicked
Add = False
Call Modi
Call EnableAll  ' Enable all the text boxes

cmdCancel.Enabled = True
cmdSave.Enabled = True
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdAdd.Enabled = False
cmdRemove.Enabled = False
cmdModify.Enabled = False
cmdClose.Enabled = False

txtCode.SetFocus
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdNext_Click()
'On Error GoTo nerr
Modify = False
Add = False
Call Modi

If rsPD.RecordCount = 0 Then  'If there are no records (data) in the table
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    Exit Sub
Else
    cmdModify.Enabled = True
End If
If rsPD.EOF = False Then rsPD.MoveNext
If rsPD.EOF = True Then rsPD.MoveLast
    showall
Exit Sub
nerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdOK_Click()   ' Hide the Data grid
FrameDed.Visible = True
FrameExch.Visible = True
Frame3.Visible = True
Framebutton.Visible = True
cmdClose.Visible = True
cmdOK.Visible = False
dg.Visible = False
End Sub

Private Sub cmdPrev_Click()
On Error GoTo perr
Modify = False
Add = False
Call Modi

If rsPD.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
If rsPD.BOF = False Then rsPD.MovePrevious
If rsPD.BOF = True Then rsPD.MoveFirst
showall
    cmdRemove.Enabled = True
Exit Sub
perr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo rerr
Modify = False
Add = False
Call Modi
 If rsPD.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    cmdSave.Enabled = False
    Exit Sub
End If
       rsPD.Delete
       Call cmdNext_Click
Exit Sub
rerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub cmdSave_Click()
On Error GoTo serr
Add = False

If Trim(txtCode.Text) = "" Or Trim(txtName.Text) = "" Or Trim(txtMnt.Text) = "" Then
    Exit Sub
End If

Call ValidateCode
If ValidCode = False Then Exit Sub

rsPD.AddNew
If txtCode.Text <> "" Then rsPD!Code = txtCode.Text
If txtName.Text <> "" Then rsPD!Name = txtName.Text
If txtDesig.Text <> "" Then rsPD!Desig = txtDesig.Text
If txtSection.Text <> "" Then rsPD!Sect = txtSection.Text
If txtBasicPay.Text <> "" Then rsPD!BasicPay = txtBasicPay.Text
If cmbYear.Text <> "" Then rsPD!Year = cmbYear.Text
If txtMnt.Text <> "" Then rsPD!Month = txtMnt.Text
If optDeduction.Value = True Then
    rsPD!Type = "Deduction"
    If txtNOD.Text <> "" Then rsPD!noofdays = txtNOD.Text
    If txtDedPerDay.Text <> "" Then rsPD!DedPerDay = txtDedPerDay.Text
    If txtTotAmtDed.Text <> "" Then rsPD!totamtded = txtTotAmtDed.Text
ElseIf optExchange.Value = True Then
    rsPD!Type = "Exchange"
    If txtTotLeavesAllowed.Text <> "" Then rsPD!totleavesallowed = txtTotLeavesAllowed.Text
    If txtLeaveTaken.Text <> "" Then rsPD!LeavesTaken = txtLeaveTaken.Text
    If txtBalDays.Text <> "" Then rsPD!BalDays = txtBalDays.Text
    If txtAmtPerDay.Text <> "" Then rsPD!amtperday = txtAmtPerDay.Text
    If txtTotAmtAllowed.Text <> "" Then rsPD!TotAmtAllowed = txtTotAmtAllowed.Text
End If

If Year(Date) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
ElseIf Year(Date) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "OASYS"
Exit Sub
End If


rsPD.Update

cmdSave.Enabled = False
cmdAdd.Enabled = True
cmdModify.Enabled = True
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdCancel.Enabled = False
cmdRemove.Enabled = True
cmdClose.Enabled = True
cmdSave.Enabled = False

Call DisableAll
cmdAdd.SetFocus

If cmdAdd.Enabled = True Then cmdAdd.SetFocus
Exit Sub
serr:
MsgBox Err.Description, vbOKOnly, "OASYS"

Call EnableAll
Call ClearAll

txtCode.SetFocus
End Sub

Private Sub AddYears()
cmbDummyYear.Clear  'First clear the Months Combo box
    If rsWD.RecordCount > 0 Then rsWD.MoveFirst
    For i = 0 To rsWD.RecordCount - 1 Step 1
            cmbDummyYear.AddItem (rsWD!Year)  ' Add Months in the combo box
        If rsWD.EOF = False Then rsWD.MoveNext
    Next
End Sub

Private Sub ClearYears()
cmbYear.Clear
If cmbDummyYear.ListCount > 0 Then
'MsgBox cmbdummyYear.ListCount
For i = 0 To cmbDummyYear.ListCount Step 1
restart:
If i = cmbDummyYear.ListCount Then Exit Sub

YR = cmbDummyYear.List(i)

    For no = 0 To cmbYear.ListCount Step 1
        If cmbYear.List(no) = YR Then
            i = i + 1  'if year already exists in the list
            no = 0
            GoTo restart
        End If
        If no = cmbYear.ListCount Then
            cmbYear.AddItem (YR) ' if year does not exist then add it
        End If
    Next no
Next i
End If
End Sub


Private Sub Form_Activate()
cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
'On Error GoTo ferr
Dim dt As String, tm As String
    dt = Format(Date, "dd-mmm-yyyy")
    tm = Format(Now, "HH:MM:SS")
    frmPayDeduct.Caption = dt + Space(150) + tm
    optDeduction.Value = 0
    optDeduction.Value = 0
  '  FrameDed.Visible = False
  '  FrameExch.Visible = False

Set rsEmpList = New ADODB.Recordset
rsEmpList.Open "select Code,Name,Desig,Sect,BasicPay from MastEmployee", conn, adOpenStatic, adLockOptimistic
cmbYear.Text = Year(Date)

Set dg.DataSource = rsEmpList
dg.Refresh

Call AddYears
Call ClearYears
'This to add Last Three Years in the year combo box
'For i = (Year(Date) - 5) To (Year(Date)) Step 1
'    cmbYear.AddItem (i)
'Next
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    cmdCancel.Enabled = False
    
If rsPD.RecordCount > 0 Then
    cmdAdd.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
End If
    cmdSave.Enabled = False
    
Exit Sub
ferr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub


Private Sub optDeduction_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub optExchange_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAmtPerDay_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtAmtPerDay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtAmtPerDay_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtBalDays_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtBalDays_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtBalDays_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtBasicPay_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtBasicPay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtBasicPay_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtCode_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Add = True Then Call Search
End If
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtDedPerDay_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtDedPerDay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDedPerDay_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtDesig_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtDesig_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtDesig_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtMnt_LostFocus()
If txtMnt.Text <> "" Then
    Call ValidateMonths
End If
End Sub

Private Sub txtname_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtname_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub showall()
'On Error GoTo showerr
Call ClearAll   ' Clear (Blank) all the text boxes

If rsPD.RecordCount = 0 Then
    cmdModify.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdRemove.Enabled = False
    Exit Sub
ElseIf rsPD.RecordCount > 0 Then
    cmdModify.Enabled = True
    cmdPrev.Enabled = True
    cmdNext.Enabled = True
    cmdRemove.Enabled = True
End If
If Not IsNull(rsPD!Name) Then txtName.Text = rsPD!Name
If Not IsNull(rsPD!Code) Then txtCode.Text = rsPD!Code
If Not IsNull(rsPD!Desig) Then txtDesig.Text = rsPD!Desig
If Not IsNull(rsPD!Month) Then txtMnt.Text = rsPD!Month
If Not IsNull(rsPD!Sect) Then txtSection.Text = rsPD!Sect
If Not IsNull(rsPD!BasicPay) Then txtBasicPay.Text = rsPD!BasicPay
If Not IsNull(rsPD!Year) Then cmbYear.Text = rsPD!Year
If Not IsNull(rsPD!Month) Then txtMnt.Text = rsPD!Month
If rsPD!Type = "Deduction" Then
    optDeduction.Value = True
ElseIf rsPD!Type = "Exchange" Then
    optExchange.Value = True
End If
If Not IsNull(rsPD!noofdays) Then txtNOD.Text = rsPD!noofdays
If Not IsNull(rsPD!DedPerDay) Then txtDedPerDay.Text = rsPD!DedPerDay
If Not IsNull(rsPD!totamtded) Then txtTotAmtDed.Text = rsPD!totamtded
If Not IsNull(rsPD!totleavesallowed) Then txtTotLeavesAllowed.Text = rsPD!totleavesallowed
If Not IsNull(rsPD!BalDays) Then txtBalDays.Text = rsPD!BalDays
If Not IsNull(rsPD!amtperday) Then txtAmtPerDay.Text = rsPD!amtperday
If Not IsNull(rsPD!TotAmtAllowed) Then txtTotAmtAllowed.Text = rsPD!TotAmtAllowed
If txtNOD.Text > 2 Then txtExcess.Text = Val(txtNOD.Text) - 2
FrameDed.Visible = True
FrameExch.Visible = True
Exit Sub
showerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txt_GotFocus()
On Error GoTo focerr
    If Trim(txtName.Text) = "" Or Trim(txtCode.Text) = "" Or Trim(txtMnt.Text) = "" Or Trim(cmbYear.Text) = "" Then
        cmdSave.Enabled = False
        
    Else
        cmdSave.Enabled = True
    End If
Exit Sub
focerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub Modi()
On Error GoTo merr
If Modify = False Then   ' If cmdmodify is clicked then
    Call DisableAll
ElseIf Modify = True Then  ' If cmdmodify is not clicked then
    Call EnableAll
End If
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub dg_DblClick()
On Error GoTo dgerr
txtCode.Text = dg.Columns(0).Text
txtName.Text = dg.Columns(1).Text
txtDesig.Text = dg.Columns(2).Text
txtSection.Text = dg.Columns(3).Text
txtBasicPay.Text = dg.Columns(4).Text
Call cmdOK_Click
cmbYear.SetFocus
Exit Sub
dgerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub dg_KeyPress(KeyAscii As Integer)
On Error GoTo dgkerr
If KeyAscii = 13 Then
txtCode.Text = dg.Columns(0).Text
txtName.Text = dg.Columns(1).Text
txtDesig.Text = dg.Columns(2).Text
txtSection.Text = dg.Columns(3).Text
txtBasicPay.Text = dg.Columns(4).Text
Call cmdOK_Click
cmbYear.SetFocus
End If
Call cmdOK_Click
cmbYear.SetFocus

Exit Sub
dgkerr:
MsgBox Err.Description, vbOKOnly, "OASYS"
End Sub

Private Sub txtMnt_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtMnt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     SendKeys "{TAB}"
    KeyAscii = 0
    Call ValidateMonths
'    Call TotalLeaveDays
End If
End Sub
Private Sub ValidateMonths()
'This function checks if the Month entered exist in the Month combobox or not
Dim no As Integer
no = 0
If txtMnt.Text <> "" Then
    For i = 0 To txtMnt.ListCount - 1 Step 1
        If txtMnt.List(i) = txtMnt.Text Then
        txtMnt.ListIndex = i
            Call TotalLeaveDays  ' Calculate Leaves and all other stuff
            no = no + 1
        End If
    Next i
End If
If no = 0 Then
    MsgBox "Select a month from List !", vbOKOnly, "OASYS"
    txtMnt.Text = ""
    txtMnt.SetFocus
End If
End Sub
Private Sub ValidateYears()
'This function checks whether the Year entered exists in the Year combobox or not
Dim no As Integer
no = 0
If cmbYear.Text <> "" Then
    For i = 0 To cmbYear.ListCount - 1 Step 1
        If cmbYear.List(i) = cmbYear.Text Then
            cmbYear.ListIndex = i
            no = no + 1
            Call AddMonths
        End If
    Next i
End If
If no = 0 Then
    MsgBox "Select an Year from the List !", vbOKOnly, "OASYS"
    cmbYear.Text = ""
    cmbYear.SetFocus
End If
End Sub
Private Sub txtMnt_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
Private Sub TotalLeaveDays()
Dim TotalLeave, no As Integer
MNT = txtMnt.ListIndex + 1  ' By this we will get the No. of the Month. eg. if January=1, February=2, March=3.....
no = 0
TotalLeave = 0
If rsTmpLS.RecordCount > 0 Then rsTmpLS.MoveFirst
For i = 0 To rsTmpLS.RecordCount - 1 Step 1
    If rsTmpLS!Code = txtCode.Text And rsTmpLS!Year = cmbYear.Text And rsTmpLS!Month = MNT Then
        TotalLeave = TotalLeave + rsTmpLS!totDays
        no = no + 1
    End If
If rsTmpLS.EOF = False Then rsTmpLS.MoveNext
Next
If no = 0 Then
 'As no leave record of this employee is found, he must be present for all days of a month
    Call FullAttendance
    GoTo continue
End If
txtNOD.Text = TotalLeave   ' Total Leave Days
If TotalLeave > 2 Then
    txtExcess.Text = Val(TotalLeave) - 2
Else
    txtExcess.Text = 0
End If

txtDedPerDay.Text = Round(Val(txtSection.Text) / 30)  'This is to make it as a round figure by removing the value of decimal point
If txtNOD.Text <> "" And txtDedPerDay.Text <> "" Then
    txtTotAmtDed.Text = Round(Val(txtExcess.Text) * Val(txtDedPerDay.Text), 2)
End If

'Calculate the details in the Exchange Frame
txtTotLeavesAllowed.Text = 2
txtLeaveTaken.Text = TotalLeave  ' Total Leaves Taken
txtBalDays.Text = Val(txtTotLeavesAllowed.Text) - Val(TotalLeave)
txtAmtPerDay.Text = Round(Val(txtSection.Text) / 30)
If txtBalDays.Text <> "" And txtAmtPerDay.Text <> "" Then
   txtTotAmtAllowed.Text = Round(Val(txtBalDays.Text) * Val(txtAmtPerDay.Text), 2)
End If

continue:

If Val(txtBalDays.Text) < 0 Then  ' If more then 2 LeaveDays then show in Deduction
    FrameDed.Visible = True
    FrameExch.Visible = False
    optDeduction.Enabled = True
    optExchange.Enabled = False
    optDeduction.Value = True
ElseIf Val(txtBalDays.Text) > 0 Then ' If less then 2 LeaveDays then show in Exchange
    FrameDed.Visible = False
    FrameExch.Visible = True
    optDeduction.Enabled = False
    optExchange.Enabled = True
    optExchange.Value = True
ElseIf Val(txtBalDays.Text) = 0 Then
    FrameDed.Visible = True
    FrameExch.Visible = True
    optDeduction.Enabled = False
    optExchange.Enabled = False
End If

If cmdSave.Enabled = True Then cmdSave.SetFocus
End Sub

'If there is no Leave Record of an employee, then it implies that
'he must have full attendance in that particular month
Private Sub FullAttendance()
If rsWD.RecordCount > 0 Then rsWD.MoveFirst

For i = 0 To rsWD.RecordCount - 1 Step 1
   If cmbYear.Text = rsWD!Year And txtMnt.Text = rsWD!Month Then
    txtTotLeavesAllowed.Text = "2"
    txtLeaveTaken.Text = "0"
    txtBalDays.Text = "2"
    txtAmtPerDay.Text = Round(Val(txtSection.Text) / 30)
    If txtBalDays.Text <> "" And txtAmtPerDay.Text <> "" Then
        txtTotAmtAllowed.Text = Round(Val(txtBalDays.Text) * Val(txtAmtPerDay.Text), 2)
    End If
    Exit Sub
    End If
If rsWD.EOF = False Then rsWD.MoveNext
Next
If rsWD.EOF = True Then
    MsgBox "Working Days not found for corresponding Year and Month !", vbCritical, "OASYS"
End If
End Sub


Private Sub txtType_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtType_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
Private Sub DisableAll()  ' Disable all the text boxes
txtCode.Enabled = False
txtName.Enabled = False
txtDesig.Enabled = False
txtSection.Enabled = False
txtBasicPay.Enabled = False
optDeduction.Enabled = False
optExchange.Enabled = False
cmbYear.Enabled = False
txtMnt.Enabled = False
cmdHelp.Enabled = False

txtNOD.Enabled = False
txtDedPerDay.Enabled = False
txtTotAmtDed.Enabled = False

txtTotLeavesAllowed.Enabled = False

txtLeaveTaken.Enabled = False
txtBalDays.Enabled = False
txtAmtPerDay.Enabled = False
txtTotAmtAllowed.Enabled = False
End Sub
Private Sub EnableAll()  ' Enable all the text boxes
txtCode.Enabled = True
txtName.Enabled = True
txtDesig.Enabled = True
txtSection.Enabled = True
txtBasicPay.Enabled = True
optDeduction.Enabled = True
optExchange.Enabled = True
cmbYear.Enabled = True
txtMnt.Enabled = True
cmdHelp.Enabled = True

txtNOD.Enabled = True
txtDedPerDay.Enabled = True
txtTotAmtDed.Enabled = True

txtTotLeavesAllowed.Enabled = True
txtLeaveTaken.Enabled = True
txtBalDays.Enabled = True
txtAmtPerDay.Enabled = True
txtTotAmtAllowed.Enabled = True
End Sub
Private Sub ClearAll()  ' Clear all the text boxes
txtCode.Text = ""
txtName.Text = ""
txtDesig.Text = ""
txtSection.Text = ""
txtBasicPay.Text = ""
cmbYear.Text = ""
txtMnt.Text = ""

txtNOD.Text = ""
txtDedPerDay.Text = ""
txtTotAmtDed.Text = ""
txtExcess.Text = ""

txtTotLeavesAllowed.Text = ""
txtLeaveTaken.Text = ""
txtBalDays.Text = ""
txtAmtPerDay.Text = ""
txtTotAmtAllowed.Text = ""
End Sub

Private Sub txtNOD_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtNOD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtNOD_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtNoDays_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtNoDays_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtNoDays_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtSection_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtSection_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtSection_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtTotAmtAllowed_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtTotAmtAllowed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtTotAmtAllowed_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtTotAmtDed_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtTotAmtDed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtTotAmtDed_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub

Private Sub txtTotLeavesAllowed_GotFocus()
Call txt_GotFocus
End Sub

Private Sub txtTotLeavesAllowed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub

Private Sub txtTotLeavesAllowed_KeyUp(KeyCode As Integer, Shift As Integer)
Call txt_GotFocus
End Sub
