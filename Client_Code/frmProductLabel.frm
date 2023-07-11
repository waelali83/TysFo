VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{82CA1823-AE1A-4850-BDC7-F24C9A05E6D0}#1.1#0"; "UstriGrid.ocx"
Begin VB.Form frmProductLabel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Product Master - Add / Update"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   0
      Left            =   0
      TabIndex        =   46
      Top             =   8775
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   0
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15505
            Text            =   "<Esc - Quit>"
            TextSave        =   "<Esc - Quit>"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraProductLabel 
      Height          =   8595
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8545
      Begin VB.CommandButton cmdPastRecords 
         Caption         =   "Paste"
         Height          =   350
         Left            =   1200
         TabIndex        =   59
         Top             =   8040
         Width           =   1000
      End
      Begin VB.CommandButton cmdCopyRecords 
         Caption         =   "Copy"
         Height          =   350
         Left            =   120
         TabIndex        =   58
         Top             =   8040
         Width           =   1000
      End
      Begin VB.Frame fraGrid3 
         Height          =   1055
         Left            =   120
         TabIndex        =   52
         Top             =   3702
         Width           =   8295
         Begin VB.TextBox txtTareWeight 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   7095
            MaxLength       =   9
            TabIndex        =   10
            Top             =   615
            Width           =   855
         End
         Begin VB.TextBox txtMaxWeight 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   4395
            MaxLength       =   9
            TabIndex        =   9
            Top             =   615
            Width           =   855
         End
         Begin VB.TextBox txtMinWeight 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2040
            MaxLength       =   9
            TabIndex        =   8
            Top             =   615
            Width           =   855
         End
         Begin VB.TextBox txtBoxesPallet 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   7095
            MaxLength       =   3
            TabIndex        =   7
            Top             =   270
            Width           =   375
         End
         Begin VB.ComboBox cboWeightType 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "frmProductLabel.frx":0000
            Left            =   2040
            List            =   "frmProductLabel.frx":0002
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label lblBxsPlt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Boxes Per Pallet"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5595
            TabIndex        =   57
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label lblWgtType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Weight Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   56
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lblMaxWgt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Weight"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3240
            TabIndex        =   55
            Top             =   630
            Width           =   975
         End
         Begin VB.Label lblMinWgt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min. Weight"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   54
            Top             =   645
            Width           =   945
         End
         Begin VB.Label lblTareWgt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tare Weight"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5955
            TabIndex        =   53
            Top             =   645
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         Height          =   645
         Left            =   120
         TabIndex        =   49
         Top             =   7200
         Width           =   8295
         Begin VB.TextBox txtRefCode 
            Height          =   285
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   23
            Tag             =   "1"
            Top             =   240
            Width           =   3165
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cross Reference"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   50
            Top             =   300
            Width           =   1410
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   7440
         TabIndex        =   25
         Top             =   8040
         Width           =   1000
      End
      Begin VB.Frame fraProductMaster 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   8295
         Begin VB.TextBox txtProdCode 
            Height          =   285
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   1
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtDescription 
            Height          =   285
            Left            =   2040
            MaxLength       =   30
            TabIndex        =   2
            Top             =   581
            Width           =   2775
         End
         Begin VB.ComboBox cboProductType 
            Height          =   315
            ItemData        =   "frmProductLabel.frx":0004
            Left            =   2040
            List            =   "frmProductLabel.frx":0006
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   922
            Width           =   1350
         End
         Begin VB.TextBox txtLabelNo 
            Height          =   285
            Left            =   6975
            MaxLength       =   6
            TabIndex        =   4
            Top             =   975
            Width           =   930
         End
         Begin VB.Label lblDivisonCodeValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   7095
            TabIndex        =   26
            Top             =   630
            Width           =   165
         End
         Begin VB.Label lblDivisionCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Division Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5475
            TabIndex        =   45
            Top             =   630
            Width           =   1155
         End
         Begin VB.Label lblProdDesc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   43
            Top             =   630
            Width           =   1620
         End
         Begin VB.Label lblProdType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   42
            Top             =   1005
            Width           =   1065
         End
         Begin VB.Label lblLabelNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5475
            TabIndex        =   41
            Top             =   1005
            Width           =   795
         End
         Begin VB.Label lblProdCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   40
            Top             =   270
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         Height          =   660
         Left            =   120
         TabIndex        =   35
         Top             =   4778
         Width           =   8295
         Begin VB.TextBox txtStrPosition 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3675
            MaxLength       =   2
            TabIndex        =   12
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtWeightLength 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   7095
            MaxLength       =   2
            TabIndex        =   13
            Top             =   240
            Width           =   360
         End
         Begin VB.TextBox txtLabelLength 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   11
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblStrPosition 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start Pos. "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2760
            TabIndex        =   38
            Top             =   270
            Width           =   840
         End
         Begin VB.Label lblWeightLength 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Weight Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5715
            TabIndex        =   37
            Top             =   270
            Width           =   1185
         End
         Begin VB.Label lblLabelLength 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   36
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   120
         TabIndex        =   31
         Top             =   5459
         Width           =   8295
         Begin VB.TextBox txtPckdtStart 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   7755
            MaxLength       =   2
            TabIndex        =   17
            Top             =   240
            Width           =   360
         End
         Begin VB.ComboBox cboGovtLot 
            Height          =   315
            ItemData        =   "frmProductLabel.frx":0008
            Left            =   3660
            List            =   "frmProductLabel.frx":004B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   210
            Width           =   630
         End
         Begin VB.ComboBox cboMfgDate 
            Height          =   315
            ItemData        =   "frmProductLabel.frx":008E
            Left            =   5535
            List            =   "frmProductLabel.frx":00BF
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   210
            Width           =   630
         End
         Begin VB.TextBox txtCommodityCode 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   14
            Top             =   210
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mfg. Date Pos. "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   6480
            TabIndex        =   60
            Top             =   270
            Width           =   1230
         End
         Begin VB.Label lblMfgDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mfg. Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4695
            TabIndex        =   34
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblGovtLot 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Govt. Lot"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2760
            TabIndex        =   33
            Top             =   240
            Width           =   705
         End
         Begin VB.Label lblCommodityCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Commodity Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   32
            Top             =   240
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1050
         Left            =   120
         TabIndex        =   27
         Top             =   6125
         Width           =   8295
         Begin VB.ComboBox cboPrdgrpcode 
            Height          =   315
            ItemData        =   "frmProductLabel.frx":00F0
            Left            =   4680
            List            =   "frmProductLabel.frx":00F2
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cboBoxSerialInd 
            Height          =   315
            ItemData        =   "frmProductLabel.frx":00F4
            Left            =   2040
            List            =   "frmProductLabel.frx":0113
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   600
            Width           =   630
         End
         Begin VB.ComboBox cboLblCodeChk 
            Height          =   315
            ItemData        =   "frmProductLabel.frx":0132
            Left            =   2040
            List            =   "frmProductLabel.frx":013F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   240
            Width           =   630
         End
         Begin VB.TextBox txtLabelCodeLength 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   7095
            MaxLength       =   2
            TabIndex        =   20
            Top             =   240
            Width           =   330
         End
         Begin VB.TextBox txtLabelCodeStr 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   4410
            MaxLength       =   2
            TabIndex        =   19
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product Group code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2880
            TabIndex        =   48
            Top             =   690
            Width           =   1635
         End
         Begin VB.Label lblBoxSerialInd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Box Serial Ind."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   270
            TabIndex        =   44
            Top             =   690
            Width           =   1170
         End
         Begin VB.Label lblLabelCodeLength 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label Code Length"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5385
            TabIndex        =   30
            Top             =   285
            Width           =   1575
         End
         Begin VB.Label lblLabelCodeString 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label Code Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2850
            TabIndex        =   29
            Top             =   285
            Width           =   1380
         End
         Begin VB.Label lblLabelCodeCheck 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label Code Check"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   255
            TabIndex        =   28
            Top             =   285
            Width           =   1530
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   350
         Left            =   6360
         TabIndex        =   24
         Top             =   8040
         Width           =   1000
      End
      Begin VB.Frame Frame5 
         Height          =   2175
         Left            =   120
         TabIndex        =   51
         Top             =   1506
         Width           =   8295
         Begin USTriSuperGrid.USTriGrid ustgrdSecondaryCustomers 
            Height          =   1925
            Left            =   75
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   165
            Width           =   8130
            _ExtentX        =   14340
            _ExtentY        =   3387
            Columns         =   2
            FixedColumns    =   0
            FixedRows       =   0
            Rows            =   2
            TopRow          =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BackColorBkg    =   -2147483633
            BackColorFixed  =   -2147483633
            BackColorSel    =   -2147483635
            BorderStyle     =   1
            ForeColor       =   -2147483640
            ForeColorFixed  =   -2147483630
            ForeColorSel    =   -2147483634
            GridColor       =   -2147483630
            GridColorFixed  =   12632256
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            MergeCells      =   0
            RowSel          =   0
            ColSel          =   0
            TextStyle       =   0
            TextStyleFixed  =   0
            WordWrap        =   -1  'True
            AllowUserResizing=   0
            ScrollBars      =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BandDisplay     =   0
            BackColorUnpopulated=   -2147483633
            AllowBigSelection=   -1  'True
            Object.WhatsThisHelpID =   0
            version         =   "1.0"
         End
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label Code Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   1380
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuError 
         Caption         =   "Show Errors"
      End
   End
End
Attribute VB_Name = "frmProductLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : frmProductLabel.frm
'*File Description              : Add/Update/Delete a Detail Inventory record
'*Author                        : US Technology
'*Date Created                  : Oct-05-02
'*Date Last Modified            : Mar-17-03
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : MasterFileFunctions
'*Components Used               : None
'*Functions Defined             : DisplayRecords, LoadTypesCombos,
'*                                RecordIsBlank, SelectComboByType,
'*                                UpdateDetailInventory,
'*                                ValidateControlTexts
'*                                ValidateMaxWgt,ValidateMinWgt
'*                                ValidateNumericFieldSize,
'*                                ValidateRecordForBlank
'*                                ValidateTareWgt,SetDivisionCode
'*                                DisplayProductDetails,
'*                                SetMasterFieldsForEdit,
'*                                SetInitialValues
'*Copyright                     : US Technology
'-------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description    Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Public fcolProduct                 As Collection
Public fcolProducts                As Collection
Public fcolLabelCodes              As Collection
Private fcolControls               As Collection

Public frsProdTypes                As ADODB.Recordset
Public frsWgtTypes                 As ADODB.Recordset
Public frsPrdGrpCode               As ADODB.Recordset

Private fsControlText               As String
Private fbFormIsLoading             As Boolean
Private fbUpdateDB                  As Boolean

Private fsCopyString                As String
Private fbCanBeCopied               As Boolean
Private fbLoadSuccess               As Boolean


Public Event RefreshGrid(ByVal sMode As String, ByVal colProduct As Collection)
'Public Event XRef(ByVal sXRefValue As String)

'Added by TCS

Private Enum ColumnName

    cnProduct = 0
    cnOriginPlant = 1
    cnCustNumber = 2
    cnCustType = 3
    cnFreezerDays = 4
    cnAction = 5
    cnBlastDays = 6
    cnPrevOriginPlant = 7
    cnPrevCustNumber = 8
    cnPrevCustType = 9
    'Added by TCS for Req 22 ver 1
    cnstatus = 10
    cnDefaultOrigin = 11
    
End Enum

'ProductCode whose secondary customers are to be displayed

Public fsProduct                As String
Public fsProductType            As String

' Used for getting the details of the error that has come

Private fcolErrors              As Collection

Private flRowIndex              As Long
Private flColumnIndex           As Long
Private fbUnload                As Boolean
Private fbCancelled             As Boolean
Private fsArrCustType()         As String
Private fsArrDefOrigin()        As String
Private fdictOrgPlant           As Dictionary
Private lgrdRowIndex            As Long

' Added by TCS
Dim sBlastInd As String

Private Const COLUMN_COUNT      As Integer = 12  ' 10 - Changed by TCS for Req 22 ver 1
'For Req 22 ver 1
'Added by TCS for OrderDetails Grid Row Selection
'Golbal Variable for gird selection
Dim PrevRowSelected As Integer
Dim bAllowSelChangeEvent As Boolean
'Added by Tcs on 17/Jan/05
Public fAddEditMode            As String   ' New_Record , Update_Record



'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cboBoxSerialInd_GotFocus()

    fsControlText = cboBoxSerialInd.Text

End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cboLblCodeChk_GotFocus()

    fsControlText = cboLblCodeChk.Text

End Sub



Private Sub cboMfgDate_Click()
txtPckdtStart.Locked = cboMfgDate.Text = "N"
If txtPckdtStart.Locked Then txtPckdtStart.Text = 0
End Sub

'******************************************************************************
'* Functional Description   :   Validation
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cboProductType_Validate(Cancel As Boolean)

Dim lRowToExclude As Long
Dim lResult As Long
    
    If fbFormIsLoading Then Exit Sub
    
    If fcolProduct Is Nothing Then
        
        lRowToExclude = -1
    
    Else
        
        'lRowToExclude = fcolProduct.Item("ROW_INDEX")
    
    End If
    
    If Not RecordIsBlank Then
    
        Dim strLabelNo As String
        Dim intLabelNo As Long
    
        intLabelNo = CLng(txtLabelNo.Text)
        strLabelNo = CStr(intLabelNo)
        
        If ItemExistsInCollection(fcolLabelCodes, _
            strLabelNo & "|" & Left(cboProductType.Text, 1), _
            strLabelNo & "|" & Left(cboProductType.Text, 1)) Then
        
            lResult = 0
            Dim rsProductCodes      As ADODB.Recordset
            Dim strProducts         As String
            Dim objProduct          As Object
         
            'Create object of ProductMaster Class in MasterFileFunctions
    
            Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
    
            'Retrieve duplicate Product Codes
        
            lResult = objProduct.GetDuplicateProductCodes(strLabelNo, _
                                                      Left(cboProductType.Text, 1), _
                                                      rsProductCodes)
            ''condition Added by  on 21-Jun-2005
            If Not rsProductCodes Is Nothing Then
                rsProductCodes.MoveFirst
                strProducts = rsProductCodes.Fields("PRODUCT_CODE").Value
                rsProductCodes.MoveNext
                While Not rsProductCodes.EOF
                    strProducts = strProducts + " - " + rsProductCodes.Fields("PRODUCT_CODE").Value
                    rsProductCodes.MoveNext
                Wend
                
                gcolErrMsg.Add strProducts
                giErrMsg = ShowErrorMsg("ProLbl032", Me.Caption, vbOKOnly, gcolErrMsg)
            End If
            GoTo CleanUpAndExit
            
        End If
    
    End If
    
    SetDivisionCode Left(Trim(cboProductType.Text), 1)
    Exit Sub

CleanUpAndExit:

    Cancel = True
    SelectComboByType cboProductType, Left(fsControlText, 1)
    cboProductType.SetFocus
    
End Sub

'******************************************************************************
'* Functional Description   :   Validation
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cboWeightType_Validate(Cancel As Boolean)

    If fbFormIsLoading Then Exit Sub

    If Left(cboWeightType.Text, 1) = "F" Then
        
        txtMaxWeight.Text = txtMinWeight.Text
        txtMaxWeight.Enabled = False
    
    Else
        
        txtMaxWeight.Enabled = True
    
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Unload form without updating DB
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cmdCancel_Click()
    
    fbUpdateDB = False
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Update DB and unload form
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cmdOK_Click()

    DoEvents
    fbUpdateDB = True
    Unload Me
    
End Sub

'Added by TCS on 22-Aug-05
'Start
Private Sub Form_Activate()
    cmdCopyRecords.Visible = False
    cmdPastRecords.Visible = False
End Sub
'End

'******************************************************************************
'* Functional Description   :   Prepares screen for display
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub Form_Load()
    
Dim varControl As Variant

    'Initializing form level variables
    
    fbUpdateDB = False
    fbFormIsLoading = True
    fbLoadSuccess = True
    
    LoadTypesCombos
    SetCombos
    cboWeightType.Enabled = True
    SetInitialValues
    Set fcolControls = New Collection
    
    For Each varControl In Me.Controls
        
        If Left(varControl.Name, 3) = "cbo" Or _
                Left(varControl.Name, 3) = "txt" Then
            
            fcolControls.Add varControl, CStr(varControl.TabIndex)
        
        End If
    
    Next

    SizeGridColumns
    FormatGrid
    
    
    
    If fAddEditMode = "Update_Record" Then
    
        If Not fcolProduct Is Nothing Then
        
            txtProdCode.Text = fcolProduct.Item("PRODUCT_CODE")
            fsProduct = fcolProduct.Item("PRODUCT_CODE")
            
            'Load Product Details
            
            LoadProductDetails
            lgrdRowIndex = 1
            If Not DisplayRecord Then
                Unload Me
                Exit Sub
            End If
                    
            ' Commented by George on 06-Dec-2005 for HD0000000534698
            'If gsUserGroup = 102 Then
                cboWeightType.Enabled = True
            'Commented by TCS on 23-Aug-2005
            'Else
            '    cboWeightType.Enabled = False
            'End If
            
            LoadSecondaryCustomers
            
            With ustgrdSecondaryCustomers
            
            For lgrdRowIndex = 1 To .Rows - 1
                .TextMatrix(lgrdRowIndex, cnAction) = TO_UPDATE
            Next
                     lgrdRowIndex = 1
            End With
            
            ustgrdSecondaryCustomers.Row = 1
            ustgrdSecondaryCustomers_Click

        Else
            
        End If
    Else
    
        ReDim fsArrDefOrigin(2)
        fsArrDefOrigin(0) = "Y"
        fsArrDefOrigin(1) = "N"
        ustgrdSecondaryCustomers.AddComboData cnDefaultOrigin, fsArrDefOrigin
        
       'Added by TCS-Ragu on 25-Aug-2005
        ustgrdSecondaryCustomers.AddComboData cnstatus, fnGetPRODStatus
    
        Set fcolProduct = New Collection
        fcolProduct.Add "New", "PRODUCT_CODE_KEY"
        fcolProduct.Add 1, "ROW_INDEX"
        fcolProduct.Add 1, "RowNum"
        lgrdRowIndex = 1
        NewRecord
     
    End If

    
    fbFormIsLoading = False
    
    'Added by TCs on 23-Aug-05
    'Start
    'Commented by George on 06-Dec-2005 for HD0000000534698
    cboWeightType.Enabled = True
    'If gsUserGroup <> "102" Then
    '    cboWeightType.Enabled = IIf(ustgrdSecondaryCustomers.CellValue(ustgrdSecondaryCustomers.Row, cnAction) = "I", True, _
    '                            IIf(ustgrdSecondaryCustomers.CellValue(ustgrdSecondaryCustomers.Row, cnAction) = "", True, False))
    'End If
    'End

    
End Sub

'******************************************************************************
'* Functional Description   :   Populate all controls in Add screen.
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub SetInitialValues()
    
    txtBoxesPallet.Text = 0
    txtLabelNo.Text = 0
    txtMinWeight.Text = "0.00"
    txtMaxWeight.Text = "0.00"
    txtTareWeight.Text = "0.00"
    'txtFreezerDays.Text = 0
    'txtBlastDays.Text = 0
    txtLabelLength.Text = 0
    txtStrPosition.Text = 0
    txtWeightLength.Text = 0
    txtLabelCodeStr.Text = 0
    txtLabelCodeLength.Text = 0
    cboGovtLot.ListIndex = 0
    cboProductType.ListIndex = 0
    'cboWeightType.ListIndex = 0 'Commented by TCS on 23-Aug-2005
    cboPrdgrpcode.ListIndex = 0
    SetDivisionCode Left(Trim(cboProductType.Text), 1)
    txtLabelLength.Text = 46
    cboMfgDate.Text = "Y"
    txtPckdtStart.Text = "29"
    
End Sub

'******************************************************************************
'* Functional Description   :   Populate combos
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub SetCombos()

    cboGovtLot.Clear
    cboGovtLot.AddItem ""
    cboGovtLot.AddItem "Y"
    cboGovtLot.AddItem "N"
    
    cboMfgDate.Clear
    cboMfgDate.AddItem "Y"
    cboMfgDate.AddItem "N"
    
    cboLblCodeChk.Clear
    cboLblCodeChk.AddItem "Y"
    cboLblCodeChk.AddItem "N"
    
    cboBoxSerialInd.Clear
    cboBoxSerialInd.AddItem ""
    cboBoxSerialInd.AddItem "Y"
    cboBoxSerialInd.AddItem "N"

End Sub

'******************************************************************************
'* Functional Description   :   Loads types combos - Product and Weight Types
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub LoadTypesCombos()

    'Load Product Types
    
    frsProdTypes.MoveFirst
    
    While Not frsProdTypes.EOF
        
        cboProductType.AddItem _
                        frsProdTypes.Fields("TYPE_CODE").Value & " - " & _
                        frsProdTypes.Fields("TYPE_SHORT_DESC").Value
        frsProdTypes.MoveNext
    
    Wend
        
    'Load Weight Types
    
    frsWgtTypes.MoveFirst
    
    While Not frsWgtTypes.EOF
        
        cboWeightType.AddItem _
                        frsWgtTypes.Fields("TYPE_CODE").Value & " - " & _
                        frsWgtTypes.Fields("TYPE_SHORT_DESC").Value
        
        frsWgtTypes.MoveNext
    
    Wend
    
    'Load Product group code Added by TCS
    
    If frsPrdGrpCode.RecordCount <> 0 Then
    
      frsPrdGrpCode.MoveFirst
    
      While Not frsPrdGrpCode.EOF
        
        cboPrdgrpcode.AddItem Trim(frsPrdGrpCode.Fields("TYPE_SHORT_DESC").Value)
        
        frsPrdGrpCode.MoveNext
    
      Wend
      
  End If
  
      cboPrdgrpcode.AddItem "", 0
    
End Sub

'******************************************************************************
'* Functional Description   :   Loads details to screen
'* Parameter Description    :   None
'* Return Type Description  :   Success of loading
'******************************************************************************

Private Function DisplayRecord() As Boolean

    DisplayRecord = True
    

  '  If fAddEditMode = "" Then Exit Function
    On Error Resume Next ' Test case
    
    With fcolProduct
        
        txtProdCode.Text = .Item("PRODUCT_CODE" & "_" & Trim(Str(lgrdRowIndex)))
        txtProdCode.Enabled = False
        txtDescription.Text = .Item("PRODUCT_DESC" & "_" & Trim(Str(lgrdRowIndex)))
        lblDivisonCodeValue.Caption = .Item("DIVISION_CODE" & "_" & Trim(Str(lgrdRowIndex)))
        SelectComboByType cboProductType, Left(.Item("PRODUCT_TYPE" & "_" & Trim(Str(lgrdRowIndex))), 1)
        txtLabelNo.Text = .Item("LABEL_NO" & "_" & Trim(Str(lgrdRowIndex)))
        SelectComboByType cboWeightType, Left(.Item("WGT_TYPE_CODE" & "_" & Trim(Str(lgrdRowIndex))), 1)
        
        If .Item("WGT_TYPE_CODE" & "_" & Trim(Str(lgrdRowIndex))) = "F" Then
            
            txtMaxWeight.Enabled = False
        
        Else
            
            txtMaxWeight.Enabled = True
        
        End If
        
        txtBoxesPallet.Text = .Item("BOXES_PER_PALLET" & "_" & Trim(Str(lgrdRowIndex)))
        txtMinWeight.Text = .Item("MIN_BOX_WGT" & "_" & Trim(Str(lgrdRowIndex)))
        txtMaxWeight.Text = .Item("MAX_BOX_WGT" & "_" & Trim(Str(lgrdRowIndex)))
        txtTareWeight.Text = .Item("BOX_TARE_WEIGHT" & "_" & Trim(Str(lgrdRowIndex)))
              
        txtLabelLength.Text = .Item("LABEL_LENGTH" & "_" & Trim(Str(lgrdRowIndex)))
        txtStrPosition.Text = .Item("LABEL_WGT_ST_POS" & "_" & Trim(Str(lgrdRowIndex)))
        txtWeightLength.Text = .Item("LABEL_WGT_LENGTH" & "_" & Trim(Str(lgrdRowIndex)))
        
              
        'txtCustomerNumber_Validate False
        
'        If Not fbLoadSuccess Then
'            DisplayRecord = False
'            Exit Function
'        End If
        
        'cboCustomerType .Item("CUSTOMER_TYPE")
        
        If fbFormIsLoading Then 'Condition added by TCS-Ragu on 05-jan-06 for Tk#547007
            txtCommodityCode.Text = .Item("GOVT_COMMODITY_CODE" & "_" & Trim(Str(lgrdRowIndex)))
        End If
        
        cboGovtLot.ListIndex = 2
        If Trim(.Item("GOVT_LOT_IND" & "_" & Trim(Str(lgrdRowIndex)))) = "" Then
            
            cboGovtLot.ListIndex = 0
        
        Else
            
            SelectComboByType cboGovtLot, .Item("GOVT_LOT_IND" & "_" & Trim(Str(lgrdRowIndex)))
        
        End If
        
        SelectComboByType cboMfgDate, Trim(.Item("PACK_DATE_IND" & "_" & Trim(Str(lgrdRowIndex))))
        
        'Added by TCS-Ragu on 17-Jun-2005
        'Start
        txtPckdtStart.Text = Trim(.Item("PACK_DATE_START" & "_" & Trim(Str(lgrdRowIndex))))
        'End
        
        SelectComboByType cboLblCodeChk, Trim(.Item("CHECK_LABEL_IND" & "_" & Trim(Str(lgrdRowIndex))))
        txtLabelCodeStr.Text = Trim(.Item("PROD_LABEL_START" & "_" & Trim(Str(lgrdRowIndex))))
        txtLabelCodeLength.Text = Trim(.Item("PROD_LABEL_LEN" & "_" & Trim(Str(lgrdRowIndex))))
        
        If Trim(.Item("BOX_SERIAL_IND" & "_" & Trim(Str(lgrdRowIndex)))) = "" Then
            
            cboBoxSerialInd.ListIndex = 0
        
        Else
            
            SelectComboByType cboBoxSerialInd, Trim(.Item("BOX_SERIAL_IND" & "_" & Trim(Str(lgrdRowIndex))))
        
        End If
        
     ' Added by TCS
        
     
        If IsNull(.Item("XREF_CODE" & "_" & Trim(Str(lgrdRowIndex)))) Then
                txtRefCode.Text = ""
        Else
                txtRefCode.Text = .Item("XREF_CODE" & "_" & Trim(Str(lgrdRowIndex)))
        End If
    
        
        'Added by TCS
        If fbFormIsLoading Then 'Condition added by TCS-Ragu on 05-jan-06 for Tk#547007
            If Trim(.Item("PRODUCT_GROUP_CODE" & "_" & Trim(Str(lgrdRowIndex)))) = "" Then
                cboPrdgrpcode.ListIndex = 0
            Else
                cboPrdgrpcode.Text = .Item("PRODUCT_GROUP_CODE" & "_" & Trim(Str(lgrdRowIndex)))
            End If
        End If
        
'        If Trim(.Item("PRODUCT_GROUP_CODE" & "_" & Trim(Str(lgrdRowIndex)))) = "" Then
'
'            cboPrdgrpcode.ListIndex = 0
'
'        Else
'
'            SelectComboByType cboPrdgrpcode, Trim(.Item("PRODUCT_GROUP_CODE" & "_" & Trim(Str(lgrdRowIndex))))
'
'        End If
        
        SetMasterFieldsForEdit True
        
    End With
    
End Function

'******************************************************************************
'* Functional Description   :   Loads master details to screen
'* Parameter Description    :   rsProduct - Details from Master table
'* Return Type Description  :   None
'******************************************************************************

Private Sub DisplayProductDetails(ByVal rsProduct As ADODB.Recordset)

Dim bClearTexts    As Boolean
Dim fld As Field

    If rsProduct Is Nothing Then
        
        bClearTexts = True
    
    ElseIf rsProduct.EOF Then
        
        bClearTexts = True
    
    End If
    
    With rsProduct
        
        If Not bClearTexts Then
            
            txtProdCode.Text = .Fields("PRODUCT_CODE")
            txtDescription.Text = .Fields("PRODUCT_DESC")
            lblDivisonCodeValue.Caption = .Fields("DIVISION_CODE")
            SelectComboByType cboProductType, .Fields("PRODUCT_TYPE")
            txtLabelNo.Text = .Fields("LABEL_NO")
            SelectComboByType cboWeightType, .Fields("WGT_TYPE_CODE")
            If IsNull(.Fields("GOVT_COMMODITY_CODE")) Then
                txtCommodityCode.Text = ""
            Else
                txtCommodityCode.Text = .Fields("GOVT_COMMODITY_CODE")
            End If
            SetMasterFieldsForEdit False
            
    'Added by TCS
             
            For Each fld In .Fields
                If "XREF_CODE" = fld.Name Then
                   If IsNull(.Fields("XREF_CODE")) Then
                      txtRefCode.Text = ""
                   Else
                      txtRefCode.Text = .Fields("XREF_CODE")
                   End If
               End If
          Next
          
        Else
            
            txtDescription.Text = ""
            lblDivisonCodeValue.Caption = ""
            cboProductType.ListIndex = -1
            txtLabelNo.Text = ""
            cboWeightType.ListIndex = -1
            txtCommodityCode.Text = ""
            'Added by TCS
            cboPrdgrpcode.ListIndex = -1
            txtRefCode.Text = ""
            
            SetMasterFieldsForEdit True
        
        End If
    
    End With
    
End Sub

'******************************************************************************
'* Functional Description   :   Enable or disable the master table details for
'*                              the product
'* Parameter Description    :   bEnable- Status
'* Return Type Description  :   None
'******************************************************************************

Private Sub SetMasterFieldsForEdit(ByVal bEnable As Boolean)
        
    txtDescription.Enabled = bEnable
    cboProductType.Enabled = bEnable
    txtLabelNo.Enabled = bEnable
'    txtCommodityCode.Enabled = bEnable
'    txtRefCode.Enabled = bEnable    ' Added by TCS
'    cboPrdgrpcode.Enabled = bEnable ' Added by TCS

End Sub

'******************************************************************************
'* Functional Description   :   Update the database
'* Parameter Description    :   None
'* Return Type Description  :   Boolean - True if Update is success
'******************************************************************************

'Private Function UpdateDetailInventory() As Boolean
'
'Dim objProduct          As Object
'Dim colProduct          As Collection
'Dim colProductClone     As Collection
'Dim colProducts         As Collection
'Dim lResult             As Long
'Dim sAction             As String
'Dim sProdKey            As String
'Dim sStatusText         As String
'Dim bDuplLblCode        As Boolean
'
'    On Error GoTo ErrHandler
'
'    UpdateDetailInventory = True
'    If ValidateRecordForBlank Then
'
'        'Add to collection and call update for Add and Update modes
'
'        If fcolProduct Is Nothing Then
'
'            sAction = TO_INSERT
'            sProdKey = Trim(txtProdCode.Text)
'
'        ElseIf Not fcolProduct Is Nothing Then
'
'            sAction = TO_UPDATE
'            sProdKey = fcolProduct.Item("PRODUCT_CODE_KEY")
'
'        End If
'
'        Set colProduct = New Collection
'        colProduct.Add sAction, "ACTION"
'        colProduct.Add 1, "RECORDCOUNT"
'        colProduct.Add Trim(txtProdCode.Text), "PRODUCT_CODE"
'
'        colProduct.Add Trim(txtDescription.Text), "PRODUCT_DESC"
'        colProduct.Add Left(cboProductType.Text, 1), "PRODUCT_TYPE"
'        colProduct.Add Trim(lblDivisonCodeValue.Caption), "DIVISION_CODE"
'        colProduct.Add Trim(txtLabelNo.Text), "LABEL_NO"
'        colProduct.Add Left(cboWeightType.Text, 1), "WGT_TYPE_CODE"
'        colProduct.Add Trim(txtCommodityCode.Text), "GOVT_COMMODITY_CODE"
'
'        colProduct.Add Trim(txtBoxesPallet.Text), "BOXES_PER_PALLET"
'        colProduct.Add Val(Trim(txtMinWeight.Text)), "MIN_BOX_WGT"
'        colProduct.Add Val(Trim(txtMaxWeight.Text)), "MAX_BOX_WGT"
'        colProduct.Add Val(Trim(txtTareWeight.Text)), "BOX_TARE_WEIGHT"
'
'        colProduct.Add "0", "BLAST_IND"  ' Changed by TCS
'        colProduct.Add 0, "FREEZE_DAYS"
'        colProduct.Add 0, "BLAST_DAYS"
'
'        colProduct.Add Val(Trim(txtLabelLength.Text)), "LABEL_LENGTH"
'        colProduct.Add Val(Trim(txtStrPosition.Text)), "LABEL_WGT_ST_POS"
'        colProduct.Add Val(Trim(txtWeightLength.Text)), "LABEL_WGT_LENGTH"
'
'        colProduct.Add gsPlantCode, "CUSTOMER_ID" 'Changed by TCS
'        colProduct.Add "S", "CUSTOMER_TYPE" 'Changed by TCS
'
'        colProduct.Add cboGovtLot.Text, "GOVT_LOT_IND"
'        colProduct.Add cboMfgDate.Text, "PACK_DATE_IND"
'
'        colProduct.Add cboLblCodeChk.Text, "CHECK_LABEL_IND"
'        colProduct.Add Val(Trim(txtLabelCodeStr.Text)), "PROD_LABEL_START"
'        colProduct.Add Val(Trim(txtLabelCodeLength.Text)), "PROD_LABEL_LEN"
'        colProduct.Add cboBoxSerialInd.Text, "BOX_SERIAL_IND"
'
'        colProduct.Add Trim(txtRefCode.Text), "XREF_CODE"  ' Added by TCS
'
'        colProduct.Add "Y", "SECOND_IND"  ' Added by TCS
'
'         'Added by TCS
'        If cboPrdgrpcode.Text = "" Then
'           colProduct.Add Null, "PRODUCT_GROUP_CODE"
'        Else
'           colProduct.Add Trim(cboPrdgrpcode.Text), "PRODUCT_GROUP_CODE"
'        End If
'        colProduct.Add sProdKey, "PRODUCT_CODE_KEY"
'
'
'        Set colProductClone = New Collection
'        Set colProductClone = colProduct
'
'        Set colProducts = New Collection
'        colProducts.Add colProduct, sProdKey
'
'        'Call the update method of ProductMaster to update the db
'
'        If Not colProducts Is Nothing Then
'
'            'Create object of ProductMaster in MasterFileFunctions
'
'            Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
'
'            sStatusText = sbStatusBar.Panels(1).Text
'            sbStatusBar.Panels(1).Text = _
'                                      "Updating Database - Please Wait..."
'            Me.MousePointer = vbHourglass
'            DoEvents
'
'            'Update Product Details
'
'            lResult = objProduct.UpdateProductForPlant(gsPlantCode, _
'                                                    colProducts, bDuplLblCode)
'
'            sbStatusBar.Panels(1).Text = sStatusText
'            Me.MousePointer = vbDefault
'
'            If lResult <> 0 Then
'
'                'If error in updating, show error window
'
'                gcolErrMsg.Add lResult
'                giErrMsg = ShowErrorMsg("ProLbl002", _
'                                        Me.Caption, _
'                                        vbOKOnly, _
'                                        gcolErrMsg)
'
'                UpdateDetailInventory = False
'                GoTo CleanUpAndExit
'
'            ElseIf bDuplLblCode Then
'
'                'Show error message in case of duplicate label code / prod
'                'type
'
'                giErrMsg = ShowErrorMsg("ProLbl014", Me.Caption, vbOKOnly)
'
'                UpdateDetailInventory = False
'                txtLabelNo.SelStart = 0
'                txtLabelNo.SelLength = Len(Trim(txtLabelNo.Text))
'                txtLabelNo.SetFocus
'                GoTo CleanUpAndExit
'
'            End If
'
'            If fcolProduct Is Nothing Then
'
'                RaiseEvent RefreshGrid("I", colProductClone)
'
'            Else
'
'                RaiseEvent RefreshGrid("U", colProductClone)
'
'            End If
'
'        End If
'
'    Else
'
'        UpdateDetailInventory = False
'
'    End If
'
'CleanUpAndExit:
'
'    Set colProduct = Nothing
'    Set colProductClone = Nothing
'    Set colProducts = Nothing
'    Set objProduct = Nothing
'    Exit Function
'
'ErrHandler:
'
'    'VB Error
'
'    UpdateDetailInventory = False
'
'    gcolErrMsg.Add Err.Description
'    giErrMsg = ShowErrorMsg("ProLbl004", Me.Caption, vbOKOnly, gcolErrMsg)
'
'    UpdateDetailInventory = False
'    GoTo CleanUpAndExit
'
'End Function

'******************************************************************************
'* Functional Description   :   Select the combo item
'* Parameter Description    :   cboCombo - Combo in which text is to be set
'*                              sText - Text to be selected in the combo
'* Return Type Description  :   None
'******************************************************************************

Private Sub SelectComboByType(ByRef cboCombo As ComboBox, _
                            ByVal sText As String)

Dim lIndex As Long

    For lIndex = 0 To cboCombo.ListCount - 1
        
        If Left(cboCombo.List(lIndex), 1) = sText Then
            
            cboCombo.ListIndex = lIndex
            Exit Sub
        
        End If
    
    Next
    
End Sub

'******************************************************************************
'* Functional Description   :   Check if the not null fields are left blank
'* Parameter Description    :   None
'* Return Type Description  :   Boolean - False if any not null field is blank
'******************************************************************************

Private Function ValidateRecordForBlank() As Boolean

Dim varControl      As Variant
Dim lCount          As Long
Dim bCancel         As Boolean


    ValidateRecordForBlank = True
    
    For lCount = 1 To fcolControls.Count - 1
       
        If Not (lCount = 5 Or lCount = 6 Or lCount = 7) Then
        If Trim(fcolControls(CStr(lCount)).Text) = "" Or _
            (lCount = 6 And Val(fcolControls(CStr(lCount)).Text) = 0) Or _
            (lCount = 7 And Val(fcolControls(CStr(lCount)).Text) = 0) Then
            
            
            
            
            ' If screen is in Update mode, it will alert as invalid entry
            ' else the method returns false
            
            If lCount >= 11 Then Exit For
                
                Select Case lCount

                    Case 1
                        giErrMsg = ShowErrorMsg("ProLbl017", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 2
                        giErrMsg = ShowErrorMsg("ProLbl018", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 3
                        giErrMsg = ShowErrorMsg("ProLbl019", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 4
                        giErrMsg = ShowErrorMsg("ProLbl020", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 5
                        giErrMsg = ShowErrorMsg("ProLbl021", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 6
                        giErrMsg = ShowErrorMsg("ProLbl022", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 7
                        giErrMsg = ShowErrorMsg("ProLbl023", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 8
                        giErrMsg = ShowErrorMsg("ProLbl024", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 9
                        giErrMsg = ShowErrorMsg("ProLbl025", _
                                                Me.Caption, _
                                                vbOKOnly)
                    Case 10
                        giErrMsg = ShowErrorMsg("ProLbl026", _
                                                Me.Caption, _
                                                vbOKOnly)
                End Select

                If fcolControls(CStr(lCount)).Enabled Then
                    
                    fcolControls(CStr(lCount)).SetFocus
                
                End If
                
                ValidateRecordForBlank = False
                
            Exit Function
        
        End If
        End If
    Next
    
    bCancel = False
    
    If Trim(cboMfgDate.Text) = "" Then
        
        giErrMsg = ShowErrorMsg("ProLbl027", _
                                Me.Caption, _
                                vbOKOnly)
        ValidateRecordForBlank = False
        cboMfgDate.SetFocus
        Exit Function
    
    End If
    
    If Trim(cboLblCodeChk.Text) = "" Then
        
        giErrMsg = ShowErrorMsg("ProLbl028", _
                                Me.Caption, _
                                vbOKOnly)
        ValidateRecordForBlank = False
        cboLblCodeChk.SetFocus
        Exit Function
    
    End If
    
    ValidateRecordForBlank = False
    
    Call txtMinWeight_Validate(bCancel)
    If bCancel Then Exit Function
    
    Call txtMaxWeight_Validate(bCancel)
    If bCancel Then Exit Function
    
    Call txtTareWeight_Validate(bCancel)
    If bCancel Then Exit Function
    
    'Call txtCustomerNumber_Validate(bCancel)
    'If bCancel Then Exit Function
    
    Call txtCommodityCode_Validate(bCancel)
    If bCancel Then Exit Function
    
    ValidateRecordForBlank = True
    
End Function

'******************************************************************************
'* Functional Description   :   Validates for Tare_Weight > 0
'* Parameter Description    :   None
'* Return Type Description  :   Boolean - True if Tare_Weight > 0
'******************************************************************************

Private Function ValidateTareWgt() As Boolean
        
    With txtTareWeight
        
        If Trim(.Text) = "" Then GoTo CleanUpAndExit
        
        If Val(.Text) <= 0 Then
            
            giErrMsg = ShowErrorMsg("ProLbl006", Me.Caption, vbOKOnly)
            ValidateTareWgt = False
            Exit Function
        
        End If
    
    End With
    
CleanUpAndExit:

    ValidateTareWgt = True
    
End Function

'******************************************************************************
'* Functional Description   :   Sets the value of Division Code
'* Parameter Description    :   sProdType - Current Product Type
'* Return Type Description  :   None
'******************************************************************************

Private Sub SetDivisionCode(ByVal sProdType As String)
        
    With lblDivisonCodeValue
        
        Select Case sProdType
            
            Case "B": .Caption = "11"
            Case "D": .Caption = "22"
            Case "I": .Caption = "11"
            Case "O": .Caption = "99"
            Case "P": .Caption = "31"
            Case "V": .Caption = "05"
            Case "C": .Caption = "98"
            Case "R": .Caption = "57"
            Case "S": .Caption = "99"
        
        End Select
    
    End With
    
End Sub

'******************************************************************************
'* Functional Description   :   Validates for Max_weight > Min_weight
'* Parameter Description    :   None
'* Return Type Description  :   Boolean - True if Maxweight> Minweight
'******************************************************************************

Private Function ValidateMaxWgt() As Boolean
        
    If Trim(txtMaxWeight.Text) = "" Then _
        GoTo CleanUpAndExit
    
    txtMinWeight.Text = Format(txtMinWeight.Text, "#0.00")
    
    If Val(txtMaxWeight.Text) < Val(txtMinWeight.Text) Then
        
        giErrMsg = ShowErrorMsg("ProLbl007", Me.Caption, vbOKOnly)
        ValidateMaxWgt = False
        Exit Function
    
    End If
        
CleanUpAndExit:

    ValidateMaxWgt = True
    
End Function

'******************************************************************************
'* Functional Description   :   Validates for
'*                              Min_weight > 0 if it is anew record
'*                              Min_Weight < Max_weight if it is upate
'* Parameter Description    :   None
'* Return Type Description  :   Boolean - True if validation is success
'******************************************************************************

Private Function ValidateMinWgt() As Boolean
            
    If Trim(txtMinWeight.Text) = "" Then GoTo CleanUpAndExit
    
    txtMinWeight.Text = Format(txtMinWeight.Text, "#0.00")
    
    If fcolProduct Is Nothing Then
        
        If Val(txtMinWeight.Text) <= -1000000 Then
            
            giErrMsg = ShowErrorMsg("ProLbl008", Me.Caption, vbOKOnly)
            ValidateMinWgt = False
            Exit Function
        
        End If
    
    Else
        
        If txtMaxWeight.Text <> "" And Left(cboWeightType.Text, 1) <> "F" Then
            
            If Val(txtMinWeight.Text) > Val(txtMaxWeight.Text) Then
                
                giErrMsg = ShowErrorMsg("ProLbl009", Me.Caption, vbOKOnly)
                ValidateMinWgt = False
                Exit Function
            
            End If
        
        End If
    
    End If

CleanUpAndExit:

    ValidateMinWgt = True
    
End Function

'******************************************************************************
'* Functional Description   :   Check if any screen fields are empty
'* Parameter Description    :   None
'* Return Type Description  :   Boolean - True if Record is blank
'******************************************************************************

Private Function RecordIsBlank() As Boolean

Dim lIndex As Long

    If Trim(txtProdCode.Text) <> "" Or _
            Trim(txtDescription.Text) <> "" Or _
            Trim(txtLabelNo.Text) <> "" Or _
            Trim(txtBoxesPallet.Text) <> "" Or _
            Trim(txtMinWeight.Text) <> "" Or _
            Trim(txtMaxWeight.Text) <> "" Or _
            Trim(txtTareWeight.Text) <> "" Or _
            Trim(txtLabelLength.Text) <> "" Or _
            Trim(txtStrPosition.Text) <> "" Or _
            Trim(txtWeightLength.Text) <> "" Or _
            Trim(txtCommodityCode.Text) <> "" Or _
            Trim(txtLabelCodeStr.Text) <> "" Or _
            Trim(txtLabelCodeLength.Text) <> "" Or _
            cboProductType.Text <> "" Or _
            cboWeightType.Text <> "" Or _
            cboGovtLot.Text <> "" Or _
            cboMfgDate.Text <> "" Or _
            cboLblCodeChk.Text <> "" Or cboPrdgrpcode.Text <> "" Then 'Added by TCS
                    
        Exit Function
    
    End If
    
    RecordIsBlank = True
    
End Function

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cboProductType_GotFocus()

    fsControlText = cboProductType.Text
    
End Sub

''******************************************************************************
''* Functional Description   :   Set a variable for original value
''* Parameter Description    :   None
''* Return Type Description  :   None
''******************************************************************************
'
'Private Sub cboBlastIndex_GotFocus()
'
'    fsControlText = cboBlastIndex.Text
'
'End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cboGovtLot_GotFocus()

    fsControlText = cboGovtLot.Text
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cboMfgDate_GotFocus()

    fsControlText = cboMfgDate.Text
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

'Private Sub cboCustomerType_GotFocus()
'
'    fsControlText = cboCustomerType.Text
'
'End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub cboWeightType_GotFocus()

    fsControlText = cboWeightType.Text
    
End Sub

'******************************************************************************
'* Functional Description   :   Update Database and unload form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

fbUnload = True
    
    If fbLoadSuccess Then
    If fbUpdateDB Then
        
        'In Update mode, validate record for not null fields
        'and then update to database
        'In Add mode, if validation for not null fileds fails,
       'unload screen without updating database
        
    
       ' If UpdateDetailInventory Then
       
        
       
           If Not UpdateSecondaryCustomers Then
             Cancel = 1
             fbUnload = False
           End If
           
'        Else
'           Cancel = 1
'        End If
        
    End If
    End If
    
    If Cancel = 1 Then
       fbUpdateDB = False
    End If

End Sub







'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtBoxesPallet_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForNumerics txtBoxesPallet, sOrigValue
    txtBoxesPallet = sOrigValue
    
End Sub

'******************************************************************************
'* Functional Description   :   Validation
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtBoxesPallet_KeyPress(KeyAscii As Integer)

    ReturnKeyForNumerics txtBoxesPallet, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Validation
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtBoxesPallet_Validate(Cancel As Boolean)

    If fbFormIsLoading Then Exit Sub
   
    With txtBoxesPallet
         
         If Trim(.Text) <> "" Then
            
            If Val(.Text) <= 0 Then
                
                giErrMsg = ShowErrorMsg("ProLbl010", Me.Caption, vbOKOnly)
                GoTo CleanUpAndExit
            
            End If
         
         End If
     
     End With

     Exit Sub
    
CleanUpAndExit:
    
    Cancel = True
    txtBoxesPallet.Text = fsControlText
    txtBoxesPallet.SelStart = 0
    txtBoxesPallet.SelLength = Len(txtBoxesPallet.Text)
    txtBoxesPallet.SetFocus
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCommodityCode_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForNumerics txtCommodityCode, sOrigValue
    txtCommodityCode = sOrigValue
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtCommodityCode_GotFocus()
    
    fsControlText = txtCommodityCode.Text
    txtCommodityCode.SelStart = 0
    txtCommodityCode.SelLength = Len(txtCommodityCode.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Calls ValidateForIntegers Function
'* Parameter Description    :   KeyAscii - value based on current key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtCommodityCode_KeyPress(KeyAscii As Integer)
    
    ReturnKeyForNumerics txtCommodityCode, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Check if Commodity Code is valid
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtCommodityCode_Validate(Cancel As Boolean)

Dim bError As Boolean

    If fbFormIsLoading Then Exit Sub
    
    With txtCommodityCode
        
        If Trim(.Text) <> "" Then
            
            If Not CommodityCodeIsValid(.Text, bError) Then
                
                If Not bError Then
                    
                    giErrMsg = ShowErrorMsg("ProLbl011", Me.Caption, vbOKOnly)
                    GoTo CleanUpAndExit
                
                End If
                
                GoTo CleanUpAndExit
            
            End If
        
        End If
        
        If Trim(.Text) = "" Then
            
            giErrMsg = ShowErrorMsg("ProLbl011", Me.Caption, vbOKOnly)
            Cancel = True
            GoTo CleanUpAndExit
        
        End If
    
    End With
    
    Exit Sub
    
CleanUpAndExit:

    Cancel = True
    txtCommodityCode.Text = fsControlText
    txtCommodityCode.SelStart = 0
    txtCommodityCode.SelLength = Len(txtCommodityCode.Text)
    txtCommodityCode.SetFocus
    
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    
    If KeyAscii >= 97 And KeyAscii <= 122 Then
                    
        KeyAscii = KeyAscii - 32
                        
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtBoxesPallet_GotFocus()
    
    fsControlText = txtBoxesPallet.Text
    txtBoxesPallet.SelStart = 0
    txtBoxesPallet.SelLength = Len(txtBoxesPallet.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtDescription_GotFocus()
    
    fsControlText = txtDescription.Text
    txtDescription.SelStart = 0
    txtDescription.SelLength = Len(txtDescription.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLabelCodeLength_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForNumerics txtLabelCodeLength, sOrigValue
    txtLabelCodeLength = sOrigValue
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtLabelCodeLength_GotFocus()
    
    fsControlText = txtLabelCodeLength.Text
    txtLabelCodeLength.SelStart = 0
    txtLabelCodeLength.SelLength = Len(txtLabelCodeLength.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLabelCodeStr_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForNumerics txtLabelCodeStr, sOrigValue
    txtLabelCodeStr = sOrigValue
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtLabelCodeStr_GotFocus()
    
    fsControlText = txtLabelCodeStr.Text
    txtLabelCodeStr.SelStart = 0
    txtLabelCodeStr.SelLength = Len(txtLabelCodeStr.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtLabelLength_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForNumerics txtLabelLength, sOrigValue
    txtLabelLength = sOrigValue
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtLabelLength_GotFocus()
    
    fsControlText = txtLabelLength.Text
    txtLabelLength.SelStart = 0
    txtLabelLength.SelLength = Len(txtLabelLength.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtLabelNo_GotFocus()
    
    fsControlText = txtLabelNo.Text
    txtLabelNo.SelStart = 0
    txtLabelNo.SelLength = Len(txtLabelNo.Text)

End Sub

Private Sub txtLabelNo_KeyPress(KeyAscii As Integer)
    ReturnKeyForNumerics txtLabelNo, KeyAscii
End Sub

'******************************************************************************
'* Functional Description   :   Validation
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtLabelNo_Validate(Cancel As Boolean)

Dim lRowToExclude As Long

    If fbFormIsLoading Then Exit Sub
    
    If fcolProduct Is Nothing Then
        
        lRowToExclude = -1
    
    Else
        
        On Error Resume Next    'Added by TCS on 24-Aug-2005
        lRowToExclude = fcolProduct.Item("ROW_INDEX")
    
    End If
    
    If lRowToExclude = -1 And Trim(txtLabelNo.Text) <> "" Then
        
        If Val(txtLabelNo.Text) <= 0 Then
            
            giErrMsg = ShowErrorMsg("ProLbl013", Me.Caption, vbOKOnly)
            GoTo CleanUpAndExit
        
        End If
    
    End If
    Dim strLabelNo As String
    Dim intLabelNo As Integer
    Dim lResult As Long
    
    intLabelNo = CInt(txtLabelNo.Text)
    strLabelNo = CStr(intLabelNo)
    
    If ItemExistsInCollection(fcolLabelCodes, _
       strLabelNo & "|" & Left(cboProductType.Text, 1), _
        strLabelNo & "|" & Left(cboProductType.Text, 1)) Then
        
         lResult = 0
         Dim rsProductCodes      As ADODB.Recordset
         Dim strProducts         As String
         Dim objProduct          As Object
         
        'Create object of ProductMaster Class in MasterFileFunctions
    
        Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
    
        'Retrieve duplicate Product Codes
    
        lResult = objProduct.GetDuplicateProductCodes(strLabelNo, _
                                                      Left(cboProductType.Text, 1), _
                                                      rsProductCodes)
        
        'Condition added by TCS-Ragu on 20-Jun-2005
        If Not rsProductCodes Is Nothing Then   'Condition added by TCS-Ragu on 20-Jun-2005
            rsProductCodes.MoveFirst

            strProducts = rsProductCodes.Fields("PRODUCT_CODE").Value
            rsProductCodes.MoveNext
            While Not rsProductCodes.EOF
                strProducts = strProducts + " - " + rsProductCodes.Fields("PRODUCT_CODE").Value
                rsProductCodes.MoveNext
            Wend
            
            gcolErrMsg.Add strProducts
            giErrMsg = ShowErrorMsg("ProLbl032", Me.Caption, vbOKOnly, gcolErrMsg)
        
        End If  'Condition added by TCS-Ragu on 20-Jun-2005
        GoTo CleanUpAndExit
    
    End If
    
    Exit Sub

CleanUpAndExit:

    Cancel = True
    txtLabelNo.Text = fsControlText
    txtLabelNo.SelStart = 0
    txtLabelNo.SelLength = Len(txtLabelNo.Text)
    txtLabelNo.SetFocus
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtMaxWeight_GotFocus()

    fsControlText = txtMaxWeight.Text
    txtMaxWeight.Text = Format(txtMaxWeight.Text, "#0.00")
    txtMaxWeight.SelStart = 0
    txtMaxWeight.SelLength = Len(txtMaxWeight.Text)

End Sub

'******************************************************************************
'* Functional Description   :   Format data
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtMaxWeight_Validate(Cancel As Boolean)

    If fbFormIsLoading Then Exit Sub

    With txtMaxWeight
    
        If Left(cboWeightType.Text, 1) = "F" Then
            
            If ValidateNegNumFieldSize(txtMinWeight, 7, 2, "Max.Weight") Then
                
                .Text = txtMinWeight.Text
                Exit Sub
            
            End If
        
        End If
        
        
        If Not ValidateNegNumFieldSize(txtMaxWeight, 7, 2, "Max.Weight") Then
            
            GoTo CleanUpAndExit
        
        End If
        
        
        If Not ValidateMaxWgt Then GoTo CleanUpAndExit

        .Text = Format(.Text, "#0.00")
    
    End With

    Exit Sub
    
CleanUpAndExit:
    
    Cancel = True
    txtMaxWeight.Text = fsControlText
    txtMaxWeight.SelStart = 0
    txtMaxWeight.SelLength = Len(txtMaxWeight.Text)
    txtMaxWeight.SetFocus

End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtMinWeight_GotFocus()

    fsControlText = txtMinWeight.Text
    txtMinWeight.Text = Format(txtMinWeight.Text, "#0.00")
    txtMinWeight.SelStart = 0
    txtMinWeight.SelLength = Len(txtMinWeight.Text)
    
End Sub




'******************************************************************************
'* Functional Description   :   Accept only numeric characters
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtLabelLength_KeyPress(KeyAscii As Integer)

    ReturnKeyForNumerics txtLabelLength, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtProdCode_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForCapsAndNumerics txtProdCode, sOrigValue
    txtProdCode = sOrigValue
            
    ustgrdSecondaryCustomers.CellValue(1, 0) = Trim(txtProdCode.Text)
        
    
End Sub

' Added by TCS
'******************************************************************************
'* Functional Description   :   Validate for alpahnumeric and upper case entries
'* Parameter Description    :   Keyascii value of key pressed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtRefCode_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtStrPosition_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForNumerics txtStrPosition, sOrigValue
    txtStrPosition = sOrigValue
    
End Sub

'******************************************************************************
'* Functional Description   :   Accept only numeric characters
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtStrPosition_KeyPress(KeyAscii As Integer)

    ReturnKeyForNumerics txtStrPosition, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the change event of the text box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub txtWeightLength_Change()

'Holds the previous value in the text

Static sOrigValue       As String

    ReturnValueForNumerics txtWeightLength, sOrigValue
    txtWeightLength = sOrigValue
    
End Sub

'******************************************************************************
'* Functional Description   :   Accept only numeric characters
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtWeightLength_KeyPress(KeyAscii As Integer)

    ReturnKeyForNumerics txtWeightLength, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Accept only numeric characters
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtLabelCodeLength_KeyPress(KeyAscii As Integer)

    ReturnKeyForNumerics txtLabelCodeLength, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Accept only numeric characters
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtLabelCodeStr_KeyPress(KeyAscii As Integer)

    ReturnKeyForNumerics txtLabelCodeStr, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Format data
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtMinWeight_Validate(Cancel As Boolean)
    
    If fbFormIsLoading Then Exit Sub
    
    With txtMinWeight
    
        .Text = Format(.Text, "#0.00")
        
        If Not ValidateNegNumFieldSize(txtMinWeight, 7, 2, "Min.Weight") Then
            
            txtMinWeight.Text = fsControlText
            GoTo CleanUpAndExit
        
        End If
     
        If Not ValidateMinWgt Then GoTo CleanUpAndExit
        
        .Text = Format(.Text, "#0.00")
        
        If Left(cboWeightType.Text, 1) = "F" Then
            
            txtMaxWeight.Text = .Text
        
        End If
        
        .Text = Format(.Text, "#0.00")
        
    End With
    
    Exit Sub

CleanUpAndExit:
    
    Cancel = True
    txtMinWeight.SelStart = 0
    txtMinWeight.SelLength = Len(txtMinWeight.Text)
    txtMinWeight.SetFocus
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtStrPosition_GotFocus()

    fsControlText = txtStrPosition.Text
    txtStrPosition.SelStart = 0
    txtStrPosition.SelLength = Len(txtStrPosition.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtTareWeight_GotFocus()

    fsControlText = txtTareWeight.Text
    txtTareWeight.SelStart = 0
    txtTareWeight.SelLength = Len(txtTareWeight.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Format data
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtTareWeight_Validate(Cancel As Boolean)

    If fbFormIsLoading Then Exit Sub

    With txtTareWeight
        
        If Trim(.Text) <> "" Then
            
            If Not ValidateNegNumFieldSize(txtTareWeight, 7, 2, "Tare Weight") Then
                
                GoTo CleanUpAndExit
            
            End If
            
            If fcolProduct Is Nothing Then
                
                If Val(.Text) <= -100000 Then
                    
                    giErrMsg = ShowErrorMsg("ProLbl006", Me.Caption, vbOKOnly)
                    GoTo CleanUpAndExit
                
                End If
            
            End If
        
        End If
    
    End With

    Exit Sub
    
CleanUpAndExit:

    Cancel = True
    txtTareWeight.Text = fsControlText
    txtTareWeight.SelStart = 0
    txtTareWeight.SelLength = Len(txtTareWeight.Text)
    txtTareWeight.SetFocus

End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtWeightLength_GotFocus()

    fsControlText = txtWeightLength.Text
    txtWeightLength.SelStart = 0
    txtWeightLength.SelLength = Len(txtWeightLength.Text)
    
End Sub

'******************************************************************************
'* Functional Description   :   Set a variable for original value
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtProdCode_GotFocus()

    fsControlText = txtProdCode.Text
    txtProdCode.SelStart = 0
    txtProdCode.SelLength = Len(txtProdCode.Text)

End Sub

'******************************************************************************
'* Functional Description   :   Validates for alphanumeric entry only.
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtProdCode_KeyPress(KeyAscii As Integer)

    ReturnKeyForCapsAndNumerics txtProdCode, KeyAscii
    
End Sub

'******************************************************************************
'* Functional Description   :   Validate Product Key for duplication
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub txtProdCode_Validate(Cancel As Boolean)
    
Dim objProduct      As Object
Dim rsProduct       As ADODB.Recordset
Dim lResult         As Long
Dim bCheckDuplicate As Boolean

    On Error GoTo ErrHandler
    
    lResult = 0
    
    If fbFormIsLoading Or fsControlText = Trim(txtProdCode.Text) Then Exit Sub
    
    If fcolProduct Is Nothing Then bCheckDuplicate = True
    
    If Not fcolProduct Is Nothing Then
        
        If txtProdCode.Text <> fcolProduct.Item("PRODUCT_CODE_KEY") Then
            
            bCheckDuplicate = True
        
        End If
    
    End If
    
    If bCheckDuplicate Then
        
        If ItemExistsInCollection(fcolProducts, Trim(txtProdCode.Text), _
                                    Trim(txtProdCode.Text)) Then
            
            giErrMsg = ShowErrorMsg("ProLbl015", Me.Caption, vbOKOnly)
            GoTo CleanUpAndExit
        
        ElseIf Trim(txtProdCode.Text) <> "" Then
        
            'Create object of ProductMaster Class in MasterFileFunctions
            
            Set objProduct = _
                CreateObject("MasterFileFunctions.ProductMaster")
                
            'Check for Product Duplication
            
            lResult = objProduct.CheckProductDuplication(Trim( _
                                            txtProdCode.Text), _
                                            rsProduct)
            If lResult <> 0 Then GoTo ErrHandler
            
            If lResult = 0 Then
                
                If Not rsProduct.EOF Then
                   
                   ' Commented to not show "Duplicate Product Code" message if it's in PROD_MSTR but not PROD_PLANT
                   'giErrMsg = ShowErrorMsg("ProLbl016", Me.Caption, vbOKOnly)
                   
                   DisplayProductDetails rsProduct
                
                Else
                    
                    SetMasterFieldsForEdit True
                
                End If
                
                If Not rsProduct.EOF Then
                    
                    txtBoxesPallet.SetFocus
                
                Else
                    
                    txtDescription.SetFocus
                
                End If
            
            End If
        
        ElseIf Trim(txtProdCode.Text) = "" Then
            
            SetMasterFieldsForEdit True
            txtDescription.SetFocus
        
        End If
    
    End If
    
    Exit Sub
    
CleanUpAndExit:

    Cancel = True
    Set objProduct = Nothing
    Set rsProduct = Nothing
    txtProdCode.Text = fsControlText
    txtProdCode.SelStart = 0
    txtProdCode.SelLength = Len(txtProdCode.Text)
    txtProdCode.SetFocus
    Exit Sub
    
ErrHandler:
    
    If lResult <> 0 Then
            
        'Server side error
        
        gcolErrMsg.Add lResult
        giErrMsg = ShowErrorMsg("ProLbl029", Me.Caption, vbOKOnly, gcolErrMsg)
    
    Else
            
        'VB error
        
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("ProLbl030", Me.Caption, vbOKOnly, gcolErrMsg)
    
    End If
    
    GoTo CleanUpAndExit
    
End Sub

'******************************************************************************
'* Functional Description   :   Validates Commodity Code
'* Parameter Description    :   sCommodityCode - Code to be validated
'* Return Type Description  :   Boolean if successful or not
'******************************************************************************

Public Function CommodityCodeIsValid(ByVal sCommodityCode As String, _
                                     ByRef bError As Boolean) As Boolean
    
Dim objProduct              As Object
Dim lResult                 As Long
Dim bCodeIsValid            As Boolean

    On Error GoTo ErrHandler
    
    bError = False
    
    'Create object of ProductMaster Class in MasterFileFunctions
    
    Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
    
    'Validate Commodity Code
    
    lResult = objProduct.ValidateCommodityCode(sCommodityCode, bCodeIsValid)
    
    If lResult <> 0 Then GoTo ErrHandler
    
    CommodityCodeIsValid = bCodeIsValid
    
    Exit Function

ErrHandler:
    
    bError = True
    CommodityCodeIsValid = False
    
    If lResult <> 0 Then
            
        'Server side error
        
        gcolErrMsg.Add lResult
        giErrMsg = ShowErrorMsg("ProLbl031", Me.Caption, vbOKOnly, gcolErrMsg)
    
    Else
            
        'VB error
        
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("ProLbl031", Me.Caption, vbOKOnly, gcolErrMsg)
    
    End If

End Function

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTareWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyV And Shift = 2 Then
    
        If CheckNegPastedSelPart(txtTareWeight.Text, _
                                    txtTareWeight.SelText, _
                                    txtTareWeight.SelStart, _
                                    Clipboard.GetText, 5, 2) Then
                                    
            txtTareWeight.Text = GetPastedText(txtTareWeight.Text, _
                                                txtTareWeight.SelText, _
                                                txtTareWeight.SelStart, _
                                                Clipboard.GetText)
            
        End If
        
    ElseIf KeyCode = 46 Then
    
        ' Check if delete key is pressed.
        ' This will not be got in the keypress event
        
        If Not CheckNegSelectedPart(txtTareWeight.Text, _
                                txtTareWeight.SelText, _
                                txtTareWeight.SelStart, 0, 5, 2) Then
                                
            KeyCode = 0
            
        End If
        
    End If
    
End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTareWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtTareWeight.Text)
    iSelLength = txtTareWeight.SelLength
    iSelStart = txtTareWeight.SelStart
    sSelText = txtTareWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 5, 2, iSelLength) Then
    
        KeyAscii = 0
        
        If sSelText = "" Then
        
            txtTareWeight.SelStart = iSelStart
            txtTareWeight.SelLength = iSelLength
        
        End If
        
    End If
    
End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTareWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String
    
    sText = Trim(txtTareWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtTareWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   To Format the field to ####0.00.
'* Parameter Description    :   None
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTareWeight_LostFocus()

Dim sText As String
    
    sText = Trim(txtTareWeight.Text)
    
    If RoundOffDecimalNumber(sText, 5, 2, True) Then
        
        txtTareWeight.Text = sText
        
    Else
    
        txtTareWeight.SetFocus
            
    End If
    
End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTareWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = MouseButtonConstants.vbRightButton Then
    
        fbCanBeCopied = False
        fsCopyString = txtTareWeight.Text
        
        If CheckNegPastedSelPart(Trim(txtTareWeight.Text), txtTareWeight.SelText, _
                                txtTareWeight.SelStart, Clipboard.GetText, 5, 2) Then
        
            fbCanBeCopied = True
            
        End If
        
    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtTareWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then
    
        If Not fbCanBeCopied Then
        
             txtTareWeight.Text = fsCopyString
             fbCanBeCopied = True
             
        End If
    End If

    sText = Trim(txtTareWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtTareWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMaxWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyV And Shift = 2 Then
    
        If CheckNegPastedSelPart(txtMaxWeight.Text, _
                                    txtMaxWeight.SelText, _
                                    txtMaxWeight.SelStart, _
                                    Clipboard.GetText, 5, 2) Then
                                    
            txtMaxWeight.Text = GetPastedText(txtMaxWeight.Text, _
                                                txtMaxWeight.SelText, _
                                                txtMaxWeight.SelStart, _
                                                Clipboard.GetText)
            
        End If
        
    ElseIf KeyCode = 46 Then
    
        ' Check if delete key is pressed.
        ' This will not be got in the keypress event
        
        If Not CheckNegSelectedPart(txtMaxWeight.Text, _
                                txtMaxWeight.SelText, _
                                txtMaxWeight.SelStart, 0, 5, 2) Then
                                
            KeyCode = 0
            
        End If
        
    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMaxWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtMaxWeight.Text)
    iSelLength = txtMaxWeight.SelLength
    iSelStart = txtMaxWeight.SelStart
    sSelText = txtMaxWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 5, 2, iSelLength) Then
    
        KeyAscii = 0
        
        If sSelText = "" Then
        
            txtMaxWeight.SelStart = iSelStart
            txtMaxWeight.SelLength = iSelLength
        
        End If
        
    End If

End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMaxWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String
    
    sText = Trim(txtMaxWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtMaxWeight.Text = sText
    
End Sub

'*******************************************************************************
'* Functional Description   :   To Format the field to ####0.00.
'* Parameter Description    :   None
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMaxWeight_LostFocus()

Dim sText As String
    
    sText = Trim(txtMaxWeight.Text)
    
    If RoundOffDecimalNumber(sText, 5, 2, True) Then
        
        txtMaxWeight.Text = sText
        
    Else
    
        txtMaxWeight.SetFocus
            
    End If
    
End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMaxWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = MouseButtonConstants.vbRightButton Then
    
        fbCanBeCopied = False
        fsCopyString = txtMaxWeight.Text
        
        If CheckNegPastedSelPart(Trim(txtMaxWeight.Text), txtMaxWeight.SelText, _
                                txtMaxWeight.SelStart, Clipboard.GetText, 5, 2) Then
        
            fbCanBeCopied = True
            
        End If
        
    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMaxWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then
    
        If Not fbCanBeCopied Then
        
             txtMaxWeight.Text = fsCopyString
             fbCanBeCopied = True
             
        End If
    End If

    sText = Trim(txtMaxWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtMaxWeight.Text = sText

End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Pasted texts and delete key.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMinWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyV And Shift = 2 Then
    
        If CheckNegPastedSelPart(txtMinWeight.Text, _
                                    txtMinWeight.SelText, _
                                    txtMinWeight.SelStart, _
                                    Clipboard.GetText, 5, 2) Then
                                    
            txtMinWeight.Text = GetPastedText(txtMinWeight.Text, _
                                                txtMinWeight.SelText, _
                                                txtMinWeight.SelStart, _
                                                Clipboard.GetText)
            
        End If
        
    ElseIf KeyCode = 46 Then
    
        ' Check if delete key is pressed.
        ' This will not be got in the keypress event
        
        If Not CheckNegSelectedPart(txtMinWeight.Text, _
                                txtMinWeight.SelText, _
                                txtMinWeight.SelStart, 0, 5, 2) Then
                                
            KeyCode = 0
            
        End If
        
    End If
    
End Sub

'*******************************************************************************
'* Functional Description   :   Validate for Numeric Entries.
'* Parameter Description    :   KeyAscii to get key press.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMinWeight_KeyPress(KeyAscii As Integer)

Dim sText       As String
Dim iSelLength  As Integer
Dim iSelStart   As Integer
Dim sSelText    As String

    sText = Trim(txtMinWeight.Text)
    iSelLength = txtMinWeight.SelLength
    iSelStart = txtMinWeight.SelStart
    sSelText = txtMinWeight.SelText

    If Not IsValidNegativeNumber(sText, KeyAscii, sSelText, _
                                iSelStart, 5, 2, iSelLength) Then
    
        KeyAscii = 0
        
        If sSelText = "" Then
        
            txtMinWeight.SelStart = iSelStart
            txtMinWeight.SelLength = iSelLength
        
        End If
        
    End If
    
End Sub

'*******************************************************************************
'* Functional Description   :   Corrects the code is '-' is entered in wrong place.
'* Parameter Description    :   KeyCode to get key code and shift
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMinWeight_KeyUp(KeyCode As Integer, Shift As Integer)

Dim sText As String
    
    sText = Trim(txtMinWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtMinWeight.Text = sText
    
End Sub

'*******************************************************************************
'* Functional Description   :   To Format the field to ####0.00.
'* Parameter Description    :   None
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMinWeight_LostFocus()

Dim sText As String
    
    sText = Trim(txtMinWeight.Text)
    
    If RoundOffDecimalNumber(sText, 5, 2, True) Then
        
        txtMinWeight.Text = sText
        
    Else
    
        txtMinWeight.SetFocus
            
    End If
    
End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMinWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = MouseButtonConstants.vbRightButton Then
    
        fbCanBeCopied = False
        fsCopyString = txtMinWeight.Text
        
        If CheckNegPastedSelPart(Trim(txtMinWeight.Text), txtMinWeight.SelText, _
                                txtMinWeight.SelStart, Clipboard.GetText, 5, 2) Then
        
            fbCanBeCopied = True
            
        End If
        
    End If

End Sub

'*******************************************************************************
'* Functional Description   :   To Validate the pasted characters
'* Parameter Description    :   Defauls parameters
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub txtMinWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim sText As String

    If Button = MouseButtonConstants.vbRightButton Then
    
        If Not fbCanBeCopied Then
        
             txtMinWeight.Text = fsCopyString
             fbCanBeCopied = True
             
        End If
    End If

    sText = Trim(txtMinWeight.Text)
    Call EvaluateNagativeNumber(sText)
    txtMinWeight.Text = sText

End Sub

' Added by TCS

Private Sub SizeGridColumns()

    With ustgrdSecondaryCustomers
        
        .ColWidth(cnProduct) = .Width / 6
        .ColWidth(cnOriginPlant) = .Width / 9
        .ColWidth(cnCustNumber) = .Width / 7
        .ColWidth(cnCustType) = .Width / 8.5
        .ColWidth(cnFreezerDays) = .Width / 10.5
        .ColWidth(cnBlastDays) = .Width / 11
        .ColWidth(cnAction) = 0
'Added by TCS for Req 22 ver1
        .ColWidth(cnstatus) = .Width / 6.5 '10    'Changed by TCS-Ragu on 25-Aug-2005
        .ColWidth(cnDefaultOrigin) = .Width / 9
    End With
    
End Sub

' Added by TCS

Private Sub FormatGrid()

Dim iColIndex               As Integer
    
    'Initializing various properties of the Grid
    
    fsProduct = Trim(txtProdCode.Text) ' Added by TCS

    With ustgrdSecondaryCustomers
        
        .FixedRows = 1
        .RowHeight(0) = 2.5 * GRID_ROW_HEIGHT
        .WordWrap = True
        .Columns = COLUMN_COUNT
        
        .CellValue(0, cnProduct) = "Product Code"
        .CellValue(0, cnOriginPlant) = "Origin"
        .CellValue(0, cnCustNumber) = "Customer No."
        .CellValue(0, cnCustType) = "Customer Type"
        .CellValue(0, cnFreezerDays) = "Freezer Days"
        .CellValue(0, cnBlastDays) = "Blast Days"
        .CellValue(0, cnAction) = "Action"
        .CellValue(0, cnPrevOriginPlant) = "Previous Orgin Plant"
        .CellValue(0, cnPrevCustNumber) = "Previous Cust Number"
        .CellValue(0, cnPrevCustType) = "Previous Cust Type"
        
'Added by TCS  for Req 22 ver1
        .CellValue(0, cnstatus) = "Default Status"
        .CellValue(0, cnDefaultOrigin) = "Default Origin"
        
        'Format the Header Row
        
        .Row = 0
               
        For iColIndex = 0 To COLUMN_COUNT - 1
            
            .ColAlignmentFixed(iColIndex) = flexAlignCenterCenter
            
            Select Case iColIndex
                
                Case cnProduct, cnOriginPlant, _
                                cnCustNumber
                    .ColAlignment(iColIndex) = flexAlignLeftCenter
                
                Case cnBlastDays, cnFreezerDays
                    .ColAlignment(iColIndex) = flexAlignRightCenter
            
            End Select
        
        Next
        
        
        
        'Set the Column Types based on the access rights
        
        'Normal   - User will not be able to edit but only browse
        'Editbox  - User will be able to handle it like a text box
        'Combobox - Control is a combo box

'        If fsAccessRights = FUNC_BROWSE_0NLY Then
'
'            .ColumnType(cnProduct) = Normal
'            .ColumnType(cnOriginPlant) = Normal
'            .ColumnType(cnCustNumber) = Normal
'            .ColumnType(cnCustType) = Normal
'            .ColumnType(cnFreezerDays) = Normal
'            .ColumnType(cnBlastDays) = Normal
        
'        Else
            
            .ColumnType(cnProduct) = Normal
            .ColumnType(cnOriginPlant) = EditBox
            .ColumnType(cnCustNumber) = EditBox
            .ColumnType(cnCustType) = ComboBox
            .ColumnType(cnFreezerDays) = EditBox
            .ColumnType(cnBlastDays) = EditBox
            
'Added by TCS  for Req 22 ver1

            'Added by TCS-Ragu on 25-Aug-2005
            'Start
            '.ColumnType(cnstatus) = EditBox    'Prev code commented
            .ColumnType(cnstatus) = ComboBox
            'End
            .ColumnType(cnDefaultOrigin) = ComboBox
'            .ColumnType(cnDefaultOrigin) = EditBox
            
'        End If
                    
                
        'Set the Column Data Types - Alphanumeric,Ascii,Numeric,Real
        
        .ColumnDataType(cnOriginPlant) = AlphaNumeric   'Datatype changed by TCS on Aug-19-2005
        .ColumnDataType(cnCustNumber) = AlphaNumeric
        .ColumnDataType(cnFreezerDays) = Numeric
        .ColumnDataType(cnBlastDays) = Numeric

        'Set the MaxLength of each column

        .MaxLength(cnOriginPlant) = 3
        .MaxLength(cnCustNumber) = 7
        .MaxLength(cnFreezerDays) = 1
        .MaxLength(cnBlastDays) = 1
        
        .ColWidth(cnPrevOriginPlant) = 0
        .ColWidth(cnPrevCustNumber) = 0
        .ColWidth(cnPrevCustType) = 0
        
        LoadOrginPlantCombo
        
        .Rows = 2
        .Row = 1
        .Column = cnOriginPlant 'Changed by TCS
        flRowIndex = 1
        flColumnIndex = cnOriginPlant
        
    End With
    
End Sub

' Added by TCS

Private Function LoadSecondaryCustomers() As Long

'Dim objProduct          As Object
'Dim rsSecCustomers      As ADODB.Recordset
Dim lRowCount           As Long
Dim lResult             As Long
Dim bFoundOriginPlant   As Boolean

    On Error GoTo ErrHandler
    
    lResult = 0
    
    'Create object of ProductMaster Class in MasterFileFunctions
    
'    Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
'
'    'Retrieve Secondary Customer details
'
'    lResult = objProduct.GetSecCustomersForProductPlant(gsPlantCode, _
'                                                        fsProduct, _
'                                                        rsSecCustomers)
'    If lResult <> 0 Then
'
'        GoTo ErrHandler
'
'    End If
'
'    If rsSecCustomers Is Nothing Then Unload Me
    
    With ustgrdSecondaryCustomers
             
        If Not fcolProduct Is Nothing Then
            
            
                        
            
            For lRowCount = 1 To Val(fcolProduct("RowNum"))
                
                    If .Rows < lRowCount + 1 Then
                        .AddItems ""
                    End If
                    .CellValue(lRowCount, cnProduct) = fsProduct
                    .CellValue(lRowCount, cnCustNumber) = _
                            fcolProduct("CUSTOMER_ID" & "_" & Trim(Str(lRowCount)))
                    .CellValue(lRowCount, cnCustType) = _
                            fcolProduct.Item("CUSTOMER_TYPE" & "_" & Trim(Str(lRowCount)))
                    .CellValue(lRowCount, cnFreezerDays) = _
                            fcolProduct.Item("FREEZE_DAYS" & "_" & Trim(Str(lRowCount)))
                    .CellValue(lRowCount, cnBlastDays) = _
                            fcolProduct.Item("BLAST_DAYS" & "_" & Trim(Str(lRowCount)))
                    .CellValue(lRowCount, cnOriginPlant) = _
                            IIf(IsNull(fcolProduct.Item("ORIGIN_PLANT_CODE" & "_" & Trim(Str(lRowCount)))), "", fcolProduct.Item("ORIGIN_PLANT_CODE" & "_" & Trim(Str(lRowCount))))
                    .CellValue(lRowCount, cnPrevOriginPlant) = _
                            IIf(IsNull(fcolProduct.Item("ORIGIN_PLANT_CODE" & "_" & Trim(Str(lRowCount)))), "", fcolProduct.Item("ORIGIN_PLANT_CODE" & "_" & Trim(Str(lRowCount))))
                    .CellValue(lRowCount, cnPrevCustNumber) = _
                            fcolProduct.Item("CUSTOMER_ID" & "_" & Trim(Str(lRowCount)))
                    .CellValue(lRowCount, cnPrevCustType) = _
                            fcolProduct.Item("CUSTOMER_TYPE" & "_" & Trim(Str(lRowCount)))
'Added by TCS for Req 22 ver 1
                    .CellValue(lRowCount, cnstatus) = _
                            IIf(IsNull(fcolProduct.Item("Default_Status" & "_" & Trim(Str(lRowCount)))), "", fcolProduct.Item("Default_Status" & "_" & Trim(Str(lRowCount))))
                    .CellValue(lRowCount, cnDefaultOrigin) = _
                            IIf(IsNull(fcolProduct.Item("Default_Origin" & "_" & Trim(Str(lRowCount)))), "", fcolProduct.Item("Default_Origin" & "_" & Trim(Str(lRowCount))))
'Addition End here
                    
                    .CellValue(lRowCount, cnAction) = ""
             Next
                                   
                
                    
           
            
        Else
            
             For lRowCount = 1 To .Rows - 1
                 
                 .CellValue(lRowCount, cnProduct) = fsProduct
                 .CellValue(lRowCount, cnAction) = NEW_ROW
             
             Next
            
        End If

        ReDim fsArrCustType(3)
        fsArrCustType(0) = "B"
        fsArrCustType(1) = "C"
        fsArrCustType(2) = "S"
        .AddComboData cnCustType, fsArrCustType
        
        'Added by TCS-Ragu on 25-Aug-2005
        .AddComboData cnstatus, fnGetPRODStatus
        
        ReDim fsArrDefOrigin(2)
        fsArrDefOrigin(0) = "Y"
        fsArrDefOrigin(1) = "N"
        .AddComboData cnDefaultOrigin, fsArrDefOrigin
        
    End With
    
CleanUpAndExit:

    'Set objProduct = Nothing
    'Set rsSecCustomers = Nothing
    Exit Function
    
ErrHandler:

    If lResult <> 0 Then
        
        'Server side error
        
        gcolErrMsg.Add lResult
        giErrMsg = ShowErrorMsg("Second001", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    Else
        
        'VB error
        
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("Second010", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    End If
    
    cmdCancel_Click
    GoTo CleanUpAndExit
    
End Function
'Added by TCS-Ragu on 25-Aug-2005
'Start
Private Function fnGetPRODStatus() As String()

    Dim objProdMstr         As Object
    Dim lResult             As Long
    Dim I As Long
    
    Dim ProdStat() As String

    On Error GoTo ErrHandler
    
    
    Dim rsProdStat As New ADODB.Recordset
    
    
    'Create object of ProductMaster Class in MasterFileFunctions
    
  Set objProdMstr = CreateObject("MasterFileFunctions.ProductMaster")
         
    'Populate the Product Types
    
    lResult = objProdMstr.GetPRODStatusTypes(rsProdStat)
    If lResult <> 0 Then GoTo ErrHandler
    
    If rsProdStat.RecordCount > 0 Then
        ReDim ProdStat(rsProdStat.RecordCount)
    End If
    
    For I = 1 To rsProdStat.RecordCount
        ProdStat(I - 1) = rsProdStat.Fields("TYPE_CODE")
        rsProdStat.MoveNext
    Next I
    
    fnGetPRODStatus = ProdStat
        
CleanUpAndExit:
    
    Set objProdMstr = Nothing
    Exit Function
    
ErrHandler:
        
End Function
'End

' Added by TCS

Private Function UpdateSecondaryCustomers() As Boolean

Dim objProduct              As Object
Dim colSecCustomers         As Collection
Dim colSecCustomer          As Collection
Dim lRowIndex               As Long
Dim lResult                 As Long
Dim bValidRow               As Boolean

' from update detail inventory
'Dim objProduct          As Object
Dim colProduct          As Collection
Dim colProductClone     As Collection
Dim colProducts         As Collection
'Dim lResult             As Long
Dim sAction             As String
Dim sProdKey            As String
Dim sStatusText         As String
Dim bDuplLblCode        As Boolean
Dim bDefaultOrgin       As Boolean 'Added by venkat to validate at least one default orgint should be Y
Dim iDefaultOriginCount As Integer

On Error GoTo ErrHandler
  
UpdateSecondaryCustomers = True

' Added by TCS
sBlastInd = "N"
iDefaultOriginCount = 0
    
If ValidateRecordForBlank Then

    bValidRow = True
    
    bDefaultOrgin = False
    
    For lRowIndex = 1 To ustgrdSecondaryCustomers.Rows - 1
    
        If Not ValidateRow(lRowIndex) Then
            
            bValidRow = False
            Exit For
            
        End If
        
        ' Added by TCS to initialize the 'Blast Ind'
         If Val(Trim(ustgrdSecondaryCustomers.CellValue(lRowIndex, cnBlastDays))) > 0 Then
             sBlastInd = "Y"
         End If
       'Addition ends here
       
       'Added by Venkat
       If ustgrdSecondaryCustomers.TextMatrix(lRowIndex, 11) = "Y" Then
            bDefaultOrgin = True
            iDefaultOriginCount = iDefaultOriginCount + 1
       End If
       'End
        
    Next
    'Added by Venkat
    If bDefaultOrgin = False And ustgrdSecondaryCustomers.Rows > 1 Then
        If ustgrdSecondaryCustomers.TextMatrix(1, 11) <> "" Then
            MsgBox "Default Location needs to be 'Y' for at least one line.", vbOKOnly + vbInformation, Me.Caption
            UpdateSecondaryCustomers = False
            GoTo CleanUpAndExit
        End If
    End If
    'End
    
    ' Only 1 default origin location allowed to be default 'Y'
    If iDefaultOriginCount > 1 Then
        MsgBox "Only 1 Default Origin allowed.", vbOKOnly + vbInformation, Me.Caption
        UpdateSecondaryCustomers = False
        GoTo CleanUpAndExit
    End If
    
    If bValidRow Then
        
        With ustgrdSecondaryCustomers
            
            'Construct the collection of records to be updated from the grid
        
            For lRowIndex = 1 To .Rows - 1

                If (Left(.CellValue(lRowIndex, cnAction), 1) = TO_INSERT Or _
                    Left(.CellValue(lRowIndex, cnAction), 1) = TO_DELETE Or _
                    Left(.CellValue(lRowIndex, cnAction), 1) = TO_DELETE_AFTER_UPDATE Or _
                    Left(.CellValue(lRowIndex, cnAction), 1) = TO_UPDATE) And _
                    Not RowIsBlank(lRowIndex) Then
                
                    'Create a colSecCustomer collection for each row of the grid
                     
                    Set colSecCustomer = New Collection
                    
                    If .CellValue(lRowIndex, cnPrevOriginPlant) = _
                            .CellValue(lRowIndex, cnOriginPlant) And _
                            .CellValue(lRowIndex, cnPrevCustNumber) = _
                            .CellValue(lRowIndex, cnCustNumber) And _
                            .CellValue(lRowIndex, cnCustType) = _
                            .CellValue(lRowIndex, cnPrevCustType) Then
                        
                        colSecCustomer.Add _
                            "U" & Left(.CellValue(lRowIndex, cnAction), 1), _
                            "ACTION"
                            
                    Else
                    
                        colSecCustomer.Add "I", "ACTION"
                        
                    End If
                    
                    If Left(.CellValue(lRowIndex, cnAction), 1) = TO_DELETE Or _
                        Left(.CellValue(lRowIndex, cnAction), 1) = TO_DELETE_AFTER_UPDATE Then
                        colSecCustomer.Add _
                                Right(Trim(.CellValue(lRowIndex, cnAction)), _
                                        Len(Trim(.CellValue(lRowIndex, cnAction))) - 1), _
                                "PRODUCT_CODE"
                    Else
                        colSecCustomer.Add _
                                Trim(.CellValue(lRowIndex, cnProduct)), _
                                "PRODUCT_CODE"
                    End If
                    colSecCustomer.Add _
                            Trim(.CellValue(lRowIndex, cnCustNumber)), _
                            "CUSTOMER_ID"
                    colSecCustomer.Add _
                            Left(Trim(.CellValue(lRowIndex, cnCustType)), 1), _
                            "CUSTOMER_TYPE"
                    colSecCustomer.Add _
                            Val(Trim(.CellValue(lRowIndex, cnFreezerDays))), _
                            "FREEZE_DAYS"
                    colSecCustomer.Add _
                            Val(Trim(.CellValue(lRowIndex, cnBlastDays))), _
                            "BLAST_DAYS"
                                                        
                    colSecCustomer.Add _
                            Trim(.CellValue(lRowIndex, cnOriginPlant)), _
                            "ORIGIN_PLANT_CODE"
                 
                 'Added by TCS on 17/01/05
                    
                    colSecCustomer.Add _
                            Trim(.CellValue(lRowIndex, cnstatus)), _
                            "Default_Status"
                    
                    colSecCustomer.Add _
                            Trim(.CellValue(lRowIndex, cnDefaultOrigin)), _
                            "Default_Origin"
                            
                    colSecCustomer.Add CStr(lRowIndex), "ROW_INDEX"
                    
                    
                            
                            
                            
                    'Append the colSecCustomer collection to the main
                    'collection colSecCustomers
                    
                    If colSecCustomers Is Nothing Then
                        
                        Set colSecCustomers = New Collection
                    
                    End If
                    
                    If .CellValue(lRowIndex, cnPrevOriginPlant) = _
                            .CellValue(lRowIndex, cnOriginPlant) Then
                        
                        colSecCustomers.Add colSecCustomer, _
                                .CellValue(lRowIndex, cnOriginPlant)
                                
                    Else
                        colSecCustomers.Add colSecCustomer, _
                                .CellValue(lRowIndex, cnOriginPlant)
                    End If
                    
                    If (.CellValue(lRowIndex, cnPrevOriginPlant) <> _
                        .CellValue(lRowIndex, cnOriginPlant) And _
                        Trim(.CellValue(lRowIndex, cnPrevOriginPlant)) <> "") Or _
                        ((.CellValue(lRowIndex, cnPrevCustNumber) <> _
                        .CellValue(lRowIndex, cnCustNumber) Or _
                        .CellValue(lRowIndex, cnPrevCustType) <> _
                        .CellValue(lRowIndex, cnCustType)) And _
                        (Trim(.CellValue(lRowIndex, cnPrevCustNumber)) <> "" Or _
                        Trim(.CellValue(lRowIndex, cnPrevCustType)) <> "")) Then

                        Set colSecCustomer = New Collection

                        colSecCustomer.Add "D", "ACTION"
                        If Left(.CellValue(lRowIndex, cnAction), 1) = TO_DELETE Or _
                            Left(.CellValue(lRowIndex, cnAction), 1) = TO_DELETE_AFTER_UPDATE Then
                            colSecCustomer.Add _
                                    Right(Trim(.CellValue(lRowIndex, cnAction)), _
                                            Len(Trim(.CellValue(lRowIndex, cnAction))) - 1), _
                                    "PRODUCT_CODE"
                        Else
                            colSecCustomer.Add _
                                    Trim(.CellValue(lRowIndex, cnProduct)), _
                                    "PRODUCT_CODE"
                        End If
                        colSecCustomer.Add _
                                Trim(.CellValue(lRowIndex, cnCustNumber)), _
                                "CUSTOMER_ID"
                        colSecCustomer.Add _
                                Left(Trim(.CellValue(lRowIndex, cnCustType)), 1), _
                                "CUSTOMER_TYPE"
                        colSecCustomer.Add _
                                Val(Trim(.CellValue(lRowIndex, cnFreezerDays))), _
                                "FREEZE_DAYS"
                        colSecCustomer.Add _
                                Val(Trim(.CellValue(lRowIndex, cnBlastDays))), _
                                "BLAST_DAYS"
                        colSecCustomer.Add _
                                Trim(.CellValue(lRowIndex, cnPrevOriginPlant)), _
                                "ORIGIN_PLANT_CODE"
                        colSecCustomer.Add _
                                Trim(.CellValue(lRowIndex, cnstatus)), _
                                "Default_Status"
                        colSecCustomer.Add _
                                Trim(.CellValue(lRowIndex, cnDefaultOrigin)), _
                                "Default_Origin"
                                
                        colSecCustomer.Add CStr(lRowIndex), "ROW_INDEX"
                        colSecCustomers.Add colSecCustomer, _
                                .CellValue(lRowIndex, cnOriginPlant) & "*DELETE*"
                    End If
                End If
            
            Next
            
         End With
         
     Else
     
       UpdateSecondaryCustomers = False
       GoTo CleanUpAndExit
     End If
            
' taken from updatedetailinventory

        If fcolProduct Is Nothing Then
            
            sAction = TO_INSERT
            sProdKey = Trim(txtProdCode.Text)
        
        ElseIf Not fcolProduct Is Nothing Then
            ''Condition Added by TCS - Venkat on 21 - Jun - 2005
            If fcolProduct(1) = "New" Then
                sAction = TO_INSERT
                sProdKey = Trim(txtProdCode.Text)
            Else
                sAction = TO_UPDATE
                sProdKey = txtProdCode.Text '  fcolProduct.Item("PRODUCT_CODE_KEY")
            End If
        End If
        
        If lgrdRowIndex > 0 Then
         UpdateRecord
        End If
        
         Set colProducts = New Collection
        
        For lRowIndex = 1 To ustgrdSecondaryCustomers.Rows - 1
        Set colProduct = New Collection
                      
        ''Venkat Doubt
        ''Condition Added by TCS - Venkat on 21 - Jun - 2005
        If ustgrdSecondaryCustomers.CellValue(lRowIndex, cnAction) = TO_IGNORE Or sAction = TO_INSERT Then
            sAction = TO_INSERT
        Else
            'sAction = ustgrdSecondaryCustomers.CellValue(lRowIndex, cnAction)  'Commented by TCS on 03-sep-2005
            sAction = Left(ustgrdSecondaryCustomers.CellValue(lRowIndex, cnAction), 1)  'Added by TCS on 03-sep-2005
        End If
        
                      
        colProduct.Add sAction, "ACTION"
        
        lgrdRowIndex = lRowIndex
        
        DisplayRecord
        
        
        
        colProduct.Add 1, "RECORDCOUNT"
        colProduct.Add Trim(txtProdCode.Text), "PRODUCT_CODE"
        
        colProduct.Add Trim(txtDescription.Text), "PRODUCT_DESC"
        colProduct.Add Left(cboProductType.Text, 1), "PRODUCT_TYPE"
        colProduct.Add Trim(lblDivisonCodeValue.Caption), "DIVISION_CODE"
        colProduct.Add Trim(txtLabelNo.Text), "LABEL_NO"
        colProduct.Add Left(cboWeightType.Text, 1), "WGT_TYPE_CODE"
        
        colProduct.Add Trim(txtCommodityCode.Text), "GOVT_COMMODITY_CODE"
        
        colProduct.Add Trim(txtBoxesPallet.Text), "BOXES_PER_PALLET"
        colProduct.Add Val(Trim(txtMinWeight.Text)), "MIN_BOX_WGT"
        colProduct.Add Val(Trim(txtMaxWeight.Text)), "MAX_BOX_WGT"
        colProduct.Add Val(Trim(txtTareWeight.Text)), "BOX_TARE_WEIGHT"
 ' Changed by TCS
        colProduct.Add sBlastInd, "BLAST_IND"
        colProduct.Add ustgrdSecondaryCustomers.CellValue(lgrdRowIndex, cnFreezerDays), "FREEZE_DAYS"
        colProduct.Add ustgrdSecondaryCustomers.CellValue(lgrdRowIndex, cnBlastDays), "BLAST_DAYS"
        
        colProduct.Add Val(Trim(txtLabelLength.Text)), "LABEL_LENGTH"
        colProduct.Add Val(Trim(txtStrPosition.Text)), "LABEL_WGT_ST_POS"
        colProduct.Add Val(Trim(txtWeightLength.Text)), "LABEL_WGT_LENGTH"
        
        colProduct.Add ustgrdSecondaryCustomers.CellValue(lgrdRowIndex, cnCustNumber), "CUSTOMER_ID"         'Changed by TCS
        colProduct.Add ustgrdSecondaryCustomers.CellValue(lgrdRowIndex, cnCustType), "CUSTOMER_TYPE"   'Changed by TCS
        ''Condition Added by TCS - Venkat on 21-Jun-2005
        colProduct.Add IIf(Trim(cboGovtLot.Text) = "", "N", cboGovtLot.Text), "GOVT_LOT_IND"
        colProduct.Add cboMfgDate.Text, "PACK_DATE_IND"
        
        'Added by TCS-Ragu on 17-Jun-2005
        colProduct.Add txtPckdtStart.Text, "PACK_DATE_START"
    
        colProduct.Add cboLblCodeChk.Text, "CHECK_LABEL_IND"
        colProduct.Add Val(Trim(txtLabelCodeStr.Text)), "PROD_LABEL_START"
        colProduct.Add Val(Trim(txtLabelCodeLength.Text)), "PROD_LABEL_LEN"
        colProduct.Add cboBoxSerialInd.Text, "BOX_SERIAL_IND"
          
        colProduct.Add Trim(txtRefCode.Text), "XREF_CODE"  ' Added by TCS
        
        colProduct.Add ustgrdSecondaryCustomers.CellValue(lgrdRowIndex, cnstatus), "DEFAULT_STATUS"
        colProduct.Add ustgrdSecondaryCustomers.CellValue(lgrdRowIndex, cnDefaultOrigin), "DEFAULT_ORIGIN"
        
        'Added by TCS-Ragu on 20-Jun-2005
        'Start
        colProduct.Add IIf(ustgrdSecondaryCustomers.CellValue(lgrdRowIndex, cnOriginPlant) = "", gsPlantCode, ustgrdSecondaryCustomers.CellValue(lgrdRowIndex, cnOriginPlant)), "ORIGIN_PLANT_CODE"
        'End
        
        colProduct.Add "Y", "SECOND_IND"  ' Added by TCS
        
        colProduct.Add Trim(cboPrdgrpcode.Text), "PRODUCT_GROUP_CODE"
       
        colProduct.Add sProdKey, "PRODUCT_CODE_KEY"
        
        colProduct.Add gsWinUserName, "USER_ID" 'harrels WR 10599 to log activtiy for tare wgt change
                
        Set colProductClone = New Collection
        Set colProductClone = colProduct
        
        colProducts.Add colProduct, sProdKey & Trim(Str(lgrdRowIndex))
        
        Next
        'Call the update method of ProductMaster to update the db
        
        If Not colProducts Is Nothing Then
            
            'Create object of ProductMaster in MasterFileFunctions
            
            Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
            
            sStatusText = sbStatusBar.Panels(1).Text
            sbStatusBar.Panels(1).Text = _
                                      "Updating Database - Please Wait..."
            Me.MousePointer = vbHourglass
            DoEvents
            
            'Update Product Details
            
            lResult = objProduct.UpdateProductForPlant(gsPlantCode, _
                                                    colProducts, bDuplLblCode)
            
            sbStatusBar.Panels(1).Text = sStatusText
            Me.MousePointer = vbDefault
            
            If lResult <> 0 Then
                
                'If error in updating, show error window
                
                gcolErrMsg.Add lResult
                giErrMsg = ShowErrorMsg("ProLbl002", _
                                        Me.Caption, _
                                        vbOKOnly, _
                                        gcolErrMsg)

                UpdateSecondaryCustomers = False
                GoTo CleanUpAndExit
                
            ElseIf bDuplLblCode Then

                'Show error message in case of duplicate label code / prod
                'type
                
                giErrMsg = ShowErrorMsg("ProLbl014", Me.Caption, vbOKOnly)

                UpdateSecondaryCustomers = False
                txtLabelNo.SelStart = 0
                txtLabelNo.SelLength = Len(Trim(txtLabelNo.Text))
                txtLabelNo.SetFocus
                GoTo CleanUpAndExit
                
            End If
            
       End If

' addition ends here
            
            
' now add/update ot_frz_duration
            
            
            'Call the update method of RateCodeMaster to update the db
'
'            If Not colSecCustomers Is Nothing Then
'
'                'Create object of ProductMaster Class in MasterFileFunctions
'
''                Set objProduct = CreateObject( _
''                                "MasterFileFunctions.ProductMaster")
'
'                'Update Secondary Customer Details
'
'                    lResult = objProduct.UpdateSecCustomersForProductPlant _
'                                                (gsPlantCode, colSecCustomers)
'
'
'                If Not colSecCustomers Is Nothing Then
'
'                    'If error occurs in updating show error window
'
'                    DisplayErrors ustgrdSecondaryCustomers, _
'                                  cnAction, COLUMN_COUNT, _
'                                  colSecCustomers, cnCustNumber, cnCustType
'
'                    gcolErrMsg.Add lResult
'                    lResult = ShowErrorMsg("Second008", Me.Caption, vbOKOnly, gcolErrMsg)
'                    Set fcolErrors = colSecCustomers
'                    PopulateErrorWindow fcolErrors, True
'
'                    UpdateSecondaryCustomers = False
'                    GoTo CleanUpandExit
'
'                End If
'
'                Set fcolErrors = Nothing
'                CloseErrorWindow
'
'            End If
'
        
        
' added from update detailinventory

        If fcolProduct Is Nothing Then
                
                RaiseEvent RefreshGrid("I", colProductClone)
            
        Else
                
                RaiseEvent RefreshGrid("U", colProductClone)
            
        End If
        
' addition ends here
        
        
Else
        
    UpdateSecondaryCustomers = False
    
End If

CleanUpAndExit:
    
    Set objProduct = Nothing
    Set colSecCustomer = Nothing
    Set colSecCustomers = Nothing
    Set colProduct = Nothing
    Set colProductClone = Nothing
    Set colProducts = Nothing
    
    Exit Function

ErrHandler:
    
    'VB Error
    
    UpdateSecondaryCustomers = False
    gcolErrMsg.Add Err.Description
    giErrMsg = ShowErrorMsg("Second009", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    GoTo CleanUpAndExit

End Function

' Added by TCS

Private Sub ustgrdSecondaryCustomers_CellDataModified()
    
    CellDataModified ustgrdSecondaryCustomers, cnAction, flRowIndex, cnProduct

End Sub

' Added by TCS

Private Sub ustgrdSecondaryCustomers_EnterCell()
    
    ValidateRow
    
    'Added by TCs on 23-Aug-05
    'Start
    'Commented by George on 06-Dec-2005 for HD0000000534698
    cboWeightType.Enabled = True
    'If gsUserGroup <> "102" Then
    '    cboWeightType.Enabled = IIf(ustgrdSecondaryCustomers.CellValue(ustgrdSecondaryCustomers.Row, cnAction) = "I", True, _
    '                            IIf(ustgrdSecondaryCustomers.CellValue(ustgrdSecondaryCustomers.Row, cnAction) = "", True, False))
    'End If
    'End
    
End Sub

' Added by TCS

Private Function ValidateRow(Optional ByVal lRow As Long = -1) As Boolean

Dim rsCustTypes         As ADODB.Recordset
Dim lColumnIndex        As Long
Dim lIndex              As Long
Dim bValidCustomer      As Boolean
Dim lRowIndex           As Long
    
    If lRow = -1 Then
    
        lRowIndex = flRowIndex
    Else
    
        lRowIndex = lRow
        
    End If

    ValidateRow = True
    
    With ustgrdSecondaryCustomers
        
        'If the previous RowIndex is a valid one, then validate it
        
        If lRowIndex <= .Rows - 1 Then
        
            'Before setting the next column, validate previous column
            'and set next column acccordingly
            
            lColumnIndex = flColumnIndex
            
            If OriginPlantCodeIsValid(.CellValue(flRowIndex, cnOriginPlant)) Then
                If DuplicateInGrid(ustgrdSecondaryCustomers, _
                        .CellValue(flRowIndex, cnOriginPlant), _
                            cnOriginPlant, flRowIndex) Then
                        
                    'Call MsgBox("Origin Plant code already exists", vbOKOnly, Me.Caption)
                    'Call ShowErrorMsg("RateMstr008", Me.Caption, vbOKOnly)
                    Call ShowErrorMsg("Second013", Me.Caption, vbOKOnly)
                    lColumnIndex = cnOriginPlant
                    .SetFocus
                    ustgrdSecondaryCustomers.GoToXY CInt(lRowIndex), cnOriginPlant  'Added by TCS on 22-Aug-05
                    GoTo CleanUpAndExit
                End If
            Else
                lColumnIndex = cnOriginPlant
                If ustgrdSecondaryCustomers.Column <> cnOriginPlant Then
                    .SetFocus
                    ustgrdSecondaryCustomers.GoToXY CInt(flRowIndex), cnOriginPlant 'Added by TCS on 22-Aug-05
                End If
                GoTo CleanUpAndExit
            End If
        
            If .CellValue(lRowIndex, cnCustNumber) <> "" Then
                
                If Not CustomerCodeIsValid(.CellValue( _
                        lRowIndex, cnCustNumber), _
                        rsCustTypes, _
                        bValidCustomer, _
                        Me.Caption) Then
                    
                    .SetFocus
                    ustgrdSecondaryCustomers.GoToXY CInt(lRowIndex), cnCustNumber   'Added by TCS on 22-Aug-05
                    GoTo CleanUpAndExit
                
                ElseIf Not bValidCustomer Then
                    
                    giErrMsg = ShowErrorMsg("Second002", Me.Caption, vbOKOnly)
                    .SetFocus
                    ustgrdSecondaryCustomers.GoToXY CInt(lRowIndex), cnCustNumber   'Added by TCS on 22-Aug-05
                    GoTo CleanUpAndExit
                    
                Else
                    
                    If DuplicateCustInGrid(.CellValue(lRowIndex, cnCustNumber), _
                                           .CellValue(lRowIndex, cnCustType), _
                                           lRowIndex) Then
                                           
                        Call ShowErrorMsg("Second014", Me.Caption, vbOKOnly)
                        lColumnIndex = cnCustNumber
                        .SetFocus
                        ustgrdSecondaryCustomers.GoToXY CInt(lRowIndex), cnCustNumber   'Added by TCS on 22-Aug-05
                        GoTo CleanUpAndExit
                        
                    End If
                    
                    If rsCustTypes.EOF Then
                        
                        rsCustTypes.MoveFirst
                    
                    End If
                    
                    lIndex = 0
                    ReDim fsArrCustType(lIndex)
                    
                    While Not rsCustTypes.EOF
                        
                        ReDim Preserve fsArrCustType(lIndex + 1)
                        fsArrCustType(lIndex) = _
                        rsCustTypes.Fields("CUSTOMER_TYPE")
                        
                        lIndex = lIndex + 1
                        rsCustTypes.MoveNext
                    
                    Wend
                    
                    .AddComboData cnCustType, fsArrCustType
                    SetCustTypeData ustgrdSecondaryCustomers, fsArrCustType, _
                                    cnCustType, lRowIndex
                
                End If
            
            End If
         
            'Before setting the current row, check if the previous row is valid
            
            If fbUnload Then

                'validate the not null fields
             
                If Trim(.CellValue(lRowIndex, cnAction)) = TO_INSERT Or _
                    Trim(.CellValue(lRowIndex, cnAction)) = TO_UPDATE Or _
                    Trim(.CellValue(lRowIndex, cnAction)) = NEW_ROW Then
                    
                        If Trim(.CellValue(lRowIndex, cnCustNumber)) = "" And _
                            Trim(.CellValue(lRowIndex, cnCustType)) = "" And _
                            Trim(.CellValue(lRowIndex, cnOriginPlant)) = "" And _
                            Trim(.CellValue(lRowIndex, cnFreezerDays)) = "" And _
                            Trim(.CellValue(lRowIndex, cnBlastDays)) = "" Then
                            
                            flRowIndex = .Row
                            flColumnIndex = .Column

                            Exit Function

                        End If

                    If Not ValidateRowForBlank( _
                            ustgrdSecondaryCustomers, lRowIndex, _
                            COLUMN_COUNT, lColumnIndex) Then
                    
                    
                        Select Case lColumnIndex
                        
                            Case cnOriginPlant
                                giErrMsg = ShowErrorMsg("Second021", _
                                                        Me.Caption, _
                                                        vbOKOnly)
                            Case cnCustNumber
                                giErrMsg = ShowErrorMsg("Second004", _
                                                        Me.Caption, _
                                                        vbOKOnly)
                            Case cnCustType
                                giErrMsg = ShowErrorMsg("Second005", _
                                                        Me.Caption, _
                                                        vbOKOnly)
                            Case cnFreezerDays
                                giErrMsg = ShowErrorMsg("Second006", _
                                                        Me.Caption, _
                                                        vbOKOnly)
                            Case cnBlastDays
                                giErrMsg = ShowErrorMsg("Second007", _
                                                        Me.Caption, _
                                                        vbOKOnly)
                        End Select

                        .SetFocus
                        GoTo CleanUpAndExit
                        
                    End If
                End If
            End If
        End If
        
        flRowIndex = .Row
        flColumnIndex = .Column
        
    End With
   
    Exit Function
    
CleanUpAndExit:

    ValidateRow = False
    Set rsCustTypes = Nothing
    ustgrdSecondaryCustomers.Row = lRowIndex
    ustgrdSecondaryCustomers.Column = ustgrdSecondaryCustomers.Column   'Changed "lColumnIndex" to "ustgrdSecondaryCustomers.Column" by TCS on 22-Aug-05
    ustgrdSecondaryCustomers.TopRow = lRowIndex

End Function

' Added by TCS
Private Sub ustgrdSecondaryCustomers_KeyDown(KeyCode As Integer, Shift As Integer)

Static lRowIndex    As Long
Dim lInvalidColumn  As Long


fsProduct = Trim(txtProdCode.Text)
    
         With ustgrdSecondaryCustomers
            Select Case KeyCode
                
'                Case vbKeyDelete
'
'                    ProcessDeleteRow ustgrdSecondaryCustomers, cnAction, _
'                                cnProduct
                                       
                Case vbKeyDown
                    
                    'If focus is on last row, add a new row if needed
                    
                    If .GridEditMode = Cell And _
                    lRowIndex = flRowIndex And _
                    flRowIndex = .Rows - 1 Then
                        
                        If ValidateRowForBlank(ustgrdSecondaryCustomers, _
                                               flRowIndex, _
                                               COLUMN_COUNT, _
                                               lInvalidColumn) Then
                            
                            HideDeletedDescription ustgrdSecondaryCustomers, _
                                                   cnAction, cnProduct
                            
                            SetLastRow ustgrdSecondaryCustomers, flRowIndex, cnAction
                            ustgrdSecondaryCustomers_EnterCell
                            .CellValue(flRowIndex, cnProduct) = fsProduct
                            .CellValue(flRowIndex, cnDefaultOrigin) = "N"   'Added by TCS on Aug-22-2005
                            
                            .Column = cnOriginPlant
                            flColumnIndex = cnOriginPlant
                       
                        End If
                    
                    End If
                    
                    ustgrdSecondaryCustomers_Click
                Case vbKeyUp
    
                    'If focus is on last last row, remove last row
                    'if it is blank
                    
                     If .GridEditMode = Cell And _
                     flRowIndex = .Rows - 2 Then
                        
                        If RowIsBlank(.Rows - 1) Then
                            
                            .Rows = .Rows - 1
                            .Row = .Rows - 1
                            flRowIndex = .Row
                        
                        End If
                    
                    End If
                    ustgrdSecondaryCustomers_Click
                Case vbKeyReturn
                
                    If .Column = cnProduct Then
                        
                        If .CellValue(flRowIndex, cnProduct) = DELETED Then
                            
                            ShowDeletedDescription ustgrdSecondaryCustomers, _
                                                   cnAction, _
                                                   cnProduct
                        
                        End If
                    
                    End If

            End Select
            
            lRowIndex = flRowIndex
        
        End With
     
End Sub

' Added by TCS
Private Sub ustgrdSecondaryCustomers_KeyPress(KeyAscii As Integer)
    
Dim lColumnIndex As Long

    'If fsAccessRights = FUNC_MODIFY Then
        
        With ustgrdSecondaryCustomers
            
            If KeyAscii = vbKeyEscape Then
                
                'If Escape is pressed from the last row which is an inserted row,
                'that row is cleared if grid has > 2 rows and the inserted row is not valid
                'even if it is in edit mode
                
                If (Left(.CellValue(flRowIndex, cnAction), 1) = NEW_ROW Or _
                Left(.CellValue(flRowIndex, cnAction), 1) = TO_INSERT) And _
                (flRowIndex = .Rows - 1) And (.Rows > 2) And _
                (Not ValidateRowForBlank(ustgrdSecondaryCustomers, flRowIndex, _
                COLUMN_COUNT, lColumnIndex)) Then
                    
                    .Rows = .Rows - 1
                    .Row = .Rows - 1
                    flRowIndex = .Row
                    
                ElseIf (.GridEditMode = Cell) Then
                    
                    Unload Me
                
                End If
                
                Exit Sub
            
            End If
            
            Select Case .Column
                
                Case cnFreezerDays, cnBlastDays
                   
                    If KeyAscii >= 97 And KeyAscii <= 122 Then
                        
                        KeyAscii = KeyAscii - 32
                    
                    End If
                       
            End Select
            
        End With
    
End Sub

' Added by TCS
'*******************************************************************************
'* Functional Description   :   Displayes a row as deleted if its marked for
'*                              deletion
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub ustgrdSecondaryCustomers_LeaveCell()
    
    With ustgrdSecondaryCustomers
        
        If .Column = cnProduct Then 'here (changed by TCS)
            
            HideDeletedDescription ustgrdSecondaryCustomers, cnAction, cnProduct
        
        End If
    
    End With
    
End Sub

' Added by TCS
'*******************************************************************************
'* Functional Description   :   Handles Rightclick menu
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub ustgrdSecondaryCustomers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'If fsAccessRights = FUNC_MODIFY Then
        
        With ustgrdSecondaryCustomers
            
            If Button = vbRightButton And .GridEditMode = Cell Then

                If Not fcolErrors Is Nothing Then

                    mnuError.Visible = True

                Else: mnuError.Visible = False

                End If

                PopupMenu mnuPopup
            End If

        End With
    
    'End If

End Sub

' Added by TCS

'******************************************************************************
'* Functional Description   :   Add a new row when tabbed from last column of
'*                              last row
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************
Private Sub ustgrdSecondaryCustomers_MouseUp(Button As Integer, _
                                        Shift As Integer, _
                                        X As Single, _
                                        Y As Single)
    
    ustgrdSecondaryCustomers_KeyDown vbKeyLeft, 0
    
End Sub

' Added by TCS

Private Function DuplicateCustInGrid(ByVal sValueID As String, _
                                ByVal sValueType As String, _
                                Optional ByVal lExcludeRow As Long = -1) _
                                As Boolean

Dim lRowIndex       As Long

    On Error GoTo ErrHandler
    
    DuplicateCustInGrid = False
    With ustgrdSecondaryCustomers
    
        For lRowIndex = 1 To .Rows - 1
        
            If .CellValue(lRowIndex, cnCustNumber) = sValueID And _
               .CellValue(lRowIndex, cnCustType) = sValueType And _
               lRowIndex <> lExcludeRow Then
               
                DuplicateCustInGrid = True
                Exit For
                
            End If
            
        Next
        
    End With

CleanUpAndExit:

    Exit Function

ErrHandler:

    GoTo CleanUpAndExit

End Function

' Added by TCS

Private Function RowIsBlank(ByVal lRowToValidate As Long) As Boolean

    With ustgrdSecondaryCustomers
        
        If .CellValue(lRowToValidate, cnCustNumber) = "" And _
            .CellValue(lRowToValidate, cnCustType) = "" And _
            .CellValue(lRowToValidate, cnBlastDays) = "" And _
            .CellValue(lRowToValidate, cnOriginPlant) = "" And _
            .CellValue(lRowToValidate, cnFreezerDays) = "" Then
        
            RowIsBlank = True
            Exit Function
            
        End If
        
        RowIsBlank = False
    
    End With
    
End Function

' Added by TCS

Private Function LoadOrginPlantCombo() As Long

Dim rsPlants            As Recordset
Dim iRecordCount        As Integer
Dim sArrOrgPlant()     As String
Dim lResult             As Long

    On Error GoTo ErrHandler
    
    Set fdictOrgPlant = New Dictionary
    LoadOrginPlantCombo = 0
    lResult = 0
    
    Set rsPlants = New Recordset
    lResult = LoadOriginPlants(rsPlants)
    
    If lResult <> 0 Then GoTo ErrHandler
    
    If Not rsPlants.EOF Then
        
        rsPlants.MoveFirst
        iRecordCount = 0
        
        While Not rsPlants.EOF
            
            ReDim Preserve sArrOrgPlant(iRecordCount + 1)
            sArrOrgPlant(iRecordCount) = _
                            rsPlants.Fields("COMPLEX_CODE").Value & " - " & _
                            rsPlants.Fields("PLANT_CODE").Value
            iRecordCount = iRecordCount + 1
            
            fdictOrgPlant.Add rsPlants.Fields("PLANT_CODE").Value, _
                            rsPlants.Fields("COMPLEX_CODE").Value
'            fdictOrgPlant.Add rsPlants.Fields("PLANT_CODE").Value, _
                            rsPlants.Fields("COMPLEX_CODE").Value & " - " & _
                            rsPlants.Fields("PLANT_CODE").Value
                            
            rsPlants.MoveNext
        
        Wend
    
    End If
    
    ustgrdSecondaryCustomers.AddComboData cnOriginPlant, sArrOrgPlant
    
CleanUpAndExit:
    
    iRecordCount = 0
    Exit Function
    
ErrHandler:
    
    'Error Handling
    
    If lResult <> 0 Then
        gcolErrMsg.Add lResult
        giErrMsg = ShowErrorMsg("Second017", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    Else
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("Second018", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    End If
    
    LoadOrginPlantCombo = lResult
    GoTo CleanUpAndExit
    
End Function

' Added by TCS

Public Function LoadOriginPlants(ByRef rsPlants As ADODB.Recordset) As Long

Dim objPlant    As Object
Dim lResult     As Long

    On Error GoTo ErrHandler
    
    LoadOriginPlants = 0
    Set objPlant = CreateObject("MasterFileFunctions.PlantMaster")
    lResult = objPlant.GetOriginPlantCodes(rsPlants)
    LoadOriginPlants = lResult
    If LoadOriginPlants <> 0 Then GoTo ErrHandler
    
CleanUpAndExit:
    Set objPlant = Nothing
    Exit Function

ErrHandler:
    
    If LoadOriginPlants <> 0 Then
        
        'Server side error
        
        gcolErrMsg.Add lResult
        giErrMsg = ShowErrorMsg("Second015", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    Else
        
        'VB error
        
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("Second016", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    End If
    GoTo CleanUpAndExit

End Function

' Added by TCS

Private Function OriginPlantCodeIsValid(ByVal sPlantcode As String) _
                                    As Boolean
    
    On Error GoTo ErrHandler
    
    If Trim(sPlantcode) = "" Then
    
        OriginPlantCodeIsValid = True
        Exit Function
    
    End If
    
    If fdictOrgPlant.Exists(sPlantcode) Then
    
        OriginPlantCodeIsValid = True
    
    Else
        
        Call ShowErrorMsg("Second019", Me.Caption, vbOKOnly)
        OriginPlantCodeIsValid = False
        
    End If
    
    Exit Function

ErrHandler:
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "Second020", Me.Caption, vbOKOnly, gcolErrMsg
    
    OriginPlantCodeIsValid = False
    
End Function

' Added by TCS

Public Sub SetCustTypeData(ByRef ustgrdGrid As USTriSuperGrid.USTriGrid, _
                           ByRef sArrCustType() As String, _
                           ByVal lCustTypeColumnIndex As Long, _
                           ByVal lRowIndex As Long)
    
Dim lIndex As Long

    With ustgrdGrid
    
        For lIndex = 0 To UBound(sArrCustType)
        
            If sArrCustType(lIndex) = .CellValue(lRowIndex, _
                                                 lCustTypeColumnIndex) Then
                Exit Sub
                
            End If
            
        Next
        
        .CellValue(lRowIndex, lCustTypeColumnIndex) = ""
        
    End With
    
End Sub

' Added by TCS

Private Sub ustgrdSecondaryCustomers_Click()
Dim vrCurrentRow As Integer, vrCurrentCol As Integer
Dim vrRowHeader As Integer

    With ustgrdSecondaryCustomers
    
    'Variables used to identify Row & Col
    vrRowHeader = 0
    bAllowSelChangeEvent = False
'MsgBox fcolProduct("GOVT_LOT_IND_1") & " ; " & fcolProduct("GOVT_LOT_IND_2")

    vrCurrentRow = .Row
    vrCurrentCol = .Column
    
    If .Row = PrevRowSelected Or (Trim(.TextMatrix(.Row, 0)) & Trim(.TextMatrix(.Row, 2)) = "") Then bAllowSelChangeEvent = True: Exit Sub
    If Trim(.TextMatrix(.Row, 2)) = "" Then bAllowSelChangeEvent = True: Exit Sub

    'Unselect previously Selected Row
    
    If PrevRowSelected > 0 And PrevRowSelected < .Rows Then
       .Row = PrevRowSelected
       'Change the Cells backcolor to white & Forecolor to blue
       ChangeUstrGridBackAndForeColor False
       .Row = vrCurrentRow
        lgrdRowIndex = PrevRowSelected
       UpdateRecord
       'SetInitialValues    'Commented by TCS on 03-Sep-2005 for Copying all values to the new record
    End If
    'Change the Cells Backcolor to blue & Forecolor to white
    ChangeUstrGridBackAndForeColor True
    
    'Set Selected column
    .Column = vrCurrentCol
    
    PrevRowSelected = .Row
    
    lgrdRowIndex = .Row

    ' Show plant  details
    'MsgBox fcolProduct("GOVT_LOT_IND_1") & " ; " & fcolProduct("GOVT_LOT_IND_2")
    
    'Condition Added by TCS-Ragu on 20-jun-2005
    'Start
    On Error Resume Next
    If ItemExistsInCollection(fcolProduct, "PRODUCT_CODE" & "_" & Trim(Str(lgrdRowIndex)), "") Then
        If fcolProduct.Item("PRODUCT_CODE" & "_" & Trim(Str(lgrdRowIndex))) <> "" Then
            DisplayRecord
        End If
    End If
    'End
    
    'MsgBox fcolProduct("GOVT_LOT_IND_1") & " ; " & fcolProduct("GOVT_LOT_IND_2")

    bAllowSelChangeEvent = True
    
    
        
        If .Column = cnProduct Then

            If .CellValue(flRowIndex, cnProduct) = DELETED Then

                ShowDeletedDescription ustgrdSecondaryCustomers, _
                                       cnAction, _
                                       cnProduct

            End If

        End If
    
    End With

End Sub

' Added by TCS

Private Sub mnuDelete_Click()

    ProcessDeleteRow ustgrdSecondaryCustomers, cnAction, cnProduct

End Sub


''******************************************************************************
'* Functional Description   :   To highlight Selected Row
'* Parameter Description    :   bChange - Boolean Value to change Backcolor
'*                              and Forecolor of Partialpallet Grid
'* Return Type Description  :   None.
'******************************************************************************

Private Function ChangeUstrGridBackAndForeColor(bChange As Boolean)

'Added by TCS
Dim I As Integer

With ustgrdSecondaryCustomers

    For I = 0 To .Columns - 1
    .Column = I
    .CellBackColor = IIf(bChange, vbBlue, vbWhite)
    .CellForeColor = IIf(bChange, vbWhite, vbBlack)
    Next
    .Column = 0

End With
End Function



'******************************************************************************
'* Functional Description   :   Populate the grid with product details
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Function LoadProductDetails() As Long

Dim objProduct      As Object
Dim rsProducts      As ADODB.Recordset
Dim lRowIndex       As Long
Dim lColIndex       As Long
Dim SFilter         As String
    On Error GoTo ErrHandler
    
    LoadProductDetails = 0
    
    FormatGrid
    
    
Set rsProducts = New ADODB.Recordset
    
    'Create object of ProductMaster Class in MasterFileFunctions
    
    Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
    
    'Retrieve Product Details
    
    LoadProductDetails = objProduct.ProdMstrPlantDetails(gsPlantCode, Trim(txtProdCode.Text), _
                                                              rsProducts)
    
    
    
    
    
    
    
    If LoadProductDetails <> 0 Then GoTo ErrHandler
    
    If rsProducts Is Nothing Then
        
        GoTo CleanUpAndExit
    
    Else
    
        Set fcolProduct = New Collection
        lRowIndex = 1
        
       
        While Not rsProducts.EOF
            For lColIndex = 0 To rsProducts.Fields.Count - 1
                fcolProduct.Add _
                        IIf(IsNull(rsProducts(lColIndex).Value), "", rsProducts(lColIndex).Value), _
                        rsProducts.Fields(lColIndex).Name & "_" & Trim(Str(lRowIndex))
            Next
        
            lRowIndex = lRowIndex + 1
              
            rsProducts.MoveNext
        Wend
         fcolProduct.Add lRowIndex - 1, "RowNum"
    End If
        
    
    
    
    
CleanUpAndExit:

    Set objProduct = Nothing
    Set rsProducts = Nothing
    Exit Function
    
ErrHandler:
   
    If LoadProductDetails <> 0 Then
        
        'Display Server error
        
        gcolErrMsg.Add LoadProductDetails
        giErrMsg = ShowErrorMsg("ProMstr002", _
                                Me.Caption, _
                                vbOKOnly, _
                                gcolErrMsg)
   Else
        
        'Display VB error
        
        LoadProductDetails = Err.Number
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("ProMstr012", _
                                Me.Caption, _
                                vbOKOnly, _
                                gcolErrMsg)
   End If
     
    
    GoTo CleanUpAndExit
    
End Function



'******************************************************************************
'* Functional Description   :   Loads details
'* Parameter Description    :   None
'* Return Type Description  :   Success of loading
'******************************************************************************

Private Function UpdateRecord() As Boolean
'If fAddEditMode = "" Then Exit Function
Dim lRowIdx As Long  'Added by TCS-Ragu on 13-10-05

On Error Resume Next ' Test case
    
    With fcolProduct
        
        'Added by TCS-Ragu on 13-10-05
        'Start
        For lRowIdx = 1 To ustgrdSecondaryCustomers.Rows - 1
            .Remove "PRODUCT_DESC" & "_" & Trim(Str(lRowIdx))
            .Add txtDescription.Text, ("PRODUCT_DESC" & "_" & Trim(Str(lRowIdx)))
                
            .Remove "PRODUCT_TYPE" & "_" & Trim(Str(lRowIdx))
            .Add cboProductType.Text, "PRODUCT_TYPE" & "_" & Trim(Str(lRowIdx))
                
            .Remove "LABEL_NO" & "_" & Trim(Str(lRowIdx))
            .Add txtLabelNo.Text, "LABEL_NO" & "_" & Trim(Str(lRowIdx))
        Next
        'End
        
        .Remove ("PRODUCT_CODE" & "_" & Trim(Str(lgrdRowIndex)))
        .Add txtProdCode.Text, "PRODUCT_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        'Commented by TCS-Ragu on 13-10-05
        '.Remove "PRODUCT_DESC" & "_" & Trim(Str(lgrdRowIndex))
        '.Add txtDescription.Text, ("PRODUCT_DESC" & "_" & Trim(Str(lgrdRowIndex)))
         
        .Remove "DIVISION_CODE" & "_" & Trim(Str(lgrdRowIndex))
        .Add lblDivisonCodeValue.Caption, "DIVISION_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        'Commented by TCS-Ragu on 13-10-05
        '.Remove "PRODUCT_TYPE" & "_" & Trim(Str(lgrdRowIndex))
        '.Add cboProductType.Text, "PRODUCT_TYPE" & "_" & Trim(Str(lgrdRowIndex))
        
        '.Remove "LABEL_NO" & "_" & Trim(Str(lgrdRowIndex))
        '.Add txtLabelNo.Text, "LABEL_NO" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "WGT_TYPE_CODE" & "_" & Trim(Str(lgrdRowIndex))
        .Add cboWeightType.Text, "WGT_TYPE_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Remove "BOXES_PER_PALLET" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtBoxesPallet.Text, "BOXES_PER_PALLET" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "MIN_BOX_WGT" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtMinWeight.Text, "MIN_BOX_WGT" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "MAX_BOX_WGT" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtMaxWeight.Text, "MAX_BOX_WGT" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "BOX_TARE_WEIGHT" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtTareWeight.Text, "BOX_TARE_WEIGHT" & "_" & Trim(Str(lgrdRowIndex))
              
        .Remove "LABEL_LENGTH" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtLabelLength.Text, "LABEL_LENGTH" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "LABEL_WGT_ST_POS" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtStrPosition.Text, "LABEL_WGT_ST_POS" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "LABEL_WGT_LENGTH" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtWeightLength.Text, "LABEL_WGT_LENGTH" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "GOVT_COMMODITY_CODE" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtCommodityCode.Text, "GOVT_COMMODITY_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "GOVT_LOT_IND" & "_" & Trim(Str(lgrdRowIndex))
        .Add cboGovtLot.Text, "GOVT_LOT_IND" & "_" & Trim(Str(lgrdRowIndex))

        .Remove "PACK_DATE_IND" & "_" & Trim(Str(lgrdRowIndex))
        .Add cboMfgDate.Text, "PACK_DATE_IND" & "_" & Trim(Str(lgrdRowIndex))
        
        'Added by TCS-Ragu on 17-Jun-2005
        .Remove "PACK_DATE_START" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtPckdtStart.Text, "PACK_DATE_START" & "_" & Trim(Str(lgrdRowIndex))

        .Remove "CHECK_LABEL_IND" & "_" & Trim(Str(lgrdRowIndex))
        .Add cboLblCodeChk.Text, "CHECK_LABEL_IND" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "PROD_LABEL_START" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtLabelCodeStr.Text, "PROD_LABEL_START" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "PROD_LABEL_LEN" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtLabelCodeLength.Text, "PROD_LABEL_LEN" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "BOX_SERIAL_IND" & "_" & Trim(Str(lgrdRowIndex))
        .Add cboBoxSerialInd.Text, "BOX_SERIAL_IND" & "_" & Trim(Str(lgrdRowIndex))
            
        .Remove "XREF_CODE" & "_" & Trim(Str(lgrdRowIndex))
        .Add txtRefCode.Text, "XREF_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        .Remove "PRODUCT_GROUP_CODE" & "_" & Trim(Str(lgrdRowIndex))
        .Add cboPrdgrpcode.Text, "PRODUCT_GROUP_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
       
        
    End With
    
End Function





Private Function NewRecord() As Boolean
'If fAddEditMode = "" Then Exit Function
On Error Resume Next ' Test case
    
    With fcolProduct
        
        
        .Add txtProdCode.Text, "PRODUCT_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        
        
        .Add txtDescription.Text, ("PRODUCT_DESC" & "_" & Trim(Str(lgrdRowIndex)))
         
        
        .Add lblDivisonCodeValue.Caption, "DIVISION_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        
        
        .Add cboProductType.Text, "PRODUCT_TYPE" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtLabelNo.Text, "LABEL_NO" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add cboWeightType.Text, "WGT_TYPE_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        
        
        .Add txtBoxesPallet.Text, "BOXES_PER_PALLET" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtMinWeight.Text, "MIN_BOX_WGT" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtMaxWeight.Text, "MAX_BOX_WGT" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtTareWeight.Text, "BOX_TARE_WEIGHT" & "_" & Trim(Str(lgrdRowIndex))
              
        
        .Add txtLabelLength.Text, "LABEL_LENGTH" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtStrPosition.Text, "LABEL_WGT_ST_POS" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtWeightLength.Text, "LABEL_WGT_LENGTH" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtCommodityCode.Text, "GOVT_COMMODITY_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add cboGovtLot.Text, "GOVT_LOT_IND" & "_" & Trim(Str(lgrdRowIndex))

        
        .Add cboMfgDate.Text, "PACK_DATE_IND" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add cboLblCodeChk.Text, "CHECK_LABEL_IND" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtLabelCodeStr.Text, "PROD_LABEL_START" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add txtLabelCodeLength.Text, "PROD_LABEL_LEN" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add cboBoxSerialInd.Text, "BOX_SERIAL_IND" & "_" & Trim(Str(lgrdRowIndex))
            
        
        .Add txtRefCode.Text, "XREF_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
        
        .Add cboPrdgrpcode.Text, "PRODUCT_GROUP_CODE" & "_" & Trim(Str(lgrdRowIndex))
        
       
        
    End With
    
End Function




Private Sub ustgrdSecondaryCustomers_Validate(Cancel As Boolean)
    Static sOrigValue       As String
    
    With ustgrdSecondaryCustomers
        If .Column = cnDefaultOrigin Then

            ReturnValueForCapsAndNumerics txtProdCode, sOrigValue
            txtProdCode = sOrigValue
                
            ustgrdSecondaryCustomers.CellValue(1, 0) = Trim(txtProdCode.Text)
        
        End If
    
    End With


End Sub
