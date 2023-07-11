VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{82CA1823-AE1A-4850-BDC7-F24C9A05E6D0}#1.1#0"; "UstriGrid.ocx"
Begin VB.Form mdi_DailyProduction 
   Caption         =   "Daily Production Browse / Update"
   ClientHeight    =   5850
   ClientLeft      =   2295
   ClientTop       =   2580
   ClientWidth     =   8190
   Icon            =   "mdi_DailyProduction.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   8190
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inquire"
            Description     =   "Inquire"
            Object.ToolTipText     =   "Inquire (F2)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Prod Break"
            Description     =   "Prod Break"
            Object.ToolTipText     =   "Prod. Breakout (F4)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Description     =   "Add"
            Object.ToolTipText     =   "Add (F5)"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      Height          =   5055
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   8175
      Begin VB.Frame fraCustomerShift 
         BorderStyle     =   0  'None
         Height          =   280
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   7005
         Begin VB.Label lblDPU 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shift"
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
            Index           =   0
            Left            =   1200
            TabIndex        =   8
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblShift 
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
            Left            =   1800
            TabIndex        =   7
            Top             =   0
            Width           =   45
         End
         Begin VB.Label lblDPU 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            Index           =   1
            Left            =   3480
            TabIndex        =   6
            Top             =   0
            Width           =   825
         End
         Begin VB.Label lblCustomer 
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
            Left            =   4560
            TabIndex        =   5
            Top             =   0
            Width           =   45
         End
      End
      Begin USTriSuperGrid.USTriGrid ustgrdDailyProduction 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7646
         Columns         =   2
         FixedColumns    =   0
         FixedRows       =   0
         Rows            =   1
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
         WordWrap        =   0   'False
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
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5550
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12523
            MinWidth        =   8819
            Text            =   "<F2 - Inquire><F4 - Prod BreakOut><F5 - Add><Esc - Quit>"
            TextSave        =   "<F2 - Inquire><F4 - Prod BreakOut><F5 - Add><Esc - Quit>"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "10:43 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_DailyProduction.frx":1CFA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_DailyProduction.frx":1E0C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_DailyProduction.frx":1F1E
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_DailyProduction.frx":2030
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_DailyProduction.frx":290A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_DailyProduction.frx":31E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save and Close"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Without Save"
      End
      Begin VB.Menu mnuWinClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actio&ns"
      Begin VB.Menu mnuInquire 
         Caption         =   "&Inquire"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuProdBreak 
         Caption         =   "&Prod. Breakout"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuErrors 
         Caption         =   "Show Errors"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Admin"
      Begin VB.Menu mnuSelectPlant 
         Caption         =   "&Select Plant"
      End
   End
   Begin VB.Menu mnuMasterFile 
      Caption         =   "&Master File"
      Begin VB.Menu mnuSecurity 
         Caption         =   "&Security Browse/Update"
      End
      Begin VB.Menu mnuSecurityRF 
         Caption         =   "R&F Security Browse/Update"
      End
      Begin VB.Menu mnuPlantMasterUpdate 
         Caption         =   "&Plant Master Update"
      End
      Begin VB.Menu mnuProductMasterUpdate 
         Caption         =   "Product &Master Browse/Update"
      End
      Begin VB.Menu mnuCustomerMasterUpdate 
         Caption         =   "&Customer Master Browse/Update"
      End
      Begin VB.Menu mnuRateCodeMasterUpdate 
         Caption         =   "&Rate Code Master Browse/Update"
      End
      Begin VB.Menu mnuCommoditiesUpdate 
         Caption         =   "C&ommodities Browse/Update"
      End
      Begin VB.Menu mnuMemoMasterUpdate 
         Caption         =   "Memo Mast&er Browse/Update"
      End
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "&Inventory"
      Begin VB.Menu mnuDetailInventoryUpdate 
         Caption         =   "&Detail Inventory Browse/Update"
      End
      Begin VB.Menu mnuDetailInventoryAdd 
         Caption         =   "Detail &Inventory Add"
      End
      Begin VB.Menu mnuInvProdStatusUpdate 
         Caption         =   "Product Status &Update"
      End
      Begin VB.Menu mnuPalletMovement 
         Caption         =   "&Pallet Movement/Relocation"
      End
      Begin VB.Menu mnuDailyProductionUpdate 
         Caption         =   "D&aily Production Browse/Update"
      End
      Begin VB.Menu mnuFreezerInventoryBrowse 
         Caption         =   "&Freezer Inventory Browse"
      End
      Begin VB.Menu mnuEndShiftCycle 
         Caption         =   "End of &Shift Cycle"
      End
      Begin VB.Menu mnuEndDayCycle 
         Caption         =   "E&nd of Day Cycle"
      End
   End
   Begin VB.Menu mnuCustomerStorage 
      Caption         =   "&Customer Storage"
      Begin VB.Menu mnuCustomerStorageBrowse 
         Caption         =   "Customer Storage &Browse"
      End
      Begin VB.Menu mnuCustomerStorageUpdate 
         Caption         =   "Customer Storage &Update"
      End
      Begin VB.Menu mnuCustomerChargesUpdate 
         Caption         =   "Customer C&harges Browse/Update"
      End
   End
   Begin VB.Menu mnuLoadout 
      Caption         =   "&LoadOut"
      Begin VB.Menu mnuOrderInformationUpdate 
         Caption         =   "Order &Information Browse/Update"
      End
      Begin VB.Menu mnuOrderSummaryBrowse 
         Caption         =   "Order Summary &Browse"
      End
      Begin VB.Menu mnuOrderShippingUpdate 
         Caption         =   "Order &Shipping Update"
      End
      Begin VB.Menu mnuOrderHistoryBrowse 
         Caption         =   "Order &History Browse"
      End
      Begin VB.Menu mnuExportOrderVerification 
         Caption         =   "Order &Verification"
      End
   End
   Begin VB.Menu mnuCommunication 
      Caption         =   "C&ommunications"
      Begin VB.Menu mnuProcessInboundReceipts 
         Caption         =   "Process Inbound &Receipts"
      End
      Begin VB.Menu mnuProcessInboundASNs 
         Caption         =   "Process Inbound &ASNs"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuProcessWIPReceipts 
         Caption         =   "Process &WIP Receipts"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInitProdSend 
         Caption         =   "&Initiate Production Send"
      End
      Begin VB.Menu mnuToBatchScanner 
         Caption         =   "&Send Product Inforamtion To Batch Scanner"
      End
      Begin VB.Menu mnuFromBatchScanner 
         Caption         =   "Receive &Pallet Information From Batch Scanner"
      End
   End
   Begin VB.Menu mnuAdjustmentLog 
      Caption         =   "A&djustment Log"
      Begin VB.Menu mnuAdjustmentLogBrowse 
         Caption         =   "Adjustment &Log Browse"
      End
   End
   Begin VB.Menu mnuProductivity 
      Caption         =   "&Productivity"
      Begin VB.Menu mnuProductivityTracking 
         Caption         =   "&Productivity Tracking"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuInventoryReports 
         Caption         =   "&Inventory"
         Begin VB.Menu mnuBlastCellUnloadingReport 
            Caption         =   "&Blast Cell Unloading Report"
         End
         Begin VB.Menu mnuDailyProductionReport 
            Caption         =   "&Daily Production Report"
         End
         Begin VB.Menu mnuDetailProductionReport 
            Caption         =   "D&etail Production Report"
         End
         Begin VB.Menu mnuCodebyCodeReport 
            Caption         =   "&Code by Code Report"
         End
         Begin VB.Menu mnuAgedProductInventoryReport 
            Caption         =   "&Aged Product Inventory Report"
         End
         Begin VB.Menu mnuFreezerInventoryReport 
            Caption         =   "&Freezer Inventory Report"
         End
         Begin VB.Menu mnuWeeklyFreezerRecapReport 
            Caption         =   "&Weekly Freezer Recap Report"
         End
         Begin VB.Menu mnuColdStorageCommodityReport 
            Caption         =   "Cold &Storage Commodity Report"
         End
         Begin VB.Menu mnuCodeVsInventoryReport 
            Caption         =   "Out of &Balance Report"
         End
         Begin VB.Menu mnuProductsNotSetUp 
            Caption         =   "&Products Not Setup Properly Report"
         End
         Begin VB.Menu mnuSummaryInventoryActivityReport 
            Caption         =   "S&ummary Inventory Activity Report"
         End
         Begin VB.Menu mnuPalletLPNDetailReport 
            Caption         =   "Pallet &LPN Detail Report"
         End
      End
      Begin VB.Menu mnuCustomerStorageReports 
         Caption         =   "&Customer Storage"
         Begin VB.Menu mnuCustomerMasterReport 
            Caption         =   "Customer &Master Report"
         End
         Begin VB.Menu mnuCustomerStorageDueReport 
            Caption         =   "Customer Storage D&ue Report"
         End
         Begin VB.Menu mnuCustomerStorageDetailReport 
            Caption         =   "Customer Storage D&etail Report"
         End
         Begin VB.Menu mnuCustomerChargesReport 
            Caption         =   "Customer C&harges Report"
         End
         Begin VB.Menu mnuAccountReceivableTransmittalRpt 
            Caption         =   "&Account Receivable Transmittal Report"
         End
         Begin VB.Menu mnuWeeklysalesSummaryReport 
            Caption         =   "&Weekly Sales Summary Report"
         End
         Begin VB.Menu mnuCustomerInventoryByProductReport 
            Caption         =   "Customer &Inventory By Product Report"
         End
         Begin VB.Menu mnuCustomerStorageRecapReport 
            Caption         =   "Customer Storage &Recap Report"
         End
         Begin VB.Menu mnuCustomerWarehouseReceipts 
            Caption         =   "Cu&stomer Warehouse Receipts"
         End
      End
      Begin VB.Menu mnuLoadoutReports 
         Caption         =   "&LoadOut"
         Begin VB.Menu mnuOrderManifestDetailReport 
            Caption         =   "Order &Manifest Detail Report"
         End
         Begin VB.Menu mnuOrderManifestSummaryReport 
            Caption         =   "&Order Manifest Summary Report"
         End
         Begin VB.Menu mnuOrderPullSheetReport 
            Caption         =   "Order &Pull Sheet Report"
         End
         Begin VB.Menu mnuOrderSummaryReport 
            Caption         =   "Order &Summary Report"
         End
         Begin VB.Menu mnuProductRecallReport 
            Caption         =   "Product &Recall Report"
         End
         Begin VB.Menu mnuOutsideCustomerBOL 
            Caption         =   "Outside Customer &BOL"
         End
         Begin VB.Menu mnuPalletDetailByOrderReport 
            Caption         =   "Pallet D&etail by Order Report"
         End
         Begin VB.Menu mnuOrderHistorySummaryReport 
            Caption         =   "Order &History Summary Report"
         End
         Begin VB.Menu mnuOrderHistoryManifestReport 
            Caption         =   "Order History Mani&fest Report"
         End
         Begin VB.Menu mnuOrderHistoryManifestSummaryReport 
            Caption         =   "Order History Manifest S&ummary Report"
         End
         Begin VB.Menu mnuOrderListingReport 
            Caption         =   "Order &Listing Report"
         End
         Begin VB.Menu mnuOrderShortageReport 
            Caption         =   "Order Shorta&ge Report"
         End
         Begin VB.Menu mnuCustomerShippingPoundsReport 
            Caption         =   "&Customer Shipping Pounds Report"
         End
         Begin VB.Menu mnuExportOrderUnverifiedCasesReport 
            Caption         =   "E&xport Order Unverified Cases Report"
         End
      End
      Begin VB.Menu mnuProductivityReport 
         Caption         =   "&Productivity"
         Begin VB.Menu mnuInboundOutboundPalletReport 
            Caption         =   "Inbound/Outbound P&allet Report"
         End
      End
      Begin VB.Menu mnuActivity 
         Caption         =   "&Activity"
         Begin VB.Menu mnuActivityReport 
            Caption         =   "&Activity Report"
         End
         Begin VB.Menu mnuBayMovementReport 
            Caption         =   "&Bay Movement Report"
         End
      End
      Begin VB.Menu mnuCorporate 
         Caption         =   "C&orporate"
         Begin VB.Menu mnuGlobalAgedProductInventory 
            Caption         =   "&Global Aged Product Inventory"
         End
         Begin VB.Menu mnuGlobalAgedProductInventoryStatus 
            Caption         =   "Global Aged Product Inventory All  &Status"
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Freezer Inventory System"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuBreakOut 
         Caption         =   "Prod. Breakout"
      End
      Begin VB.Menu mnuAddProduct 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuError 
         Caption         =   "Show Errors"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "mdi_DailyProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : mdi_DailyProduction.frm
'*File Description              : Used to browse/update Daily Production.
'*Author                        : US Technology
'*Date Created                  : Oct-21-02
'*Date Last Modified            : Mar-18-03
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : InventoryFunctions
'*Components Used               : USTriGrid
'*Functions Defined             : 1. FormatGrid
'*                                2. LoadProductionDetails
'*                                3. AddRow
'*                                4. ErrorSelected
'*                                5. RefreshRow
'*                                6. SaveProductionChanges
'*                                7. ValidateRow
'*                                8. GridEdited
'*                                9. LoadDailyProduction
'*                               10. IsValidGridData
'*                               11. SetMainMenu
'*Copyright                     : US Technology
'------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description      Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************

Option Explicit

Private Enum ButtonIndex

    biCut = 1
    biCopy = 2
    biPaste = 3
    biInquire = 5
    biProdBreak = 6
    biAdd = 7
    
End Enum

Private Enum ColumnName

    cnProductCode = 0
    cnBeginWIPBoxes = 1
    cnBeginWIPWeight = 2
    cnProductionBoxes = 3
    cnProductionWeight = 4
    cnEndingWIPBoxes = 5
    cnEndingWIPWeight = 6
    cnTotalBoxes = 7
    cnTotalWeight = 8
    cnAvgWeight = 9
    cnAction = 10
    cnTruckBoxes = 11
    cnTruckWeight = 12

End Enum

Private flRowsRetrieved             As Long
Private flRowIndex                  As Long
Private flColumnIndex               As Long
Private flLoadSuccess               As Long
Private fbCloseWithoutUpdate        As Boolean
Private fbValidated                 As Boolean
Private fcolErrors                  As Collection

Private Const COLUMN_COUNT          As Integer = 13
Private Const CHILD_FORM_WIDTH = 10680

Public fsAccessRights               As String
Public fsShift                      As String
Public fsCustomerId                 As String
Public fsCustomerType               As String
Public fbInquireCancelled           As Boolean
Public fbFormLoaded                 As Boolean

'******************************************************************************
'* Functional Description   :   Checks whether to unload without saving or
'*                              update changes to database and then unload
'* Parameter Description    :   Status of cancel and unload mode
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    On Error GoTo ErrHandler
           
    If Not fbCloseWithoutUpdate Then
    
        If fsAccessRights = FUNC_MODIFY Then
           
            If Not SaveProductionChanges Then
            
                Cancel = 1
                Exit Sub
                
            End If
            
        Else
        
            Unload Me
            
        End If
        
    End If
    
    fbFormLoaded = False
    
CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws012", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
    
End Sub

'**********************************************************************************
'* Functional Description   :   The entry point to this mdi from modMain. Invokes
'                               Form_Load()
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'**********************************************************************************

Public Function LoadDailyProduction() As Boolean
    
    On Error GoTo ErrHandler
    LoadDailyProduction = True
    Load Me
    If flLoadSuccess <> 0 Then GoTo ErrHandler
    Exit Function
    
ErrHandler:
    
    LoadDailyProduction = False
    
End Function

'******************************************************************************
'* Functional Description   :   Unload MDI Window
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Terminate()
    
    Unload Me

End Sub

Private Sub mnuAccountReceivableTransmittalRpt_Click()
'Added by TCS
    AccRecvTransReportClick
    SetNoTopmost
End Sub

Private Sub mnuActivityReport_Click()
    ActivityReportClick
    SetNoTopmost
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Add Product.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuAddProduct_Click()
    
    mnuAdd_Click
    
End Sub

Private Sub mnuBayMovementReport_Click()
    BayMovementReportClick
    SetNoTopmost
End Sub

'******************************************************************************
'* Functional Description   :   Shows the End of the Day cycle form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuEndDayCycle_Click()

    EndDayCycleClick

End Sub

'******************************************************************************
'* Functional Description   :   Shows the End of the Shift cycle form.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuEndShiftCycle_Click()

    EndOfShiftCycleClick

End Sub

Private Sub mnuExportOrderUnverifiedCasesReport_Click()
    ExportOrderUnverifiedCasesReportClick
    SetNoTopmost
End Sub

Private Sub mnuExportOrderVerification_Click()
    ExportOrderVerification
End Sub

Private Sub mnuGlobalAgedProductInventory_Click()
    GlobalAgedProductInventoryReportClick
End Sub

Private Sub mnuGlobalAgedProductInventoryStatus_Click()

GlobalAgedProductInventoryStatusReportClick


End Sub

'******************************************************************************
'* Functional Description   :   Shows the help search screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuHelpSearchForHelpOn_Click()

Dim nRet As Integer
 
    If Len(App.HelpFile) = 0 Then
        ShowErrorMsg "Help001", "Freezer Inventory System", _
                     vbOKOnly
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


'******************************************************************************
'* Functional Description   : Calls Receive Pallet Info from Scanner.
'* Parameter Description    : None.
'* Return Type Description  : None.
'******************************************************************************

Private Sub mnuPalletFromScanner_Click()
    
     Call ReceivePalletInfoFromScanner
     SetNoTopmost
     
End Sub

'******************************************************************************
'* Functional Description   : Calls Send Product Info to Scanner.
'* Parameter Description    : None.
'* Return Type Description  : None.
'******************************************************************************

Private Sub mnuProductToScanner_Click()

    Call SendProductInfoToScanner

End Sub



'Added by TCS on 08-Sep-2005
Private Sub mnuInitProdSend_Click()
    mdi_InitProdSend.Show
End Sub

Private Sub mnuOrderHistoryManifestSummaryReport_Click()

    OrderHistoryManifestSummaryReportClick
    SetNoTopmost
    
End Sub

' Added by TCS
Private Sub mnuProductsNotSetUp_Click()
    
    ProductsNotSetupReportClick
    SetNoTopmost
    
End Sub

'******************************************************************************
'* Functional Description   :   Unloads form after updating database.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuSave_Click()
    
    If Not SaveProductionChanges Then
        Exit Sub
    End If
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Security browse/update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuSecurity_Click()

    SecurityClick

End Sub

'******************************************************************************
'* Functional Description   :   Shows the RF Scanner Security browse/update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuSecurityRF_Click()
    
    SecurityRFClick
    
End Sub

'*******************************************************************************
'* Functional Description   :   Handles Inventory menu item Process WIP
'*                          :   Receipts.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************
Private Sub mnuProcessWIPReceipts_Click()

    ProcessWIPReceiptsClick
    
End Sub

'*******************************************************************************
'* Functional Description   :   Handles Inventory menu item Process Inbound
'*                          :   ASNs.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************
Private Sub mnuProcessInBoundASNs_Click()

    ProcessInboundASNsClick
    
End Sub

'*******************************************************************************
'* Functional Description   :   Handles Inventory menu item Process Inbound
'*                          :   Receipts.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************
Private Sub mnuProcessInBoundReceipts_Click()

    ProcessInboundReceiptsClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets the visibility of the Status bar and
'*                              tool bar and sets the form to standard size.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Activate()

    On Error GoTo ErrHandler
    fbCloseWithoutUpdate = False

    If Not gbCloseErrWindow Then
        Set gfrmCurrentMdi = Me
        PopulateErrorWindow fcolErrors
    End If
    
    SetViewMenu Me
    Form_Resize
    
CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws013", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Hide the Error Window
'* Parameter Description    :   Status of cancel
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo ErrHandler

    Set fcolErrors = Nothing
    CloseErrorWindow
    If (Forms.Count - 2) = 1 Then
        frmMainMenu.sbMainMenu.Visible = True
        frmMainMenu.mnuViewStatusBar.Checked = True
    End If
    
CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws014", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Display error window
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuError_Click()
    
    PopulateErrorWindow fcolErrors

End Sub

'******************************************************************************
'* Functional Description   :   Display error window
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuErrors_Click()
    
    PopulateErrorWindow fcolErrors

End Sub

'******************************************************************************
'* Functional Description   :   The main menu is made invisible and the size
'*                              of the menu is set.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()
   
Dim lResult         As Long

    On Error GoTo ErrHandler
    
    flLoadSuccess = 0
    lResult = 0

    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    
    frmMainMenu.sbMainMenu.Panels(1).Text = _
        "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    
    mnuAdmin.Visible = gbEnableAdminMenu
    mnuExportOrderVerification.Visible = gbEnableExportVerifyMenu
    
    'Added by TCS (To show the Batch Scanner Menu)
    mnuFromBatchScanner.Visible = gbEnableBatchScannerMenu
    mnuToBatchScanner.Visible = gbEnableBatchScannerMenu
    'Addition ends here
    
    mnuInitProdSend.Visible = frmMainMenu.mnuInitProdSend.Visible   'Added by TCS on 08-Sep-2005
    mnuSelectPlant.Enabled = frmMainMenu.mnuSelectPlant.Enabled 'Added by TCS on 08-Sep-2005

    Me.Move 0, 0, frmMainMenu.Width / 1.03, frmMainMenu.Height / 1.15
    
    'Format the Daily Production Grid.
    
    lResult = FormatGrid
    If lResult <> 0 Then GoTo ErrHandler
    
    'Load the Daily Production details for the selected Customer and Shift.
        
    If LoadProductionDetails <> 0 Then GoTo ErrHandler

    SetMainMenu
    Exit Sub

CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:

    fbFormLoaded = False
    flLoadSuccess = lResult
    SetMainMenu
    Form_Terminate
        
End Sub

'******************************************************************************
'* Functional Description   :   Set the properties of the Main Menu
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub SetMainMenu()

    frmMainMenu.sbMainMenu.Visible = False
    frmMainMenu.mnuViewStatusBar.Checked = False
    frmMainMenu.MousePointer = vbDefault
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
 End Sub

'******************************************************************************
'* Functional Description   :   The controls are positioned according to form
'*                              size
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Resize()

Dim lColIndex       As Long

    On Error GoTo ErrHandler
    
    If Me.WindowState = vbMinimized Or _
    frmMainMenu.WindowState = vbMinimized Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Width < CHILD_FORM_WIDTH Then Me.Width = CHILD_FORM_WIDTH
        If Me.Height < CHILD_FORM_HEIGHT Then Me.Height = CHILD_FORM_HEIGHT
    End If
    SetFramePosition Me, fraMain
    
    fraCustomerShift.Left = fraMain.Width / 2 - fraCustomerShift.Width / 2
    
    With ustgrdDailyProduction
        .Move 100, 600, _
                fraMain.Width - 200, fraMain.Height - 700
        For lColIndex = 0 To COLUMN_COUNT - 1
            Select Case lColIndex
                Case cnProductCode
                    .ColWidth(lColIndex) = .Width / 9.5
                Case cnBeginWIPBoxes, cnBeginWIPWeight, cnProductionWeight, _
                     cnProductionBoxes
                    .ColWidth(lColIndex) = .Width / 9.5
                Case cnEndingWIPWeight, cnTotalWeight
                    .ColWidth(lColIndex) = .Width / 9.5
                Case cnAvgWeight, cnTotalBoxes
                    .ColWidth(lColIndex) = .Width / 13
                Case cnEndingWIPBoxes
                    .ColWidth(lColIndex) = .Width / 14
                Case cnTruckBoxes, cnTruckWeight, cnAction
                    .ColWidth(lColIndex) = 0
            End Select
        Next
    
    End With
        
CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws015", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Format the look and feel of the grid.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Function FormatGrid() As Long

Dim lRowCount       As Long
Dim lColIndex       As Long

    On Error GoTo ErrHandler
    
    FormatGrid = 0
    
    With ustgrdDailyProduction
        
        .Redraw = False
        .Rows = 2
        .FixedRows = 1
        .RowHeight(0) = 3 * GRID_ROW_HEIGHT
        .WordWrap = True
        .Columns = COLUMN_COUNT
        
        .CellValue(0, cnProductCode) = "Product Code"
        .CellValue(0, cnBeginWIPBoxes) = "Beginning WIP Boxes"
        .CellValue(0, cnBeginWIPWeight) = "Beginning WIP Weight"
        .CellValue(0, cnProductionBoxes) = "Production Boxes"
        .CellValue(0, cnProductionWeight) = "Production Weight"
        .CellValue(0, cnEndingWIPBoxes) = "Ending WIP Boxes"
        .CellValue(0, cnEndingWIPWeight) = "Ending WIP Weight"
        .CellValue(0, cnTotalBoxes) = "Total Boxes"
        .CellValue(0, cnTotalWeight) = "Total Weight"
        .CellValue(0, cnAvgWeight) = "Avg. Weight"
        .CellValue(0, cnAction) = "Action"
        .CellValue(0, cnTruckBoxes) = "Truck Boxes"
        .CellValue(0, cnTruckWeight) = "Truck Weight"
        
        'Format the Header Row
        
        .Row = 0
        
        For lColIndex = 0 To COLUMN_COUNT - 1
           .Column = lColIndex
           .CellFontBold = True
        Next
        
        'Set the Column Types based on the access rights.
        
        If fsAccessRights = FUNC_MODIFY Then
  
            .ColumnType(cnProductCode) = Normal
            .ColumnType(cnBeginWIPBoxes) = Normal
            .ColumnType(cnBeginWIPWeight) = Normal
            .ColumnType(cnProductionBoxes) = Normal
            .ColumnType(cnProductionWeight) = Normal
            .ColumnType(cnEndingWIPBoxes) = EditBox
            .ColumnType(cnEndingWIPWeight) = EditBox
            .ColumnType(cnTotalBoxes) = Normal
            .ColumnType(cnTotalWeight) = Normal
            .ColumnType(cnAvgWeight) = Normal
            .ColumnType(cnAction) = Normal
            .ColumnType(cnTruckBoxes) = Normal
            .ColumnType(cnTruckWeight) = Normal
            
            .ColAlignment(cnProductCode) = flexAlignLeftCenter
            
            .ColumnDataType(cnEndingWIPBoxes) = Numeric
            .ColumnDataType(cnEndingWIPWeight) = Real
            
            .MaxLength(cnEndingWIPBoxes) = 6
            .MaxLength(cnEndingWIPWeight) = 11
            .DecimalLength(cnEndingWIPWeight) = 2
            .AllowNegative(cnEndingWIPBoxes) = True
            .AllowNegative(cnEndingWIPWeight) = True
            
        ElseIf fsAccessRights = FUNC_BROWSE_0NLY Then
        
            .ColumnType(cnProductCode) = Normal
            .ColumnType(cnBeginWIPBoxes) = Normal
            .ColumnType(cnBeginWIPWeight) = Normal
            .ColumnType(cnProductionBoxes) = Normal
            .ColumnType(cnProductionWeight) = Normal
            .ColumnType(cnEndingWIPBoxes) = Normal
            .ColumnType(cnEndingWIPWeight) = Normal
            .ColumnType(cnTotalBoxes) = Normal
            .ColumnType(cnTotalWeight) = Normal
            .ColumnType(cnAvgWeight) = Normal
            .ColumnType(cnAction) = Normal
            .ColumnType(cnTruckBoxes) = Normal
            .ColumnType(cnTruckWeight) = Normal
            
            .ColAlignment(cnProductCode) = flexAlignLeftCenter
            
        End If
   
        For lColIndex = 0 To COLUMN_COUNT - 1
            .Column = lColIndex
            .ColAlignmentFixed(lColIndex) = flexAlignCenterCenter
        Next
        
        .Row = 0
        .Row = 1
        .Column = 0
        flRowIndex = 1
        flColumnIndex = 0
        
        .TextMatrix(.Row, cnProductCode) = ""
        .TextMatrix(.Row, cnBeginWIPBoxes) = "0"
        .TextMatrix(.Row, cnBeginWIPWeight) = "0.00"
        .TextMatrix(.Row, cnProductionBoxes) = "0"
        .TextMatrix(.Row, cnProductionWeight) = "0.00"
        .TextMatrix(.Row, cnEndingWIPBoxes) = "0"
        .TextMatrix(.Row, cnEndingWIPWeight) = "0.00"
        .TextMatrix(.Row, cnTotalBoxes) = "0"
        .TextMatrix(.Row, cnTotalWeight) = "0.00"
        .TextMatrix(.Row, cnAvgWeight) = "0.00"
        .TextMatrix(.Row, cnAction) = ""
        .TextMatrix(.Row, cnTruckBoxes) = "0"
        .TextMatrix(.Row, cnTruckWeight) = "0.00"
        
        .Redraw = True
        
    End With
    
    ' Change the File Menu
    
    If fsAccessRights = FUNC_MODIFY Then
        mnuSave.Visible = True
        mnuClose.Visible = True
        mnuExit.Visible = True
        mnuWinClose.Visible = False
    Else
        mnuWinClose.Visible = True
        mnuExit.Visible = True
        mnuSave.Visible = False
        mnuClose.Visible = False
    End If
    
    ' Change the Actions Menu.
    
    If fsAccessRights = FUNC_MODIFY Then
        mnuInquire.Enabled = True
        mnuInquire.Visible = True
        mnuAdd.Enabled = True
        mnuAdd.Visible = True
        mnuProdBreak.Enabled = True
        mnuProdBreak.Visible = True
    Else
        mnuInquire.Enabled = True
        mnuInquire.Visible = True
        mnuAdd.Enabled = False
        mnuAdd.Visible = False
        mnuProdBreak.Enabled = True
        mnuProdBreak.Visible = True
    End If
    
    ' Change the tool bar.
        
    If fsAccessRights = FUNC_MODIFY Then
        tlbToolBar.Buttons(5).Visible = True
        tlbToolBar.Buttons(5).Enabled = True
        tlbToolBar.Buttons(6).Visible = True
        tlbToolBar.Buttons(6).Enabled = True
        tlbToolBar.Buttons(7).Visible = True
        tlbToolBar.Buttons(7).Enabled = True
    Else
        tlbToolBar.Buttons(5).Visible = True
        tlbToolBar.Buttons(5).Enabled = True
        tlbToolBar.Buttons(6).Visible = True
        tlbToolBar.Buttons(6).Enabled = True
        tlbToolBar.Buttons(7).Visible = False
        tlbToolBar.Buttons(7).Enabled = False
    End If
    
    ' Change the Status bar.
    
    If fsAccessRights = FUNC_MODIFY Then
        sbStatusBar.Panels(1).Text = _
                "<F2 - Inquire><F4 - Prod. Breakout><F5 - Add><Esc - Quit>"
    Else
        sbStatusBar.Panels(1).Text = _
                "<F2 - Inquire><F4 - Prod. Breakout><Esc - Quit>"
    End If
    
CleanUpAndExit:
        
    Exit Function
        
ErrHandler:
        
    ustgrdDailyProduction.Redraw = True
    
    FormatGrid = Err.Number
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws033", Me.Caption, vbOKOnly, gcolErrMsg
    
    GoTo CleanUpAndExit
        
End Function

'******************************************************************************
'* Functional Description   :   Load the grid with the values.
'* Parameter Description    :   None.
'* Return Type Description  :   LoadProductionDetails - error code, if any.
'******************************************************************************

Private Function LoadProductionDetails() As Long

Dim objDailyProd    As Object
Dim rsDailyProd     As ADODB.Recordset
Dim lRowCount       As Long
Dim lBegWIPBox      As Long
Dim dblBegWIPWgt    As Double
Dim lProdBox        As Long
Dim dblProdWgt      As Double
Dim lEndWIPBox      As Long
Dim dblEndWIPWgt    As Double
Dim lTotalBox       As Long
Dim dblTotalWgt     As Double
Dim dblAvgWgt       As Double

    On Error GoTo ErrHandler
    
    LoadProductionDetails = 0
    flRowsRetrieved = 0
    fbFormLoaded = True
    
    'Display the Customer and the Shift in the screen.
        
    lblCustomer.Caption = fsCustomerId & " " & Mid(fsCustomerType, 5)
    lblShift.Caption = fsShift
    
    Set objDailyProd = CreateObject("InventoryFunctions.DailyProduction")
    
    'Get the production details for the selected Customer and Shift.
    
    LoadProductionDetails = _
                    objDailyProd.GetProductionDetails(gsPlantCode, _
                                                      fsShift, _
                                                      fsCustomerId, _
                                                      Mid(fsCustomerType, 1, 1), _
                                                      rsDailyProd)
                                      
    If LoadProductionDetails <> 0 Then GoTo ErrHandler
    
    If Not rsDailyProd.EOF Then
    
        rsDailyProd.MoveFirst
        
    Else
    
        If fsAccessRights = FUNC_BROWSE_0NLY Then
        
            ShowErrorMsg "DailyProBws001", Me.Caption, vbOKOnly
            
            flRowsRetrieved = 0
            Unload Me
            Exit Function
            
        Else
        
            ShowErrorMsg "DailyProBws032", Me.Caption, vbOKOnly
            flRowsRetrieved = 0
            
            With ustgrdDailyProduction
            
                'Setting the focus to EndingWIPBoxes column
                'while loading the daily production details.
    
                .Row = 1
                .Column = cnEndingWIPBoxes
                
                .CellValue(1, cnProductionBoxes) = "0"
                .CellValue(1, cnEndingWIPBoxes) = "0"
                .CellValue(1, cnBeginWIPBoxes) = "0"
                .CellValue(1, cnProductionWeight) = "0.00"
                .CellValue(1, cnEndingWIPWeight) = "0.00"
                .CellValue(1, cnBeginWIPWeight) = "0.00"
                
                RefreshRow 1
                
                'Make the Ending WIP Boxes and the Ending WIP Weight fields
                'non-editable based on the access rights of the user.
                
                .ColumnType(cnEndingWIPBoxes) = Normal
                .ColumnType(cnEndingWIPWeight) = Normal
                
            End With
            
            Exit Function
            
        End If
        
    End If
    
    lRowCount = 1
    
    With ustgrdDailyProduction
    
        .Redraw = False
        
        While Not rsDailyProd.EOF
        
            .Rows = lRowCount + 1
            
            'Get the Boxes and Weight values for each record.
            
            lBegWIPBox = rsDailyProd.Fields("BEG_WIP_BOX").Value
            dblBegWIPWgt = rsDailyProd.Fields("BEG_WIP_WGT").Value
            lProdBox = rsDailyProd.Fields("PROD_BOX").Value
            dblProdWgt = rsDailyProd.Fields("PROD_WGT").Value
            lEndWIPBox = rsDailyProd.Fields("END_WIP_BOX").Value
            dblEndWIPWgt = rsDailyProd.Fields("END_WIP_WGT").Value
            lTotalBox = lProdBox + lEndWIPBox - lBegWIPBox
            dblTotalWgt = dblProdWgt + dblEndWIPWgt - dblBegWIPWgt
            
            'Calculate the Average Weight for the record.
            
            If lTotalBox <> 0 Then
                dblAvgWgt = dblTotalWgt / lTotalBox
            Else
                dblAvgWgt = 0
            End If
            
            'Modify the Average Weight based on the upper and the lower limit.
            
            If dblAvgWgt >= 1000 Then
                dblAvgWgt = 999.99
            ElseIf dblAvgWgt <= -100 Then
                dblAvgWgt = -99.99
            End If
            
            'Set the Boxes and Weight values for each
            'record to the Daily Production Grid.
            
            .CellValue(lRowCount, cnProductCode) = _
                    rsDailyProd.Fields("PRODUCT_CODE").Value
            .CellValue(lRowCount, cnBeginWIPBoxes) = _
                    lBegWIPBox
            .CellValue(lRowCount, cnBeginWIPWeight) = _
                    Format(dblBegWIPWgt, "#0.00")
            .CellValue(lRowCount, cnProductionBoxes) = _
                    lProdBox
            .CellValue(lRowCount, cnProductionWeight) = _
                    Format(dblProdWgt, "#0.00")
            .CellValue(lRowCount, cnEndingWIPBoxes) = _
                    lEndWIPBox
            .CellValue(lRowCount, cnEndingWIPWeight) = _
                    Format(dblEndWIPWgt, "#0.00")
            .CellValue(lRowCount, cnTotalBoxes) = _
                    lTotalBox
            .CellValue(lRowCount, cnTotalWeight) = _
                    Format(dblTotalWgt, "#0.00")
            .CellValue(lRowCount, cnAvgWeight) = _
                    Format(dblAvgWgt, "#0.00")
            .CellValue(lRowCount, cnTruckBoxes) = _
                   rsDailyProd.Fields("TRUCK_BOX").Value
            .CellValue(lRowCount, cnTruckWeight) = _
                   Format(rsDailyProd.Fields("TRUCK_WGT").Value, _
                          "#0.00")
            
            lRowCount = lRowCount + 1
            rsDailyProd.MoveNext
            
        Wend
        
        .Redraw = True
        
        'Make the Ending WIP Boxes and the Ending WIP Weight fields
        'editable based on the access rights of the user.
        
        If fsAccessRights = FUNC_MODIFY Then
        
            ustgrdDailyProduction.ColumnType(cnEndingWIPBoxes) = EditBox
            ustgrdDailyProduction.ColumnType(cnEndingWIPWeight) = EditBox
            
        End If
        
    End With
    
    'No of records retrived from the database
    'for the selected Customer and Shift.
    
    flRowsRetrieved = lRowCount - 1
    
    'Set the focus to the Ending WIP Boxes column.
    
    ustgrdDailyProduction.Row = 1
    ustgrdDailyProduction.Column = cnEndingWIPBoxes
    
    
CleanUpAndExit:
    
    Set objDailyProd = Nothing
    Set rsDailyProd = Nothing
    Exit Function

ErrHandler:

    frmMainMenu.MousePointer = vbDefault

    If LoadProductionDetails <> 0 Then
    
        gcolErrMsg.Add LoadProductionDetails
        ShowErrorMsg "DailyProBws009", Me.Caption, vbOKOnly, gcolErrMsg
        
    Else
    
        gcolErrMsg.Add Err.Description
        ShowErrorMsg "DailyProBws016", Me.Caption, vbOKOnly, gcolErrMsg
        
    End If
    
    frmMainMenu.sbMainMenu.Visible = False
    frmMainMenu.mnuViewStatusBar.Checked = False
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    GoTo CleanUpAndExit
    
End Function

'******************************************************************************
'* Functional Description   :   Unloads form without updating database
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuClose_Click()

    fbCloseWithoutUpdate = True
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the Help Contents.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuHelpContents_Click()

Dim nRet As Integer
    
    If Len(App.HelpFile) = 0 Then
        ShowErrorMsg "Help001", "Freezer Inventory System", _
                     vbOKOnly
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Commodities Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCommoditiesUpdate_Click()

    CommoditiesUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Customer Master Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerMasterUpdate_Click()

    CustomerMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Memo Master Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuMemoMasterUpdate_Click()

    MemoMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Plant Master Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuPlantMasterUpdate_Click()

    PlantMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Rate Code Master Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuRateCodeMasterUpdate_Click()

    RateCodeMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Rate Product Master
'*                              Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuProductMasterUpdate_Click()

    ProductMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Order Shipping Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOrderShippingUpdate_Click()

    OrderShippingUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Daily Production Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuDailyProductionUpdate_Click()

    DailyProductionUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Captures Ctrl+C key for Copy action.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCopy_Click()

    SendKeys "^C"
    
End Sub

'******************************************************************************
'* Functional Description   :   Captures Ctrl+X key for Cut action.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCut_Click()

    SendKeys "^X"
    
End Sub

'******************************************************************************
'* Functional Description   :   Captures Ctrl+V key for Paste action.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuPaste_Click()

    SendKeys "^V"
    
End Sub

'******************************************************************************
'* Functional Description   :   Exit the Application.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuExit_Click()

    Unload Me
    Unload frmMainMenu
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the Status Bar.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuViewStatusBar_Click()

    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    gbStatusBarIsVisible = mnuViewStatusBar.Checked
    Form_Resize
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the toolbar.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuViewToolbar_Click()

    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tlbToolBar.Visible = mnuViewToolbar.Checked
    gbToolBarIsVisible = mnuViewToolbar.Checked
    Form_Resize
    
End Sub

'******************************************************************************
'* Functional Description   :   Unloads form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWinClose_Click()

    Unload Me

End Sub

'******************************************************************************
'* Functional Description   :   Cascade the windows.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWindowCascade_Click()

    frmMainMenu.Arrange vbCascade
    
End Sub

'******************************************************************************
'* Functional Description   :   Tiles the windows horizontal.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWindowTileHorizontal_Click()

    frmMainMenu.Arrange vbTileHorizontal
    
End Sub

'******************************************************************************
'* Functional Description   :   Tiles the window vertical.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWindowTileVertical_Click()

    frmMainMenu.Arrange vbTileVertical
    
End Sub

'******************************************************************************
'* Functional Description   :   Show About dialog box.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuHelpAbout_Click()

    frmAbout.Show vbModal, frmMainMenu
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the toolbar button clicks.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub tlbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case biCut
            SendKeys "^X"
        Case biCopy
            SendKeys "^C"
        Case biPaste
            SendKeys "^V"
        Case biInquire
            mnuInquire_Click
        Case biProdBreak
            mnuProdBreak_Click
        Case biAdd
            mnuAdd_Click
    End Select
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Detail Inventory Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuDetailInventoryUpdate_Click()

    DetailInventoryUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Pallet Movement.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuPalletMovement_Click()

    PalletMovementClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Freezer Inventory Browse.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuFreezerInventoryBrowse_Click()

    FreezerInventoryBrowseClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Order Summary Browse.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOrderSummaryBrowse_Click()

    OrderSummaryBrowseClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Order Information Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOrderInformationUpdate_Click()

    OrderInformationUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Customer Storage Browse.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerStorageBrowse_Click()

    CustomerStorageBrowseClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Inquire.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuInquire_Click()

Dim sData As String
    
    On Error GoTo ErrHandler
    
    'Validate any invalid entry in the Grid.
    
    If Not IsValidGridData Then Exit Sub
    
    If fsAccessRights = FUNC_MODIFY Then
    
        'Check if any updations are made in the Grid. If so, display
        'the confirmation message to save the production data.
        
        If GridEdited Then
        
            If ShowErrorMsg("DailyProBws002", Me.Caption, vbYesNo) = vbYes Then
                
                'Save the daily production details after confirmation.
                
                If Not SaveProductionChanges Then
                
                    Exit Sub
                    
                End If
                
            End If
            
        End If
        
    End If
    
    'Display the inquire screen.
    
    Set frmInquire.gfrmGeneral = gfrmDailyProductUpdate
    frmInquire.Show vbModal
    
    'Set the screen for the new Customer and Shift in the inquire screen.
    
    If Not fbInquireCancelled Then
    
        'Set the default values for the Boxes and the Weight fields in the Grid.
        
        sData = "" & vbTab & "0" & vbTab & "0.00" & vbTab & _
                    "0" & vbTab & "0.00" & vbTab & "0" & vbTab & _
                    "0.00" & vbTab & "0" & vbTab & "0.00" & vbTab & "0.00"
                    
        ustgrdDailyProduction.AddItems sData, 1
        ustgrdDailyProduction.Rows = 2
        flRowIndex = 1
        
        'Load the production data for the selected Customer and Shift.
        
        LoadProductionDetails
        
    End If
    
CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws017", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Add.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuAdd_Click()
        
    'Validate for any invalid entry in the Grid.
    
    If Not IsValidGridData Then Exit Sub
          
    'Display the Add screen.
    
    Set frmAddProduct.gfrmGeneral = Me
    frmAddProduct.Show vbModal
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item customer Storage Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerStorageUpdate_Click()

    CustomerStorageUpdateClick

End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Customer Charges Update.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerChargesUpdate_Click()

    CustomerChargesUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Product Break Out.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuProdBreak_Click()
    
    On Error GoTo ErrHandler
    
    'Validate for any invalid entry in the Grid.
    
    If Not IsValidGridData Then Exit Sub
    
    With ustgrdDailyProduction
    
        'Do not display the Product Breakout screen, if there
        'are no records in the daily production grid.
        
        If flRowsRetrieved = 0 And .TextMatrix(.Row, cnProductCode) = "" Then

            Exit Sub
            
        End If
        
        'Pass the details for the selected record in the Daily
        'Production Grid to the Product Breakout Screen.
        
        frmProductBreak.fsProductCode = .TextMatrix(.Row, cnProductCode)
        frmProductBreak.flTruckBoxes = .TextMatrix(.Row, cnTruckBoxes)
        frmProductBreak.fdblTruckWgt = .TextMatrix(.Row, cnTruckWeight)
        frmProductBreak.flTotalBoxes = .TextMatrix(.Row, cnProductionBoxes)
        frmProductBreak.fdblTotalWgt = .TextMatrix(.Row, cnProductionWeight)

    End With
    
    'Display the Product Breakout screen.
    
    frmProductBreak.Show vbModal
    
CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws018", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item Product Break Out.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuBreakOut_Click()
    
    mnuProdBreak_Click
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the blast cell unloading report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuBlastCellUnloadingReport_Click()
    
    BlastCellUnloadingReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the daily production report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuDailyProductionReport_Click()
    
    DailyProductionReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the detail production report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuDetailProductionReport_Click()
    
    DetailProductionReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the code by code report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCodebyCodeReport_Click()
    
    CodebyCodeReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the aged product inventory report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuAgedProductInventoryReport_Click()
    
    AgedProductInventoryReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the freezer inventory report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuFreezerInventoryReport_Click()
    
    FreezerInventoryReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the weekly freezer recap report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuWeeklyFreezerRecapReport_Click()
    
    WeeklyFreezerRecapReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows message box that Cold Storage Commodity
'*                          :   report printing is ON.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuColdStorageCommodityReport_Click()
    
    ColdStorageCommodityReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the code by code vs inventory report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCodeVsInventoryReport_Click()
    
    CodeVsInventoryReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the first screen of customer master report.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCustomerMasterReport_Click()
    
    CustomerMasterReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the customer storage due report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCustomerStorageDueReport_Click()
    
    CustomerStorageDueReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the customer storage detail report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCustomerStorageDetailReport_Click()
    
    CustomerStorageDetailReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the customer charges report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCustomerChargesReport_Click()
    
    CustomerChargesReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Weekly Sales Summary report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuWeeklysalesSummaryReport_Click()
    
    WeeklySalesSummaryReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the message box that Customer inventory
'*                          :   by product report printing is ON.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCustomerInventoryByProductReport_Click()
    
    CustomerInventoryByProductReportClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the customer storage recap report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCustomerStorageRecapReport_Click()

    CustomerStorageRecapReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the customer warehouse receipts report
'*                          :   screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuCustomerWarehouseReceipts_Click()
    
    CustomerWarehouseReceiptsClick
    SetNoTopmost
    
End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Order Manifest detail report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuOrderManifestDetailReport_Click()
    
    OrderManifestDetailReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Order Manifest Summary report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuOrderManifestSummaryReport_Click()
    
    OrderManifestSummaryReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Order pullsheet report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuOrderPullSheetReport_Click()
    
    OrderPullSheetReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Order summary report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuOrderSummaryReport_Click()
    
    OrderSummaryReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Product recall report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuProductRecallReport_Click()
    
    ProductRecallReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Outside customer BOL report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuOutsideCustomerBOL_Click()
    
    OutsideCustomerBOLClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Inbound/Outbound pallet report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuInboundOutboundPalletReport_Click()
    
    InboundOutboundPalletReportClick
    SetNoTopmost

End Sub

'*******************************************************************************
'* Functional Description   :   Shows the Pallet detail by order report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuPalletDetailByOrderReport_Click()
    
    PalletDetailByOrderReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the screen to select the plant id for
'*                              an admin user.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuSelectPlant_Click()

    SelectPlant

End Sub

'******************************************************************************
'* Functional Description   :   Handles the Menu item AdjustmentLogBrowse
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuAdjustmentLogBrowse_Click()

    AdjustmentLogBrowseClick

End Sub

'******************************************************************************
'* Functional Description   :   Handles the Cell Data Modified event
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Sub ustgrdDailyProduction_CellDataModified()

Dim iInvalidColumn As Integer

    On Error GoTo ErrHandler
    
    If Not IsValidGridData Then
        
        Exit Sub
        
    Else
             
        CellDataModified ustgrdDailyProduction, cnAction, flRowIndex
        
    End If

CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws019", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Adds a row in the grid with the entered values.
'* Parameter Description    :   sProdCode       - ProductCode (validated)
'*                              lEndingWIPBoxes - Ending WIP Boxes
'*                              dblEndingWIPWgt - Ending WIP Weight
'* Return Type Description  :   True, if new row is added, False if not.
'******************************************************************************

Public Function AddRow(ByVal sProdCode As String, _
                       ByVal lEndingWIPBoxes As Long, _
                       ByVal dblEndingWIPWgt As Double) As Boolean
Dim lCount As Long

    On Error GoTo ErrHandler
    
    AddRow = False
    
    With ustgrdDailyProduction
        
        'Check for the duplicate entry for the Product Code in the Grid.
        
        For lCount = 1 To .Rows - 1
        
            If .CellValue(lCount, cnProductCode) = sProdCode Then
            
                ShowErrorMsg "DailyProBws004", Me.Caption, vbOKOnly
                
                AddRow = False
                Exit Function
                
            End If
            
        Next
        
        'If only one row is present and if it is a blank one, reuse that.
        
        If Not (flRowsRetrieved = 0 And .CellValue(.Row, cnProductCode) = "") Then
        
            'Otherwise create a new row.
            
            .Rows = .Rows + 1
            
        End If
        
        .Row = .Rows - 1
        
        'Set the user inputs into the Daily Production Grid from the Add Screen.
        
        .CellValue(.Row, cnProductCode) = sProdCode
        .CellValue(.Row, cnBeginWIPBoxes) = "0"
        .CellValue(.Row, cnBeginWIPWeight) = "0.00"
        .CellValue(.Row, cnProductionBoxes) = "0"
        .CellValue(.Row, cnProductionWeight) = "0.00"
        .CellValue(.Row, cnEndingWIPBoxes) = lEndingWIPBoxes
        .CellValue(.Row, cnEndingWIPWeight) = Format(dblEndingWIPWgt, "#0.00")
        .CellValue(.Row, cnTotalBoxes) = "0"
        .CellValue(.Row, cnTotalWeight) = "0.00"
        .CellValue(.Row, cnAvgWeight) = "0.00"
        .CellValue(.Row, cnAction) = TO_INSERT
        .CellValue(.Row, cnTruckBoxes) = "0"
        .CellValue(.Row, cnTruckWeight) = "0.00"
        
        'Set the Total and the Average fields based on the Ending WIP fields.
        
        RefreshRow .Row
        
        'Set the focus to the newly added record.
        
        .TopRow = .Row
        .Column = cnEndingWIPBoxes
        
        'Make the Ending WIP Boxes and the Ending WIP Weight fields editable.
        
        .ColumnType(cnEndingWIPBoxes) = EditBox
        .ColumnType(cnEndingWIPWeight) = EditBox
        
    End With
    AddRow = True
    
CleanUpAndExit:
        
    Exit Function
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws020", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Function

'******************************************************************************
'* Functional Description   :   Handles the KeyPress event of the grid.
'* Parameter Description    :   KeyAscii - Ascii code of the key entered
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ustgrdDailyProduction_KeyPress(KeyAscii As Integer)

Dim lColumnIndex As Long
    
    On Error GoTo ErrHandler
    
    With ustgrdDailyProduction
    
        If KeyAscii = vbKeyEscape Then
        
            If .Column <> cnEndingWIPBoxes And _
               .Column <> cnEndingWIPWeight Then
               
                If Not IsValidGridData Then Exit Sub
                
                If GridEdited Then

                    If fsAccessRights = FUNC_MODIFY Then
                    
                        If Not SaveProductionChanges Then
                        
                            Exit Sub
                            
                        End If
                            
                    End If
                    
                End If
                
                Unload Me
                
            ElseIf (.GridEditMode = Cell) Then
            
                If Not IsValidGridData Then Exit Sub
                
                If GridEdited Then
                
                    If fsAccessRights = FUNC_MODIFY Then
                        
                        If Not SaveProductionChanges Then
                        
                            Exit Sub
                            
                        End If
                        
                    End If
                    
                End If
                
                Unload Me
                
            Else
            
                .GridEdit vbKeyEscape
                
            End If
            
        End If
        
    End With

CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws022", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Handles the Enter Cell event of the grid.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ustgrdDailyProduction_EnterCell()
    
    On Error GoTo ErrHandler
    
    If fbValidated Then
    
        fbValidated = False
        ustgrdDailyProduction.Row = flRowIndex
        ustgrdDailyProduction.Column = flColumnIndex
        Exit Sub
        
    End If
    
    IsValidGridData

CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws021", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Validates a whole row, or a single column
'* Parameter Description    :   lRow,iInvalidColumn
'* Return Type Description  :   Boolean value to return if valid row or not
'******************************************************************************

Private Function ValidateRow(ByVal lRow As Long, _
                             ByRef iInvalidColumn As Integer) _
                             As Boolean

    On Error GoTo ErrHandler
    
    ValidateRow = True
    iInvalidColumn = -1
    
    With ustgrdDailyProduction

        If .CellValue(lRow, cnEndingWIPWeight) > 99999999.99 Then

            ShowErrorMsg "DailyProBws007", Me.Caption, vbOKOnly
            iInvalidColumn = cnEndingWIPWeight
            ValidateRow = False
            fbValidated = True
            
        ElseIf .CellValue(lRow, cnEndingWIPBoxes) > 999999 Then
            
            ShowErrorMsg "DailyProBws008", Me.Caption, vbOKOnly
            iInvalidColumn = cnEndingWIPBoxes
            ValidateRow = False
            fbValidated = True
            
        End If
        
    End With

CleanUpAndExit:
        
    Exit Function
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws024", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Function

'******************************************************************************
'* Functional Description   :   Refreshes the calculations in a row
'* Parameter Description    :   lRow - the row to be refreshed
'* Return Type Description  :   None.
'******************************************************************************

Private Sub RefreshRow(ByVal lRow As Long)

    On Error GoTo ErrHandler
    
    With ustgrdDailyProduction
        
        .Redraw = False
        
        If .CellValue(lRow, cnProductionBoxes) = "" Then _
            .CellValue(lRow, cnProductionBoxes) = "0"
            
        If .CellValue(lRow, cnEndingWIPBoxes) = "" Then _
            .CellValue(lRow, cnEndingWIPBoxes) = "0"
            
        If .CellValue(lRow, cnBeginWIPBoxes) = "" Then _
            .CellValue(lRow, cnBeginWIPBoxes) = "0"
            
        If .CellValue(lRow, cnProductionWeight) = "" Then _
            .CellValue(lRow, cnProductionWeight) = "0.00"
            
        If .CellValue(lRow, cnEndingWIPWeight) = "" Then _
            .CellValue(lRow, cnEndingWIPWeight) = "0.00"
            
        If .CellValue(lRow, cnBeginWIPWeight) = "" Then _
            .CellValue(lRow, cnBeginWIPWeight) = "0.00"
            
        'Calculate the Total Boxes and set the value to the Grid.
        
        .CellValue(lRow, cnTotalBoxes) = CStr( _
                CDbl(.CellValue(lRow, cnProductionBoxes)) + _
                CDbl(.CellValue(lRow, cnEndingWIPBoxes)) - _
                CDbl(.CellValue(lRow, cnBeginWIPBoxes)))
                
        'Calculate the Total Weight and set the value
        'to the Grid in the format 0.00.
        
        .CellValue(lRow, cnTotalWeight) = Format(CStr( _
                CDbl(.CellValue(lRow, cnProductionWeight)) + _
                CDbl(.CellValue(lRow, cnEndingWIPWeight)) - _
                CDbl(.CellValue(lRow, cnBeginWIPWeight))), "#0.00")

        'Calculate the Average Weight based on the Total Boxes field.
        
        If .CellValue(lRow, cnTotalBoxes) <> 0 Then
            .CellValue(lRow, cnAvgWeight) = CStr( _
                    CDbl(.CellValue(lRow, cnTotalWeight)) / _
                    CDbl(.CellValue(lRow, cnTotalBoxes)))
        Else
            .CellValue(lRow, cnAvgWeight) = "0"
        End If
        
        'Modify the Average Weight based on the upper and the lower limit.
        
        If .CellValue(lRow, cnAvgWeight) >= 1000 Then
            .CellValue(lRow, cnAvgWeight) = 999.99
        ElseIf .CellValue(lRow, cnAvgWeight) <= -100 Then
            .CellValue(lRow, cnAvgWeight) = -99.99
        End If
        
        'Set the Average Weight to the Daily Production
        ' Grid in the format 0.00.
        
        .CellValue(lRow, cnAvgWeight) = _
                Format(.CellValue(lRow, cnAvgWeight), "#0.00")
                
        .Redraw = True
        
    End With

CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:

    ustgrdDailyProduction.Redraw = True
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws025", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Update table with changes done in grid
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean value to indicate whether saved or not
'******************************************************************************

Private Function SaveProductionChanges() As Boolean

Dim objDailyProd    As Object
Dim colProdRows     As Collection
Dim colDailyProd    As Collection
Dim lRowCount       As Long
Dim lResult         As Long
Dim lInvalidColumn  As Long
Dim lRowIndex       As Long
Dim lColumnIndex    As Long
Dim sStatusText     As String
    
    On Error GoTo ErrHandler
    
    SaveProductionChanges = False
    
    If Not GridEdited Then
        
        'No data to update, exit.
        
        SaveProductionChanges = True
        Exit Function
        
    End If
    
    If Not IsValidGridData Then Exit Function
    
    With ustgrdDailyProduction
        
        'If Product code is empty, this has to be a default empty row.
        
        If Trim(.CellValue(1, cnProductCode)) = "" Then
            
            'No data to update, exit.
            
            SaveProductionChanges = True
            Exit Function
            
        End If
        
        ustgrdDailyProduction_EnterCell
        
        lRowIndex = .Row
        lColumnIndex = .Column
        
        'Construct the collection of records to be updated from the grid.
        
        For lRowCount = 1 To .Rows - 1
        
            If Trim(.CellValue(lRowCount, cnAction)) <> "" Then
                
                'Create a colProdRows collection for each row of the grid
                
                Set colProdRows = New Collection
                
                If Left(.CellValue(lRowCount, cnAction), 1) = TO_INSERT Or _
                    Left(.CellValue(lRowCount, cnAction), 1) = TO_UPDATE Then
                        
                    colProdRows.Add gsPlantCode, "PLANT_CODE"
                    colProdRows.Add fsShift, "SHIFT_CODE"
                    colProdRows.Add fsCustomerId, "CUSTOMER_ID"
                    colProdRows.Add Mid(fsCustomerType, 1, 1), "CUSTOMER_TYPE"

                    colProdRows.Add .CellValue(lRowCount, cnProductCode), _
                                    "PRODUCT_CODE"
                    colProdRows.Add .CellValue(lRowCount, cnEndingWIPBoxes), _
                                    "END_WIP_BOX"
                    colProdRows.Add .CellValue(lRowCount, cnEndingWIPWeight), _
                                    "END_WIP_WGT"
                    colProdRows.Add Left(.CellValue(lRowCount, cnAction), 1), _
                                    "ACTION"
                    colProdRows.Add CStr(lRowCount), "ROW_INDEX"
                    
                    'Append the colProdRows collection to the
                    'main collection colDailyProd
                    
                    If colDailyProd Is Nothing Then
                        Set colDailyProd = New Collection
                    End If
                    
                    colDailyProd.Add colProdRows, CStr(lRowCount)
                End If
            End If
        Next
        
        If Not colDailyProd Is Nothing Then
            
            'Call the update method of DailyProduction to update the database.
            
            Set objDailyProd = _
                        CreateObject("InventoryFunctions.DailyProduction")
                        
            sStatusText = sbStatusBar.Panels(1).Text
            sbStatusBar.Panels(1).Text = _
                                       "Updating Database - Please Wait..."
            Me.MousePointer = vbHourglass
            DoEvents
            
            lResult = objDailyProd.UpdateProductionDetails(colDailyProd)
            
            Me.MousePointer = vbDefault
            
            If Not colDailyProd Is Nothing Then

                'Display Error
                
                mnuErrors.Visible = True
                DisplayErrors ustgrdDailyProduction, cnAction, _
                                 COLUMN_COUNT, colDailyProd, cnProductCode

                lResult = ShowErrorMsg("DailyProBws006", Me.Caption, vbOKOnly)

                Set fcolErrors = New Collection
                Set fcolErrors = colDailyProd
                PopulateErrorWindow fcolErrors
                
                'Set the focus back to the cell which had the
                'focus before the error window appeared.
                
                .Row = lRowIndex
                .Column = lColumnIndex
                .SetFocus
                
                SaveProductionChanges = False
                GoTo CleanUpAndExit
                
             ElseIf lResult <> 0 Then
                        
                gcolErrMsg.Add lResult
                ShowErrorMsg "DailyProBws010", Me.Caption, vbOKOnly, gcolErrMsg
                
                .SetFocus
                SaveProductionChanges = False
                GoTo CleanUpAndExit
                
             Else
             
                SaveProductionChanges = True
                
             End If
             
        Else
        
            SaveProductionChanges = True
            
        End If
        
        'If Updates successful or No updates in screen
        
        Set fcolErrors = Nothing
        mnuErrors.Visible = False
        CloseErrorWindow
         
    End With
    
CleanUpAndExit:

    Me.MousePointer = vbDefault
    sbStatusBar.Panels(1).Text = sStatusText
    
    Set objDailyProd = Nothing
    Set colProdRows = Nothing
    Set colDailyProd = Nothing
    Exit Function
    
ErrHandler:
    
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws026", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
    
End Function

'******************************************************************************
'* Functional Description   :   Handles the mouse down event on the grid and
'*                              displays the popup menu with actions
'* Parameter Description    :   Button pressed and status of shift
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ustgrdDailyProduction_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ErrHandler
    
    With ustgrdDailyProduction
    
        If Button = vbRightButton And .GridEditMode = Cell Then
        
            'Display Error
            
            If Not fcolErrors Is Nothing Then
                mnuError.Visible = True
            Else
                mnuError.Visible = False
            End If
            
            'Display Error
            
            If fsAccessRights = FUNC_MODIFY Then
                PopupMenu mnuPopup
            Else
                mnuAddProduct.Visible = False
                mnuAddProduct.Enabled = False
                PopupMenu mnuPopup
            End If
            
        End If
    End With

CleanUpAndExit:
        
    Exit Sub
    
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws027", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'**************************************************************************************
'* Functional Description   :   Highlights row corresponding to error selected by user
'* Parameter Description    :   sKey - The key fields separated by " | "
'* Return Type Description  :   None.
'**************************************************************************************

Public Sub ErrorSelected(ByVal sKey As String)

Dim lRowIndex As Long
    
    On Error GoTo ErrHandler
    
    With ustgrdDailyProduction
        For lRowIndex = 1 To .Rows - 1
            If .CellValue(lRowIndex, cnProductCode) = _
            sKey Then
                .Row = lRowIndex
                .TopRow = lRowIndex
                .Column = cnProductCode
                .SetFocus
                GoTo CleanUpAndExit
            End If
        Next
    End With
    
CleanUpAndExit:
        
    Exit Sub
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws028", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Sub

'******************************************************************************
'* Functional Description   :   Checks if the grid has been edited
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean.
'******************************************************************************

Private Function GridEdited() As Boolean

Dim lRowCt As Long
    
    On Error GoTo ErrHandler
    With ustgrdDailyProduction
        For lRowCt = 1 To .Rows - 1
            If .TextMatrix(lRowCt, cnAction) <> "" And _
               .TextMatrix(lRowCt, cnAction) <> "A" Then
                GridEdited = True
                Exit Function
            End If
        Next
    End With
    GridEdited = False
    
CleanUpAndExit:
        
    Exit Function
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws029", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
        
End Function

'******************************************************************************
'* Functional Description   :   Validates the Grid values
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean.
'******************************************************************************

Private Function IsValidGridData() As Boolean

Dim iInvalidColumn As Integer

    On Error GoTo ErrHandler
    
    IsValidGridData = True
    
    With ustgrdDailyProduction
    
        If .Column = cnEndingWIPBoxes Or .Column = cnEndingWIPWeight Then
            
            If .Column = cnEndingWIPWeight And _
                        Val(.CellValue(.Row, cnEndingWIPWeight)) = 0 Then
                        
                .CellValue(.Row, cnEndingWIPWeight) = 0
                
            End If
        
        
            If .Column = cnEndingWIPWeight Then
            
                .CellValue(.Row, cnEndingWIPWeight) = _
                    Format(.CellValue(.Row, cnEndingWIPWeight), "#0.00")
                    
            End If
            
            If .Column = cnEndingWIPBoxes And _
                        Val(.CellValue(.Row, cnEndingWIPBoxes)) = 0 Then
                        
                .CellValue(.Row, cnEndingWIPBoxes) = 0
                
            End If
            
        End If
        
        If ValidateRow(flRowIndex, iInvalidColumn) Then
        
            RefreshRow .Row
            flRowIndex = .Row
            flColumnIndex = .Column
            
        Else
        
            .Row = flRowIndex
            .Column = iInvalidColumn
            IsValidGridData = False
            .SetFocus
            
        End If
    
    End With
   
CleanUpAndExit:
        
    Exit Function
        
ErrHandler:
        
    gcolErrMsg.Add Err.Description
    ShowErrorMsg "DailyProBws023", Me.Caption, vbOKOnly, gcolErrMsg
    GoTo CleanUpAndExit
    
End Function


Private Sub mnuOrderHistoryBrowse_Click()

' Added by TCS

OrderHistoryBrowse1Click

End Sub


Private Sub mnuOrderHistorySummaryReport_Click()

' Added by TCS
    
    OrderHistorySummaryReportClick
    SetNoTopmost
    
End Sub


Private Sub mnuOrderHistoryManifestReport_Click()

' Added by TCS
  
    OrderHistoryManifestReportClick
    SetNoTopmost
    

End Sub

' Added by TCS
Private Sub mnuOrderListingReport_Click()

    OrderListingReportClick
    SetNoTopmost

End Sub
Private Sub mnuToBatchScanner_Click()
'Added by TCS
 ToBatchScannerClick
 
End Sub

Private Sub mnuFromBatchScanner_Click()
' Added by TCS
     
     FromBatchScannerClick
     

End Sub
'Added By TCS to Print Order Shortage Report
Private Sub mnuOrderShortageReport_Click()
    OrderShortageReportClick
    SetNoTopmost
End Sub

Private Sub mnuSummaryInventoryActivityReport_Click()
    
    SummaryInventoryActivityReportClick
    SetNoTopmost
    
End Sub

' Added by TCS
Private Sub mnuDetailInventoryAdd_Click()
DetailInventoryAddClick
End Sub
' Added by TCS
'For Req 4(B) Part II
Private Sub mnuInvProdStatusUpdate_Click()
Inventory_ProductStatusUpdate
End Sub

'For Req 10v1
'Added by TCS
Private Sub mnuCustomerShippingPoundsReport_Click()
    Cust_Ship_Pound_Report
End Sub

'Added By TCS
'Req 3v1
Private Sub mnuProductivityTracking_Click()
    ProductivityTracking
    'SetNoTopmost
End Sub


Private Sub mnuPalletLPNDetailReport_Click()
    PalletLPNDetailReportClick
    SetNoTopmost
End Sub


