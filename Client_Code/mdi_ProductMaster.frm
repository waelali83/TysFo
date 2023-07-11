VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{82CA1823-AE1A-4850-BDC7-F24C9A05E6D0}#1.1#0"; "UstriGrid.ocx"
Begin VB.Form mdi_ProductMaster 
   Caption         =   "Product Master Browse/Update"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   Icon            =   "mdi_ProductMaster.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   8460
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMain 
      Height          =   5055
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   8175
      Begin USTriSuperGrid.USTriGrid ustgrdProductMaster 
         Height          =   4695
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8281
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
   Begin MSComctlLib.Toolbar tlbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print (Ctrl + P)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inquire"
            Description     =   "Inquire"
            Object.ToolTipText     =   "Inquire (F2)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Description     =   "Add"
            Object.ToolTipText     =   "Add (F5)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Update"
            Description     =   "Update"
            Object.ToolTipText     =   "Update (F6)"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5550
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12999
            MinWidth        =   8819
            Text            =   "<F2 - Inquire><F4 - XRef><F5 - Add><F6 - Update><F9 - Secondary><Ctrl+P - Print><Esc - Quit>"
            TextSave        =   "<F2 - Inquire><F4 - XRef><F5 - Add><F6 - Update><F9 - Secondary><Ctrl+P - Print><Esc - Quit>"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "10:49 AM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":1CFA
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":1E0C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":1F1E
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":2030
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":26AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":2F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":385E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":4138
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_ProductMaster.frx":4A12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSaveAndClose 
         Caption         =   "&Save And Close"
      End
      Begin VB.Menu mnuCloseWithoutSave 
         Caption         =   "&Close Without Save"
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
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
         Enabled         =   0   'False
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
         Visible         =   0   'False
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
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Update"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuPrintPM 
         Caption         =   "&Print"
         Shortcut        =   ^P
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
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuInquirePopup 
         Caption         =   "&Inquire"
      End
      Begin VB.Menu mnuXrefPopup 
         Caption         =   "&XRef"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddPopup 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuUpdatePopup 
         Caption         =   "&Update"
      End
      Begin VB.Menu mnuSecondaryPopup 
         Caption         =   "&Secondary"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuError 
         Caption         =   "Show E&rrors"
      End
   End
   Begin VB.Menu mnuNoAccessPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuInquirePopupBrowse 
         Caption         =   "&Inquire"
      End
      Begin VB.Menu mnuXrefPopupBrowse 
         Caption         =   "&XRef"
      End
      Begin VB.Menu mnuSecondaryPopupBrowse 
         Caption         =   "&Secondary"
      End
   End
End
Attribute VB_Name = "mdi_ProductMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*File Name                     : mdi_ProductMaster.frm
'*File Description              : Used to browse/update Product Master
'*Author                        : US Technology
'*Date Created                  : Sep-01-02
'*Date Last Modified            : Mar-17-03
'*Version                       : 2.0
'*Layer                         : Client
'*Project Referenced            : MasterFileFunctions
'*Components Used               : USTriGrid
'*Functions Defined             : ShowAddScreen,FormatGrid,
'*                                LoadProductDetails, LoadTypeRecordSets,
'*                                PrintProductMasterReports,
'*                                SetLabelCodesCollection
'*                                SetMenuForAccessRights
'*                                SetSecondaryCustomerMenu, UpdateProductPlant
'*                                SetProductCodesCollection,
'*                                ShowUpdateScreen
'*Copyright                     : US Technology
'-------------------------------------------------------------------------------
'*Change History:
'*Change Code   Source(defect ID)  Change Description    Date     Author
'*Initial Release                                      Dec-19-02  US Technology
'*Second Release                                       Apr-05-03  US Technology
'******************************************************************************
Option Explicit

Private Enum ButtonIndex

    biCut = 1
    biCopy = 2
    biPaste = 3
    biPrint = 5
    biInquire = 7
'    biXRef = 8
'    biSecondary = 9
    biAdd = 9
    biUpdate = 10
    
End Enum

Private Enum ColumnName

    cnProductCode = 0
    cnProductDescription = 1
    cnDivisionCode = 2
    cnPrdTyp = 3
    cnLblNbr = 4
    cnWgtTyp = 5
    cnBxsPlt = 6
    cnMinWgt = 7
    cnMaxWgt = 8
    cnTareWgt = 9
    cnBlastInd = 10
    cnFrzDays = 11
    cnBlstDays = 12
    cnLblLen = 13
    cnStrPos = 14
    cnWgtLen = 15
    cnCustNum = 16
    cnCustType = 17
    cnComdCode = 18
    cnGovtLot = 19
    cnMfgDate = 20
    cnLabelCodeChk = 21
    cnLabelCodeStr = 22
    cnXRef = 23
    cnSecondIndex = 24
    cnProductCodeKey = 25
    cnAction = 26
    cnLabelCodeLen = 27
    cnBoxSerialInd = 28
    cnProductGroupCode = 29 'Added by TCS on 03/28/04
    cnMfgDatePos = 30 'Added by TCS-Ragu on 17-Jun-2005
   
End Enum

Private frsProdTypes            As ADODB.Recordset
Private frsWgtTypes             As ADODB.Recordset
Private frsPrdGrpCode           As ADODB.Recordset 'Added by TCS
Private frsProdDet              As ADODB.Recordset


Private fbCloseWithoutUpdate    As Boolean
Private fbCrossRefProductsOnly  As Boolean
Private fbAddOnly               As Boolean


'Added by TCS
Private sFltrProdCode           As String
Private sFltrDivCode            As String
Private sFltrProdType           As String
Private sFltrLabelNo            As String
'''

Private fcolProductCodes        As Collection
Private fcolErrors              As Collection

Private flLoadSuccess           As Long

'Private WithEvents ffrmXRef     As frmProductLabel
'Private WithEvents ffrmSecCust  As frmSecondary

Private WithEvents ffrmInquire  As frmProductCodeInquire
Attribute ffrmInquire.VB_VarHelpID = -1
Private WithEvents ffrmAddUpd   As frmProductLabel
Attribute ffrmAddUpd.VB_VarHelpID = -1

Private Const COLUMN_COUNT      As Integer = 31 'Added by TCS
Private Const CHILD_FORM_WIDTH = 12500

Public fsAccessRights           As String

'Added by TCS
Private WithEvents ffrmProductMasterInquire As FrmProductMasterInquire
Attribute ffrmProductMasterInquire.VB_VarHelpID = -1

'******************************************************************************
'* Functional Description   :   Refresh the grid with new details
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ffrmAddUpd_RefreshGrid(ByVal sMode As String, _
                                   ByVal colProduct As Collection)
    
Dim lRowIndex       As Long



 LoadProductDetails
    
 Exit Sub

    With ustgrdProductMaster
        
        'Insertion to Grid
        
        If sMode = "I" Then
            
            .AddItems ""
            lRowIndex = .Rows - 1
             
             If .CellValue(lRowIndex - 1, cnProductCode) = "" Then
                
                .RemoveItem (lRowIndex)
                lRowIndex = lRowIndex - 1
                fbAddOnly = False
            
            End If
        
        'Updation to Grid
        
        ElseIf sMode = "U" Then
            
            lRowIndex = .Row
        
        End If
        
        'Appending or Updating the grid
        
        .CellValue(lRowIndex, cnProductCode) = _
                colProduct.Item("PRODUCT_CODE")
                
        If Left(.CellValue(lRowIndex, cnAction), 1) = TO_DELETE Or _
                Left(.CellValue(lRowIndex, cnAction), 1) = _
                TO_DELETE_AFTER_UPDATE Then
            
            .CellValue(lRowIndex, cnProductDescription) = DELETED
            .CellValue(lRowIndex, cnAction) = _
                            Left(.CellValue(lRowIndex, cnAction), 1) & _
                                      colProduct.Item("PRODUCT_DESC")
        
        Else
            
            .CellValue(lRowIndex, cnProductDescription) = _
                colProduct.Item("PRODUCT_DESC")
        
        End If
        
        .CellValue(lRowIndex, cnPrdTyp) = _
                WriteTypeTextToGrid( _
                colProduct.Item("PRODUCT_TYPE"), _
                frsProdTypes)
        .CellValue(lRowIndex, cnDivisionCode) = _
                colProduct.Item("DIVISION_CODE")
        .CellValue(lRowIndex, cnLblNbr) = _
                colProduct.Item("LABEL_NO")
        .CellValue(lRowIndex, cnWgtTyp) = _
                WriteTypeTextToGrid( _
                colProduct.Item("WGT_TYPE_CODE"), _
                frsWgtTypes)
        .CellValue(lRowIndex, cnBxsPlt) = _
                colProduct.Item("BOXES_PER_PALLET")
        .CellValue(lRowIndex, cnMinWgt) = _
                Format(colProduct.Item("MIN_BOX_WGT"), _
                "#0.00")
        .CellValue(lRowIndex, cnMaxWgt) = _
                Format(colProduct.Item("MAX_BOX_WGT"), _
                "#0.00")
        .CellValue(lRowIndex, cnTareWgt) = _
                Format(colProduct.Item("BOX_TARE_WEIGHT"), _
                "#0.00")
        .CellValue(lRowIndex, cnBlastInd) = _
                colProduct.Item("BLAST_IND")
        .CellValue(lRowIndex, cnFrzDays) = _
                colProduct.Item("FREEZE_DAYS")
        .CellValue(lRowIndex, cnBlstDays) = _
                colProduct.Item("BLAST_DAYS")
        .CellValue(lRowIndex, cnLblLen) = _
                colProduct.Item("LABEL_LENGTH")
        .CellValue(lRowIndex, cnStrPos) = _
                colProduct.Item("LABEL_WGT_ST_POS")
        .CellValue(lRowIndex, cnWgtLen) = _
                colProduct.Item("LABEL_WGT_LENGTH")
        .CellValue(lRowIndex, cnCustNum) = _
                colProduct.Item("CUSTOMER_ID")
        .CellValue(lRowIndex, cnComdCode) = _
                colProduct.Item("GOVT_COMMODITY_CODE")
        .CellValue(lRowIndex, cnGovtLot) = _
                colProduct.Item("GOVT_LOT_IND")
        .CellValue(lRowIndex, cnMfgDate) = _
                colProduct.Item("PACK_DATE_IND")
                
        'Added by TCS-Ragu on 17-Jun-2005
        'Start
        .CellValue(lRowIndex, cnMfgDatePos) = _
                colProduct.Item("PACK_DATE_START")
        'End
        
        .CellValue(lRowIndex, cnLabelCodeChk) = _
                colProduct.Item("CHECK_LABEL_IND")
        .CellValue(lRowIndex, cnLabelCodeStr) = _
                colProduct.Item("PROD_LABEL_START")
        .CellValue(lRowIndex, cnLabelCodeLen) = _
                colProduct.Item("PROD_LABEL_LEN")
        .CellValue(lRowIndex, cnProductCodeKey) = _
                colProduct.Item("PRODUCT_CODE")
        .CellValue(lRowIndex, cnBoxSerialInd) = _
                colProduct.Item("BOX_SERIAL_IND")
        .CellValue(lRowIndex, cnProductGroupCode) = _
                colProduct.Item("PRODUCT_GROUP_CODE") 'Added by TCS on 28/03/2004
        
        .CellValue(lRowIndex, cnXRef) = _
                colProduct.Item("XREF_CODE") 'Added by TCS on 28/03/2004
        
        'Set Customer Type
' *********** commented by TCS ***********************
'        Select Case colProduct.Item("CUSTOMER_TYPE")
'
'            Case "S"
'                .CellValue(lRowIndex, cnCustType) = _
'                colProduct.Item("CUSTOMER_TYPE") & _
'                " - Ship To"
'            Case "C"
'                .CellValue(lRowIndex, cnCustType) = _
'                colProduct.Item("CUSTOMER_TYPE") & _
'                " - Corporate"
'            Case "B"
'                .CellValue(lRowIndex, cnCustType) = _
'                colProduct.Item("CUSTOMER_TYPE") & _
'                " - Bill To"
        
'        End Select
        
        .TopRow = lRowIndex
        .Row = lRowIndex
        
    End With
    
End Sub

'******************************************************************************
'* Functional Description   :   Obtains the value of the Product Code from
'*                              frmProductCodeInquire screen.
'* Parameter Description    :   sProductCode - The value from Inquire screen.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ffrmInquire_ProductMasterInquire(ByVal sProductCode As String)

Dim lRowIndex As Long
    
    'Search for the inquired value in the grid
    
    With ustgrdProductMaster
        
        lRowIndex = SoftSeek(ustgrdProductMaster, sProductCode, cnProductCode)
        .TopRow = lRowIndex
        .Row = lRowIndex
        .Column = cnProductCode
    
    End With
    
End Sub

'******************************************************************************
'* Functional Description   :   Assigns T to Second_Ind filed in grid.
'* Parameter Description    :   sXRefValue - The value from XREF screen.
'* Return Type Description  :   None.
'******************************************************************************

'Private Sub ffrmSecCust_SecondaryCustomers()
'
'    'Set cnSecondIndex to true if secondary customer updated
'
'    With ustgrdProductMaster
'
'        .CellValue(.Row, cnSecondIndex) = "T"
'        CellDataModified ustgrdProductMaster, cnAction, .Row, cnProductDescription
'
'    End With
'
'End Sub
'
'******************************************************************************
'* Functional Description   :   Obtains the value of the XRef from
'*                              frmXRef screen.
'* Parameter Description    :   sXRefValue - The value from XREF screen.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ffrmXRef_XRef(ByVal sXRefValue As String)

    'Set the XRef value in the grid

    With ustgrdProductMaster

        .CellValue(.Row, cnXRef) = sXRefValue
        CellDataModified ustgrdProductMaster, cnAction, .Row, cnProductDescription

    End With

End Sub

Private Sub ffrmProductMasterInquire_PopulateProductMaster(ByVal sProductCode _
                                                As String, ByVal sDivisionCode _
                                                As String, ByVal sProductType _
                                                As String, ByVal sLabelNo As String)

sFltrProdCode = sProductCode
sFltrDivCode = sDivisionCode
sFltrProdType = sProductType
sFltrLabelNo = sLabelNo

End Sub

'******************************************************************************
'* Functional Description   :   Sets the StatusBar and ToolBar Visiblity and
'*                              resizes the form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Activate()

    fbCrossRefProductsOnly = False
    fbCloseWithoutUpdate = False
    
    'Populate Error Window
    
    If Not gbCloseErrWindow Then
        
        Set gfrmCurrentMdi = Me
        PopulateErrorWindow fcolErrors
    
    End If
       
    SetViewMenu Me
    Form_Resize
    
End Sub

'******************************************************************************
'* Functional Description   :   Hide the Error Window
'* Parameter Description    :   Status of Cancel buttton.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Terminate()

    Unload Me

End Sub

'******************************************************************************
'* Functional Description   :   Hide the Error Window
'* Parameter Description    :   Status of Cancel buttton.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Unload(Cancel As Integer)
    
    'Setting the error collection to nothing and unloading the error window

    Set fcolErrors = Nothing
    CloseErrorWindow
    
    'Check to ensure that this is the only form loaded and hence while
    'unloading set the status bar visibility as true.This is to ensure
    'that only one status bar is visible throughout the application. 2 is
    'subtracted to exempt the MDI parent(frmMainMenu) and the error window

    If (Forms.Count - 2) = 1 Then
    
        frmMainMenu.sbMainMenu.Visible = True
        frmMainMenu.mnuViewStatusBar.Checked = True
        
    End If
    
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

'*******************************************************************************
'* Functional Description   :   The entry point to this mdi from modMain. Invokes
'                               Form_Load()
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Public Function LoadProductMaster() As Boolean
    
    On Error GoTo ErrHandler
    
    
    
    'Call Form Load
    
    Set ffrmProductMasterInquire = New FrmProductMasterInquire
    
    If ffrmProductMasterInquire.LoadFormProductMasterBrowseUpdate Then
        LoadProductMaster = True
        Load Me
    Else
        Unload Me
        LoadProductMaster = False
    End If
    
    If flLoadSuccess <> 0 Then GoTo ErrHandler
    Exit Function
    
ErrHandler:

    LoadProductMaster = False
    
End Function


'******************************************************************************
'* Functional Description   :   Loads the form to the standard size.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Load()

Dim lResult As Long
    
    On Error GoTo ErrHandler
    
    fbAddOnly = False
    lResult = 0
    flLoadSuccess = 0
    
    'Initializing the status bar

    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.sbMainMenu.Panels(1).Text = _
        "Loading From Database - Please Wait..."
    frmMainMenu.MousePointer = vbHourglass
    DoEvents
    
    Me.Move 0, 0, frmMainMenu.Width / 1.03, frmMainMenu.Height / 1.15
    
    'Enabling Admin menu if user has admin privilege

    mnuAdmin.Visible = gbEnableAdminMenu
    mnuExportOrderVerification.Visible = gbEnableExportVerifyMenu
    
    'Added by TCS (To show the Batch Scanner Menu)
    mnuFromBatchScanner.Visible = gbEnableBatchScannerMenu
    mnuToBatchScanner.Visible = gbEnableBatchScannerMenu
    'Addition ends here
    
    mnuInitProdSend.Visible = frmMainMenu.mnuInitProdSend.Visible   'Added by TCS on 08-Sep-2005
    mnuSelectPlant.Enabled = frmMainMenu.mnuSelectPlant.Enabled 'Added by TCS on 08-Sep-2005

    'Set visibility of File & Action Menu items, Status bar
    'and toolbar based on access privileges
    
    SetMenuForAccessRights
    
    'Load Type Details
    
    lResult = LoadTypeRecordsets
    
    'Get the Product details
    
    If lResult = 0 Then
        
        lResult = LoadProductDetails
        If lResult <> 0 Then GoTo ErrHandler
    
    Else: If lResult <> 0 Then GoTo ErrHandler
    
    End If
    
    'Reset Main Menu and status bar
    
    SetMainMenu
        
    Exit Sub

ErrHandler:

    flLoadSuccess = lResult
    
    'Reset Main Menu and status bar
    
    frmMainMenu.sbMainMenu.Visible = True
    frmMainMenu.mnuViewStatusBar.Checked = True
    frmMainMenu.MousePointer = vbDefault
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    'Call of Form_Terminate to unload the form

    Form_Terminate
       
End Sub

'******************************************************************************
'* Functional Description   :   Update Database and unload form
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Call of various methods. On failure of any of these methods, the unload
    'is cancelled. Saving of changes to database is applicable only for user
    'with access rights Modify.

    If fsAccessRights = FUNC_MODIFY And _
        Not fbCloseWithoutUpdate And flLoadSuccess = 0 Then
        
        If Not UpdateProductPlant Then Cancel = 1
    End If
End Sub

'******************************************************************************
'* Functional Description   :   Positions and resizes the form controls when the
'*                              form is resized.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub Form_Resize()

Dim iColIndex       As Integer

    'Checking window state and setting the window state accordingly

    If Me.WindowState = vbMinimized Or _
    frmMainMenu.WindowState = vbMinimized Then Exit Sub
    
    'Ensuring that the child form's width and height does not reduce beyond
    'the level set
    
    If Me.WindowState <> vbMaximized Then
        
        If Me.Width < CHILD_FORM_WIDTH Then Me.Width = CHILD_FORM_WIDTH
        If Me.Height < CHILD_FORM_HEIGHT Then Me.Height = CHILD_FORM_HEIGHT
    
    End If
    
    SetFramePosition Me, fraMain
    
    'Resizing of the grid control

    With ustgrdProductMaster
        
        .Move 100, 200, _
        fraMain.Width - 200, fraMain.Height - 300
                    
        For iColIndex = 0 To COLUMN_COUNT - 1
        
            ' Set the width for the grid columns
            
            Select Case iColIndex
                
                Case cnProductDescription
                    .ColWidth(iColIndex) = .Width / 3
                
                Case cnComdCode
                    .ColWidth(iColIndex) = .Width / 8
                
                Case cnProductCode
                    .ColWidth(iColIndex) = .Width / 5.5
                    
                Case cnCustType
                    .ColWidth(iColIndex) = 0  'Changed to 0 by TCS
                    
                Case cnPrdTyp
                    .ColWidth(iColIndex) = .Width / 7
                    
                Case cnCustNum
                    .ColWidth(iColIndex) = 0  'Changed to 0 by TCS
                    
                Case cnDivisionCode, cnStrPos
                    .ColWidth(iColIndex) = .Width / 11.5
                    
                Case cnFrzDays, cnBlstDays
                    .ColWidth(iColIndex) = 0 'Added by TCS
                    
                Case cnWgtTyp
                    .ColWidth(iColIndex) = .Width / 9
                    
                Case cnMinWgt, cnMaxWgt, cnTareWgt
                    .ColWidth(iColIndex) = .Width / 9
                
                Case cnWgtLen
                    .ColWidth(iColIndex) = .Width / 13
                    
                Case cnLabelCodeLen, cnLblLen, _
                        cnMfgDate, cnMfgDatePos, cnLabelCodeStr 'Case cnMfgDatePos added by TCS-Ragu on 17-Jun-2005
                    .ColWidth(iColIndex) = .Width / 11
                
                Case cnLblNbr, cnBxsPlt, _
                         cnGovtLot, _
                        cnLabelCodeChk, cnBoxSerialInd
                    .ColWidth(iColIndex) = .Width / 10.5
                
                Case cnBlastInd
                    .ColWidth(iColIndex) = 0  'Added by TCS
                    
                Case cnProductGroupCode
                .ColWidth(iColIndex) = .Width / 6  'Added by TCS
                
                Case cnAction, cnXRef, cnSecondIndex, cnProductCodeKey
                    .ColWidth(iColIndex) = 0
                    
            End Select
            
        Next
        
    End With
    
End Sub

'******************************************************************************
'* Functional Description   :   Format the look and feel of the grid.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub FormatGrid()
    
Dim iColIndex               As Integer
        
    'Initializing various properties of the Grid
    
    With ustgrdProductMaster
        
        .FixedRows = 1
        .RowHeight(0) = 3 * GRID_ROW_HEIGHT
        .WordWrap = True
        .Columns = COLUMN_COUNT
        
        ' Set Grid columns
        
        .CellValue(0, cnProductCode) = "Product Code"
        .CellValue(0, cnProductDescription) = "Product Description"
        .CellValue(0, cnPrdTyp) = "Product Type"
        .CellValue(0, cnDivisionCode) = "Division Code"
        .CellValue(0, cnLblNbr) = "Label No."
        .CellValue(0, cnWgtTyp) = "Weight Type"
        .CellValue(0, cnBxsPlt) = "Boxes Per Pallet"
        .CellValue(0, cnMinWgt) = "Min. Weight"
        .CellValue(0, cnMaxWgt) = "Max. Weight"
        .CellValue(0, cnTareWgt) = "Tare Weight"
        .CellValue(0, cnBlastInd) = "Blast Ind."
        .CellValue(0, cnFrzDays) = "Freezer Days"
        .CellValue(0, cnBlstDays) = "Blast Days"
        .CellValue(0, cnLblLen) = "Label Length"
        .CellValue(0, cnStrPos) = "Start Pos."
        .CellValue(0, cnWgtLen) = "Weight Length"
        .CellValue(0, cnCustNum) = "Customer No."
        .CellValue(0, cnCustType) = "Customer Type"
        .CellValue(0, cnComdCode) = "Commodity Code"
        .CellValue(0, cnGovtLot) = "Govt. Lot"
        .CellValue(0, cnMfgDate) = "Mfg. Date"
        
        'Added by TCS-Ragu on 17-Jun-2005
        'Start
        .CellValue(0, cnMfgDatePos) = "Mfg. Date Pos"
        'End
        
        .CellValue(0, cnLabelCodeChk) = "Label Code Check"
        .CellValue(0, cnLabelCodeStr) = "Label Code Start"
        .CellValue(0, cnLabelCodeLen) = "Label Code Length"
        .CellValue(0, cnBoxSerialInd) = "Box Serial Ind."
        '.CellValue(0, cnXRef) = "XREF"
        '.CellValue(0, cnSecondIndex) = "SecInd"
        .CellValue(0, cnAction) = "Action"
        .CellValue(0, cnProductCodeKey) = "ProdCodeKey"
        .CellValue(0, cnProductGroupCode) = "ProdGrpCode" ' Added by TCS
        
        'Format the Header Row
        
        .Row = 0
        
        For iColIndex = 0 To COLUMN_COUNT - 1
           
           .Column = iColIndex
           .CellFontBold = True
        
        Next
        
        'Set the Column Types based on the access rights
        
        'Normal   - User will not be able to edit but only browse
        'Editbox  - User will be able to handle it like a text box
        'Combobox - Control is a combo box

        If fsAccessRights = FUNC_BROWSE_0NLY Then
            
            For iColIndex = 0 To COLUMN_COUNT - 1
                
                .ColumnType(iColIndex) = Normal
            
            Next
        
        End If
       
        'Set Combo Alignment to be displayed
      
        For iColIndex = 0 To COLUMN_COUNT - 1
            
            .ColAlignmentFixed(iColIndex) = flexAlignCenterCenter
            
            Select Case iColIndex
                
                Case cnProductCode, cnProductDescription, cnLblNbr, _
                        cnComdCode, cnGovtLot, _
                        cnWgtTyp, cnPrdTyp, cnBoxSerialInd, _
                        cnMfgDate, cnLabelCodeChk, cnDivisionCode, cnProductGroupCode  'Added by TCS
                    .ColAlignment(iColIndex) = flexAlignLeftBottom
                
                Case Else
                    .ColAlignment(iColIndex) = flexAlignRightBottom
            
            End Select
        
        Next
        
        .Rows = 2
        
    End With
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets the menus as per the access rights
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub SetMenuForAccessRights()

    'Setting visibility of File Menu items based on access priveleges
    
    If fsAccessRights = FUNC_MODIFY Then
    
        mnuClose.Visible = False
        
    Else
    
        mnuSaveAndClose.Visible = False
        mnuCloseWithoutSave.Visible = False
        
    End If
    
    'Set status bar,Actions menu and toolbar for BROWSE ONLY access
    
    If fsAccessRights = FUNC_BROWSE_0NLY Then
    
        mnuAdd.Enabled = False
        mnuUpdate.Enabled = False
        'mnuXref.Enabled = True
        'mnuXrefPopup.Enabled = True
        'tlbToolBar.Buttons("XRef").Visible = True
        tlbToolBar.Buttons("Add").Visible = False
        tlbToolBar.Buttons("Update").Visible = False
        'mnuXrefPopup.Visible = True
              
           
         '   mnuSecondary.Enabled = True
          '  tlbToolBar.Buttons("Secondary").Visible = True
            sbStatusBar.Panels(1).Text = "<F2 - Inquire>" & _
                                         "<Ctrl+P - Print>" & _
                                         "<Esc - Quit>"
        
                
    ElseIf fsAccessRights = FUNC_MODIFY Then
    
        'Set status bar,Actions menu and toolbar for MODIFY access
    
        mnuAdd.Enabled = True
        mnuUpdate.Enabled = True
        'mnuXref.Enabled = True
        tlbToolBar.Buttons("Add").Visible = True
        tlbToolBar.Buttons("Update").Visible = True
        SetSecondaryCustomerMenu
        
    End If
 
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
'* Functional Description   :   Enables the F9 menu for plants 35,40,47,53,206
'*                              and disables it for others
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub SetSecondaryCustomerMenu()
    
 
        sbStatusBar.Panels(1).Text = "<F2 - Inquire>" & _
                                    "<F5 - Add><F6 - Update>" & _
                                    "<Ctrl+P - Print>" & _
                                    "<Esc - Quit>"
    
 
    
End Sub

'******************************************************************************
'* Functional Description   :   Populate the types recorsets to format display
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Function LoadTypeRecordsets() As Long

Dim objProdMstr         As Object
Dim lResult             As Long

    On Error GoTo ErrHandler
    
    LoadTypeRecordsets = 0
    
    'Create object of ProductMaster Class in MasterFileFunctions
    
  Set objProdMstr = CreateObject("MasterFileFunctions.ProductMaster")
         
    'Populate the Product Types
    
    lResult = objProdMstr.GetProductTypes(frsProdTypes)
    If lResult <> 0 Then GoTo ErrHandler
    
    'Populate the Weight Types

    lResult = objProdMstr.GetWeightTypes(frsWgtTypes)
    If lResult <> 0 Then GoTo ErrHandler
    
    'Populate the Product group code ADDED BY TCS
    
    lResult = objProdMstr.GetProductGrpCode(frsPrdGrpCode)
    If lResult <> 0 Then GoTo ErrHandler
    
CleanUpAndExit:
    
    Set objProdMstr = Nothing
    Exit Function
    
ErrHandler:

    frmMainMenu.sbMainMenu.Visible = False
    frmMainMenu.mnuViewStatusBar.Checked = False
    frmMainMenu.MousePointer = vbDefault
    frmMainMenu.sbMainMenu.Panels(1).Text = STATUS_TEXT
    
    If lResult <> 0 Then
        
        'Display Server error
        
        LoadTypeRecordsets = lResult
        gcolErrMsg.Add LoadTypeRecordsets
        giErrMsg = ShowErrorMsg("ProMstr002", Me.Caption, _
                                         vbOKOnly, gcolErrMsg)
    Else
    
        'Display VB error
    
        LoadTypeRecordsets = Err.Number
        gcolErrMsg.Add LoadTypeRecordsets
        giErrMsg = ShowErrorMsg("ProMstr002", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    End If

   GoTo CleanUpAndExit
   
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
    ustgrdProductMaster.Redraw = False
    FormatGrid
    lRowIndex = 0
    
    If frsProdDet Is Nothing Then
    'Create object of ProductMaster Class in MasterFileFunctions
    
    Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
    
    'Retrieve Product Details
    
    LoadProductDetails = objProduct.GetProductDetailsForPlant(gsPlantCode, _
                                                              rsProducts)
    
    Else
    Set rsProducts = frsProdDet
    
    
    End If
    
    If LoadProductDetails <> 0 Then GoTo ErrHandler
    
    If rsProducts Is Nothing Then mnuClose_Click
    
    'Added by TCS
    If IsFilterAvailable(SFilter) Then
       rsProducts.Filter = SFilter
      ' rsProducts.Filter = 0
       If rsProducts.EOF Or rsProducts.BOF Then
'            giErrMsg = ShowErrorMsg("ProMstr011", _
                                    Me.Caption, _
                                    vbOKOnly)
'            fbAddOnly = True
            'form Termination
'            LoadProductDetails = 1
           ' GoTo CleanUpAndExit
       End If

    End If
    
    With ustgrdProductMaster
    
        If Not rsProducts.EOF Then
        
            rsProducts.MoveFirst
            lRowIndex = 1
            
            'Add the records in the recordset to the grid
            
            While Not rsProducts.EOF
                
                If .Rows < lRowIndex + 1 Then
                    .AddItems ""
                End If
                
                If .CellBackColor = COLOR_YELLOW Then
                    .Row = lRowIndex
                    
                    For lColIndex = 0 To COLUMN_COUNT - 1
                        .Column = lColIndex
                        .CellBackColor = COLOR_WHITE
                    Next
                
                End If
                               
                'Read values of each field for a record
                
                .CellValue(lRowIndex, cnProductCode) = _
                        rsProducts.Fields("PRODUCT_CODE").Value
                .CellValue(lRowIndex, cnProductDescription) = _
                        rsProducts.Fields("PRODUCT_DESC").Value
                .CellValue(lRowIndex, cnPrdTyp) = _
                        WriteTypeTextToGrid( _
                        rsProducts.Fields("PRODUCT_TYPE").Value, _
                        frsProdTypes)
                .CellValue(lRowIndex, cnDivisionCode) = _
                        rsProducts.Fields("DIVISION_CODE").Value
                .CellValue(lRowIndex, cnLblNbr) = _
                        rsProducts.Fields("LABEL_NO").Value
                .CellValue(lRowIndex, cnWgtTyp) = _
                        WriteTypeTextToGrid( _
                        rsProducts.Fields("WGT_TYPE_CODE").Value, _
                        frsWgtTypes)
                .CellValue(lRowIndex, cnBxsPlt) = _
                        rsProducts.Fields("BOXES_PER_PALLET").Value
                .CellValue(lRowIndex, cnMinWgt) = _
                        Format(rsProducts.Fields("MIN_BOX_WGT").Value, _
                        "#0.00")
                .CellValue(lRowIndex, cnMaxWgt) = _
                        Format(rsProducts.Fields("MAX_BOX_WGT").Value, _
                        "#0.00")
                .CellValue(lRowIndex, cnTareWgt) = _
                        Format(rsProducts.Fields("BOX_TARE_WEIGHT").Value, _
                        "#0.00")
                        
                .CellValue(lRowIndex, cnBlastInd) = "0" 'Changed by TCS
                        
                .CellValue(lRowIndex, cnFrzDays) = 0 'Changed by TCS
                
                .CellValue(lRowIndex, cnBlstDays) = 0 'Changed by TCS
                        
                        
                .CellValue(lRowIndex, cnLblLen) = _
                        rsProducts.Fields("LABEL_LENGTH").Value
                .CellValue(lRowIndex, cnStrPos) = _
                        rsProducts.Fields("LABEL_WGT_ST_POS").Value
                .CellValue(lRowIndex, cnWgtLen) = _
                        rsProducts.Fields("LABEL_WGT_LENGTH").Value
                .CellValue(lRowIndex, cnCustNum) = _
                        rsProducts.Fields("CUSTOMER_ID").Value
                
                'Added by TCS-Ragu on 06-Jan-06 for Tk#547007
                .CellValue(lRowIndex, cnCustType) = _
                        rsProducts.Fields("CUSTOMER_TYPE").Value

                If Not IsNull(rsProducts.Fields("GOVT_COMMODITY_CODE").Value) Then
                    .CellValue(lRowIndex, cnComdCode) = _
                        rsProducts.Fields("GOVT_COMMODITY_CODE").Value
                End If
                
                .CellValue(lRowIndex, cnGovtLot) = _
                        rsProducts.Fields("GOVT_LOT_IND").Value
                .CellValue(lRowIndex, cnMfgDate) = _
                        rsProducts.Fields("PACK_DATE_IND").Value
                        
                'Added by TCS-Ragu on 17-Jun-2005
                'Start
                .CellValue(lRowIndex, cnMfgDatePos) = _
                        rsProducts.Fields("PACK_DATE_START").Value
                'End
                
                .CellValue(lRowIndex, cnLabelCodeChk) = _
                        rsProducts.Fields("CHECK_LABEL_IND").Value
                .CellValue(lRowIndex, cnLabelCodeStr) = _
                        rsProducts.Fields("PROD_LABEL_START").Value
                .CellValue(lRowIndex, cnLabelCodeLen) = _
                        rsProducts.Fields("PROD_LABEL_LEN").Value
                .CellValue(lRowIndex, cnProductCodeKey) = _
                        rsProducts.Fields("PRODUCT_CODE").Value
                
                If Not IsNull(rsProducts.Fields("XREF_CODE").Value) Then
                    
                    .CellValue(lRowIndex, cnXRef) = _
                        Trim(rsProducts.Fields("XREF_CODE").Value)
                
                End If
                
                If Not IsNull(rsProducts.Fields("BOX_SERIAL_IND").Value) Then
                    
                    .CellValue(lRowIndex, cnBoxSerialInd) = _
                        rsProducts.Fields("BOX_SERIAL_IND").Value
                
                End If
                
                ' If is Added by TCS
                
                If Not IsNull(rsProducts.Fields("PRODUCT_GROUP_CODE").Value) Then
                    
                    .CellValue(lRowIndex, cnProductGroupCode) = _
                        rsProducts.Fields("PRODUCT_GROUP_CODE").Value
                
                End If
                
'                'Set Customer Type
' ******************* commented by TCS
'
'                Select Case rsProducts.Fields("CUSTOMER_TYPE").Value
'
'                    Case "S"
'                        .CellValue(lRowIndex, cnCustType) = _
'                        rsProducts.Fields("CUSTOMER_TYPE").Value & _
'                        " - Ship To"
'                    Case "C"
'                        .CellValue(lRowIndex, cnCustType) = _
'                        rsProducts.Fields("CUSTOMER_TYPE").Value & _
'                        " - Corporate"
'                    Case "B"
'                        .CellValue(lRowIndex, cnCustType) = _
'                        rsProducts.Fields("CUSTOMER_TYPE").Value & _
'                        " - Bill To"
'
'                End Select
'
                lRowIndex = lRowIndex + 1
                rsProducts.MoveNext
                
            Wend
            
            .Column = 1
            .Column = 0
            .Row = 1
            
        Else
            
            giErrMsg = ShowErrorMsg("ProMstr011", _
                                    Me.Caption, _
                                    vbOKOnly)
            fbAddOnly = True
            
        End If
        
    End With
    
    ustgrdProductMaster.Redraw = True
    
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
     
    mnuClose_Click
    GoTo CleanUpAndExit
    
End Function

'******************************************************************************
'* Functional Description   :   Update the Product_Plant table
'* Parameter Description    :   None.
'* Return Type Description  :   Boolean - True if Update is success.
'******************************************************************************

Private Function UpdateProductPlant() As Boolean

Dim objProduct         As Object
Dim colProducts        As Collection
Dim colProduct         As Collection
Dim lRowCount          As Long
Dim lResult            As Long
Dim lRowIndex          As Long
Dim lColumnIndex       As Long
Dim sStatusText        As String

    On Error GoTo ErrHandler
    
    UpdateProductPlant = True
    
    With ustgrdProductMaster
        
        lRowIndex = .Row
        lColumnIndex = .Column
            
        'Construct the collection of records to be updated from the grid
    
        For lRowCount = 1 To .Rows - 1
            
            If Trim(.CellValue(lRowCount, cnAction)) <> "" Then
            
                'Create a colProduct collection for each row of the grid
                
                Set colProduct = New Collection
                
                If Left(.CellValue(lRowCount, cnAction), 1) = TO_DELETE Or _
                    Left(.CellValue(lRowCount, cnAction), 1) = _
                             TO_DELETE_AFTER_UPDATE Or _
                    Left(.CellValue(lRowCount, cnAction), 1) = TO_INSERT Or _
                    Left(.CellValue(lRowCount, cnAction), 1) = TO_UPDATE Then
                    
                    colProduct.Add _
                        Left(Trim(.CellValue(lRowCount, cnAction)), 1), _
                        "ACTION"
                    colProduct.Add -1, "RECORDCOUNT"
                    
                    colProduct.Add _
                            "Y", _
                            "SECOND_IND"  'Changed by TCS
                            
                    colProduct.Add .CellValue(lRowCount, cnXRef), _
                            "XREF_CODE"
                    colProduct.Add _
                            Trim(.CellValue(lRowCount, cnProductCodeKey)), _
                            "PRODUCT_CODE_KEY"
                    colProduct.Add CStr(lRowCount), "ROW_INDEX"
                    
                    'Added by TCS-Ragu on 05-Jan-06 for Tk547007
                    'Start
                    colProduct.Add _
                            Trim(.CellValue(lRowCount, cnCustNum)), _
                            "CUSTOMER_ID"
                    colProduct.Add _
                            Trim(.CellValue(lRowCount, cnCustType)), _
                            "CUSTOMER_TYPE"
                    'End
                    
                    'Append the colProduct collection to the main collection
                    'colProducts
                    
                    If colProducts Is Nothing Then
                        
                        Set colProducts = New Collection
                    
                    End If
                    
                    colProducts.Add colProduct, _
                                .CellValue(lRowCount, cnProductCodeKey)
                
                End If
            
            End If
        
        Next
        
        'Call the update method of ProductMaster to update the db
        
        If Not colProducts Is Nothing Then
            
            'Create object of ProductMaster Class in MasterFileFunctions

            Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
            
            sStatusText = sbStatusBar.Panels(1).Text
            sbStatusBar.Panels(1).Text = _
                                   "Updating Database - Please Wait..."
            Me.MousePointer = vbHourglass
            DoEvents
            
            'Update Product Details
            
            lResult = objProduct.UpdateProductDetailsForPlant( _
                                                            gsPlantCode, _
                                                            colProducts, False)
            
            Me.MousePointer = vbDefault
            sbStatusBar.Panels(1).Text = sStatusText
            
            If Not colProducts Is Nothing Then
            
                'If error occurs during updating, display error window
                
                mnuErrors.Visible = True
                DisplayErrors ustgrdProductMaster, cnAction, _
                            COLUMN_COUNT, colProducts, cnProductCode

                lResult = ShowErrorMsg("ProMstr003", Me.Caption, vbOKOnly)

                ustgrdProductMaster.Row = lRowIndex
                ustgrdProductMaster.Column = lColumnIndex
                ustgrdProductMaster.SetFocus
                ustgrdProductMaster.TopRow = lRowIndex

                Set fcolErrors = New Collection
                Set fcolErrors = colProducts
                PopulateErrorWindow fcolErrors
                
                UpdateProductPlant = False
                
                GoTo CleanUpAndExit
                
            Else: If lResult <> 0 Then GoTo ErrHandler
            
            End If
            
        End If
        
        Set fcolErrors = Nothing
        mnuErrors.Visible = False
        CloseErrorWindow
        
    End With

CleanUpAndExit:
    
    Set colProduct = Nothing
    Set colProducts = Nothing
    Set objProduct = Nothing
    Exit Function

ErrHandler:

    UpdateProductPlant = False
    
    If lResult <> 0 Then
        
        'Display Server error
        
        gcolErrMsg.Add lResult
        giErrMsg = ShowErrorMsg("ProMstr004", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    Else
        
        'Display VB error
        
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("ProMstr005", Me.Caption, _
                               vbOKOnly, gcolErrMsg)
    End If

    GoTo CleanUpAndExit
    
End Function

'******************************************************************************
'* Functional Description   :   Displays the screen to add a new record
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub ShowAddScreen()

Dim colLblCodes As Collection

    Set ffrmAddUpd = New frmProductLabel
    Set ffrmAddUpd.frsWgtTypes = frsWgtTypes
    Set ffrmAddUpd.frsProdTypes = frsProdTypes
    Set ffrmAddUpd.frsPrdGrpCode = frsPrdGrpCode 'Added by TCS
    
    SetLabelCodesCollection "I", colLblCodes
    
    Set ffrmAddUpd.fcolLabelCodes = colLblCodes
    
    SetProductCodeCollection

    Set ffrmAddUpd.fcolProducts = fcolProductCodes
    
    ffrmAddUpd.Show vbModal
     

End Sub

'******************************************************************************
'* Functional Description   :   Form the collection of Product Codes
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub SetProductCodeCollection()
    
Dim lIndex As Long
    
On Error Resume Next
    Set fcolProductCodes = New Collection
    
    With ustgrdProductMaster
        
        For lIndex = 1 To .Rows - 1
            
            fcolProductCodes.Add .CellValue(lIndex, cnProductCodeKey), _
                            CStr(.CellValue(lIndex, cnProductCodeKey))
        
        Next
    
    End With
    
End Sub

'******************************************************************************
'* Functional Description   :   Displays the screen to update a record
'* Parameter Description    :   None
'* Return Type Description  :   None
'******************************************************************************

Private Sub ShowUpdateScreen()

Dim colProduct  As Collection
Dim colLblCodes As Collection

    On Error GoTo ErrHandler
       
    Set ffrmAddUpd = New frmProductLabel
    
     
    
    Set ffrmAddUpd.frsWgtTypes = frsWgtTypes
    Set ffrmAddUpd.frsProdTypes = frsProdTypes
    Set ffrmAddUpd.frsPrdGrpCode = frsPrdGrpCode 'Added by TCS
    
    SetLabelCodesCollection "U", colLblCodes
    Set ffrmAddUpd.fcolLabelCodes = colLblCodes
    
     ffrmAddUpd.fAddEditMode = "Update_Record"
    
    ' Check if there are any records
    
    If fbAddOnly Then
        
        Exit Sub
    
    End If
    
    SetProductCodeCollection
    
    Set ffrmAddUpd.fcolProducts = fcolProductCodes
    Set colProduct = New Collection
    
    With ustgrdProductMaster
    
        'Add cell value of the grid to collection
        
        colProduct.Add Trim(.CellValue(.Row, cnProductCode)), _
                "PRODUCT_CODE"
        
        If Trim(.CellValue(.Row, cnProductDescription)) <> DELETED Then
            
            colProduct.Add Trim(.CellValue(.Row, cnProductDescription)), _
                "PRODUCT_DESC"
        
        Else
            
            colProduct.Add Right(Trim(.CellValue(.Row, cnAction)), _
                    Len(Trim(.CellValue(.Row, cnAction))) - 1), _
                    "PRODUCT_DESC"
        
        End If
        
        colProduct.Add Left(.CellValue(.Row, cnPrdTyp), 1), _
                "PRODUCT_TYPE"
        colProduct.Add .CellValue(.Row, cnDivisionCode), _
                "DIVISION_CODE"
        colProduct.Add Trim(.CellValue(.Row, cnLblNbr)), _
                "LABEL_NO"
        colProduct.Add Left(.CellValue(.Row, cnWgtTyp), 1), _
                "WGT_TYPE_CODE"
        colProduct.Add Trim(.CellValue(.Row, cnBxsPlt)), _
                "BOXES_PER_PALLET"
        colProduct.Add Trim(.CellValue(.Row, cnMinWgt)), _
                "MIN_BOX_WGT"
        colProduct.Add Trim(.CellValue(.Row, cnMaxWgt)), _
                "MAX_BOX_WGT"
        colProduct.Add Trim(.CellValue(.Row, cnTareWgt)), _
                "BOX_TARE_WEIGHT"
        colProduct.Add Trim(.CellValue(.Row, cnBlastInd)), _
                "BLAST_IND"
        colProduct.Add Trim(.CellValue(.Row, cnFrzDays)), _
                "FREEZE_DAYS"
        colProduct.Add Trim(.CellValue(.Row, cnBlstDays)), _
                "BLAST_DAYS"
        colProduct.Add Trim(.CellValue(.Row, cnLblLen)), _
                "LABEL_LENGTH"
        colProduct.Add Trim(.CellValue(.Row, cnStrPos)), _
                "LABEL_WGT_ST_POS"
        colProduct.Add Trim(.CellValue(.Row, cnWgtLen)), _
                "LABEL_WGT_LENGTH"
        colProduct.Add gsPlantCode, _
                "CUSTOMER_ID"
        colProduct.Add "S", _
                "CUSTOMER_TYPE"
        colProduct.Add Trim(.CellValue(.Row, cnComdCode)), _
                "GOVT_COMMODITY_CODE"
        colProduct.Add .CellValue(.Row, cnGovtLot), _
                "GOVT_LOT_IND"
        colProduct.Add Trim(.CellValue(.Row, cnMfgDate)), _
                "PACK_DATE_IND"
        
        'Added by TCS-Ragu on 17-Jun-2005
        'Start
        colProduct.Add Trim(.CellValue(.Row, cnMfgDatePos)), _
                "PACK_DATE_START"
        'End
        
        colProduct.Add Trim(.CellValue(.Row, cnLabelCodeChk)), _
                "CHECK_LABEL_IND"
        colProduct.Add Trim(.CellValue(.Row, cnLabelCodeStr)), _
                "PROD_LABEL_START"
        colProduct.Add Trim(.CellValue(.Row, cnLabelCodeLen)), _
                "PROD_LABEL_LEN"
        colProduct.Add .CellValue(.Row, cnBoxSerialInd), _
                "BOX_SERIAL_IND"
        colProduct.Add Left(.CellValue(.Row, cnProductGroupCode), 1), _
                "PRODUCT_GROUP_CODE"  'Added by TCS
        colProduct.Add Trim(.CellValue(.Row, cnProductCodeKey)), _
                "PRODUCT_CODE_KEY"
                                
        colProduct.Add "Y", "SECOND_IND" 'Added by TCS
                
        colProduct.Add Trim(.CellValue(.Row, cnXRef)), _
                "XREF_CODE"
                
        colProduct.Add .Row, _
                "ROW_INDEX"
    End With
    
    Set ffrmAddUpd.fcolProduct = colProduct
   
    ffrmAddUpd.Show vbModal
    
CleanUpAndExit:

    Exit Sub

ErrHandler:

    If Err.Number = 364 Then
    
        GoTo CleanUpAndExit
     
    Else
    
        'VB error
        
        MsgBox Err.Description, _
               vbInformation + vbOKOnly, _
               Me.Caption
               
    End If

End Sub

'******************************************************************************
'* Functional Description   :   Forms the collection of all existing
'*                              LabelCode+ ProductTyps combinations
'* Parameter Description    :   sMode - Called from Add or Update mode
'*                              colCollection - Collection to be populated
'* Return Type Description  :   None
'******************************************************************************

Private Function SetLabelCodesCollection(ByVal sMode As String, _
                        ByRef colCollection As Collection)

Dim lRowIndex       As Long

    On Error GoTo ErrHandler
    
    Set colCollection = New Collection
    
    With ustgrdProductMaster
    
        If sMode = "I" Then
            
            For lRowIndex = 1 To .Rows - 1
                
                colCollection.Add .CellValue(lRowIndex, cnLblNbr) & "|" _
                            & Left(.CellValue(lRowIndex, cnPrdTyp), 1), _
                            .CellValue(lRowIndex, cnLblNbr) & "|" _
                            & Left(.CellValue(lRowIndex, cnPrdTyp), 1)
            
            Next
            
        ElseIf sMode = "U" Then
            
            For lRowIndex = 1 To .Rows - 1
                
                If lRowIndex <> .Row Then
                    
                    colCollection.Add .CellValue(lRowIndex, cnLblNbr) & "|" _
                            & Left(.CellValue(lRowIndex, cnPrdTyp), 1), _
                            .CellValue(lRowIndex, cnLblNbr) & "|" _
                            & Left(.CellValue(lRowIndex, cnPrdTyp), 1)
                
                End If
            
            Next
        
        End If
        
    End With
    
CleanUpAndExit:
    
    Exit Function

ErrHandler:

    'Display VB error
    If Err.Number = 457 Then
        Resume Next
    Else
    
    gcolErrMsg.Add Err.Description
    giErrMsg = ShowErrorMsg("ProMstr013", Me.Caption, _
                               vbOKOnly, gcolErrMsg)
    GoTo CleanUpAndExit
    End If
End Function

'******************************************************************************
'* Functional Description   :   Prints ProductMaster Report
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Function PrintProductMasterReport() As Boolean
    
Dim objProduct      As Object
Dim rsReport        As ADODB.Recordset
Dim sStatusText     As String
Dim lResult         As Long
    
    On Error GoTo ErrHandler
    
        PrintProductMasterReport = True
        
        'Create object of ProductMaster Class in MasterFileFunctions
        
        Set objProduct = CreateObject("MasterFileFunctions.ProductMaster")
        
        'Set status bar and mouse pointer
        
        sStatusText = sbStatusBar.Panels(1).Text
        sbStatusBar.Panels(1).Text = "Loading From Database - Please Wait..."
        Me.MousePointer = vbHourglass
        DoEvents
        
        'Retrieve Product Master Report details
        
        lResult = objProduct.GetDetailsForProductMasterReport(gsPlantCode, _
                  fbCrossRefProductsOnly, rsReport)
        
        'Reset status bar and mouse pointer
                 
        Me.MousePointer = vbDefault
        sbStatusBar.Panels(1).Text = sStatusText
        
        If lResult <> 0 Then GoTo ErrHandler
        
        If rsReport.RecordCount = 0 Then
            
            'Display No data for report
         
            giErrMsg = ShowErrorMsg("ProMstr015", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
            PrintProductMasterReport = False
        
        Else
            
            sStatusText = sbStatusBar.Panels(1).Text
            sbStatusBar.Panels(1).Text = "Printing Report - Please Wait..."
            Me.MousePointer = vbHourglass
            DoEvents
            
            'Print report
            
            PrintReport "Product Master Report", _
                    "ProductMaster.rpt", _
                    rsReport, _
                    gbPrintWithoutPreview
        
        End If
        
        Me.MousePointer = vbDefault
        sbStatusBar.Panels(1).Text = sStatusText
            
CleanUpAndExit:

    gbPrintWithoutPreview = False
    Set objProduct = Nothing
    Set rsReport = Nothing

    Exit Function
    
ErrHandler:
    
    PrintProductMasterReport = False
    Me.MousePointer = vbDefault
    sbStatusBar.Panels(1).Text = sStatusText
    
    If lResult <> 0 Then
    
        'Display Server error
        
        gcolErrMsg.Add lResult
        giErrMsg = ShowErrorMsg("ProMstr014", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    Else
        
        'Display VB error
        
        gcolErrMsg.Add Err.Description
        giErrMsg = ShowErrorMsg("ProMstr006", Me.Caption, _
                                vbOKOnly, gcolErrMsg)
    End If

    GoTo CleanUpAndExit
    
End Function

'******************************************************************************
'* Functional Description   :   Display screen to add a new record
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuAdd_Click()

    If fsAccessRights = FUNC_MODIFY Then
        
        ShowAddScreen
    
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Display screen to add a new record
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuAddPopup_Click()
    
    mnuAdd_Click
    
End Sub

'******************************************************************************
'* Functional Description   :   Unloads form without updating DB
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuClose_Click()

    fbCloseWithoutUpdate = True
    Unload Me
    
End Sub

'******************************************************************************
'* Functional Description   :   Unloads form without updating DB
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCloseWithoutSave_Click()

    fbCloseWithoutUpdate = True
    Unload Me

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

'Added by TCS on 08-Sep-2005
Private Sub mnuInitProdSend_Click()
    mdi_InitProdSend.Show
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
'* Functional Description   : Calls Send Product Info to Scanner.
'* Parameter Description    : None.
'* Return Type Description  : None.
'******************************************************************************

Private Sub mnuProductToScanner_Click()

    Call SendProductInfoToScanner

End Sub

' Added by TCS
Private Sub mnuProductsNotSetUp_Click()
    
    ProductsNotSetupReportClick
    SetNoTopmost
    
End Sub

'******************************************************************************
'* Functional Description   :   Unloads form without updating DB
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuSaveAndClose_Click()

    fbCloseWithoutUpdate = False
    Unload Me

End Sub

'******************************************************************************
'* Functional Description   :   Implements Cut functionality in the menu
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCut_Click()

    SendKeys "^X"
    
End Sub

'******************************************************************************
'* Functional Description   :   Implements Copy functionality in the menu
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCopy_Click()

    SendKeys "^C"
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the inquire screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuInquirePopup_Click()
    
    mnuInquire_Click
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the inquire screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuInquirePopupBrowse_Click()

    mnuInquire_Click

End Sub

'******************************************************************************
'* Functional Description   :   Shows the inquire screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuInquirePopupNoAccess_Click()
    
    mnuInquire_Click
    
End Sub

'******************************************************************************
'* Functional Description   :   Implements Paste functionality in the menu
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************
Private Sub mnuPaste_Click()

    SendKeys "^V"
    
End Sub

'******************************************************************************
'* Functional Description   :   Prompts the user whether to print
'*                              CrossRef Products only.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuPrintPM_Click()

Dim lAnswer As Long
    
    ' Check if there are any records
    
    If fbAddOnly Then
    
        Exit Sub
        
    End If
    
    If UpdateProductPlant Then
    
        If LoadProductDetails <> 0 Then GoTo ErrHandler

        lAnswer = ShowErrorMsg("ProMstr007", Me.Caption, vbYesNo)

        If lAnswer = vbYes Then
            
            fbCrossRefProductsOnly = True
        
        ElseIf lAnswer = vbNo Then
            
            fbCrossRefProductsOnly = False
        
        End If
        
        ustgrdProductMaster.SetFocus
        
        If Not PrintProductMasterReport Then GoTo ErrHandler
        SetNoTopmost
    End If
    
    Exit Sub
    
ErrHandler:

    Exit Sub

End Sub

'******************************************************************************
'* Functional Description   :   Exits the application
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuExit_Click()

    Unload Me
    Unload frmMainMenu
    
End Sub



'******************************************************************************
'* Functional Description   :   Display screen to Update a record
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuUpdate_Click()

    If fsAccessRights = FUNC_MODIFY Then
        
        ShowUpdateScreen
    
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Display screen to Update a record
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuUpdatePopUp_Click()
    
    mnuUpdate_Click
    
End Sub

'******************************************************************************
'* Functional Description   :   Sets the visibility of the Status bar as per
'*                              the menu selection
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
'* Functional Description   :   Sets the visibility of the Tool bar as per
'*                              the menu selection
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
'* Functional Description   :   Arranges the windows in a cascading fashion.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWindowCascade_Click()

    frmMainMenu.Arrange vbCascade
    
End Sub

'******************************************************************************
'* Functional Description   :   Tiles the windows horizontally.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWindowTileHorizontal_Click()

    frmMainMenu.Arrange vbTileHorizontal
    
End Sub

'******************************************************************************
'* Functional Description   :   Tiles the windows vertically.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWindowTileVertical_Click()

    frmMainMenu.Arrange vbTileVertical
    
End Sub



'******************************************************************************
'* Functional Description   :   Handles the toolbar button clicks.
'* Parameter Description    :   Button - The button that the user has clicked.
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
        Case biPrint
            gbPrintWithoutPreview = True
            mnuPrintPM_Click
        Case biInquire
            mnuInquire_Click
        Case biAdd
            ShowAddScreen
        Case biUpdate
            ShowUpdateScreen
    End Select
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the help contents.
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
'* Functional Description   :   Shows the inquire screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuInquire_Click()
    
    If fbAddOnly Then
        
        Exit Sub
    
    End If
    
    Set ffrmInquire = New frmProductCodeInquire
    ffrmInquire.Show vbModal

End Sub



'******************************************************************************
'* Functional Description   :   Gives description about the application.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuHelpAbout_Click()

    frmAbout.Show vbModal, frmMainMenu
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Commodities Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCommoditiesUpdate_Click()

    CommoditiesUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Customer Master Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerMasterUpdate_Click()

    CustomerMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Memo Master Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuMemoMasterUpdate_Click()

    MemoMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Plant Master Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuPlantMasterUpdate_Click()

    PlantMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Rate Code Master Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuRateCodeMasterUpdate_Click()

    RateCodeMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Product Master Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuProductMasterUpdate_Click()

    ProductMasterUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Order Shipping Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOrderShippingUpdate_Click()

    OrderShippingUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Daily Production Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuDailyProductionUpdate_Click()

    DailyProductionUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Detail Inventory Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuDetailInventoryUpdate_Click()

    DetailInventoryUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Pallet Movement screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuPalletMovement_Click()

    PalletMovementClick
    
End Sub


'******************************************************************************
'* Functional Description   :   Shows the main Freezer Inventory Browse screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuFreezerInventoryBrowse_Click()

    FreezerInventoryBrowseClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Order Summary Browse screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOrderSummaryBrowse_Click()

    OrderSummaryBrowseClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Order Information Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOrderInformationUpdate_Click()

    OrderInformationUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the main Customer Storage Browse screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerStorageBrowse_Click()

    CustomerStorageBrowseClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the Adjustment Log Browse screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuAdjustmentLogBrowse_Click()

    AdjustmentLogBrowseClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the Customer Storage Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerStorageUpdate_Click()

    CustomerStorageUpdateClick
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the Customer Charges Update screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerChargesUpdate_Click()

    CustomerChargesUpdateClick
    
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



'******************************************************************************
'* Functional Description   :   Shows the Blast Cell Unloading Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuBlastCellUnloadingReport_Click()
    
    BlastCellUnloadingReportClick
    SetNoTopmost
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the Daily Production Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuDailyProductionReport_Click()
    
    DailyProductionReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Detail Production Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuDetailProductionReport_Click()
    
    DetailProductionReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Code By Code Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCodebyCodeReport_Click()
    
    CodebyCodeReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Aged Product Inventory Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuAgedProductInventoryReport_Click()
    
    AgedProductInventoryReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Freezer Inventory Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuFreezerInventoryReport_Click()
    
    FreezerInventoryReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Weekly Freezer Recap Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWeeklyFreezerRecapReport_Click()
    
    WeeklyFreezerRecapReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows message box that Cold Storage
'*                              Commodity report printing is ON.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuColdStorageCommodityReport_Click()
    
    ColdStorageCommodityReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Code By Code Vs Inventory Report
'*                              screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCodeVsInventoryReport_Click()
    
    CodeVsInventoryReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the first screen of Customer Master report.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerMasterReport_Click()
    
    CustomerMasterReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Customer Storage Due Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerStorageDueReport_Click()
    
    CustomerStorageDueReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Customer Storage Detail Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerStorageDetailReport_Click()
    
    CustomerStorageDetailReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Customer Charges Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerChargesReport_Click()
    
    CustomerChargesReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Weekly Sales Summary Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuWeeklysalesSummaryReport_Click()
    
    WeeklySalesSummaryReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the message box that Customer Inventory
'*                              By Product Report printing is ON.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerInventoryByProductReport_Click()
    
    CustomerInventoryByProductReportClick
    SetNoTopmost
    
End Sub

'******************************************************************************
'* Functional Description   :   Shows the Customer Storage Recap Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerStorageRecapReport_Click()

    CustomerStorageRecapReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Customer Warehouse Receipts Report
'*                              screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuCustomerWarehouseReceipts_Click()
    
    CustomerWarehouseReceiptsClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Order Manifest Detail Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

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

'******************************************************************************
'* Functional Description   :   Shows the Order Pullsheet Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOrderPullSheetReport_Click()
    
    OrderPullSheetReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Order Summary Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOrderSummaryReport_Click()
    
    OrderSummaryReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Product Recall Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuProductRecallReport_Click()
    
    ProductRecallReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Outside Customer BOL Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuOutsideCustomerBOL_Click()
    
    OutsideCustomerBOLClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Inbound/Outbound Pallet Report screen
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuInboundOutboundPalletReport_Click()
    
    InboundOutboundPalletReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the Pallet Detail By Order Report screen.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuPalletDetailByOrderReport_Click()
    
    PalletDetailByOrderReportClick
    SetNoTopmost

End Sub

'******************************************************************************
'* Functional Description   :   Shows the screen to select the plant id for an
'*                              admin user.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub mnuSelectPlant_Click()
    
    SelectPlant
    
End Sub

'******************************************************************************
'* Functional Description   :   Display screen to update current row
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ustgrdProductMaster_DblClick()
    
    If fsAccessRights = FUNC_MODIFY Then
        
        ShowUpdateScreen
    
    End If

End Sub

'******************************************************************************
'* Functional Description   :   Show the ProductDesc if it is marked for delete
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ustgrdProductMaster_EnterCell()

    With ustgrdProductMaster
        
        If .Column = cnProductDescription Then
            
            ShowDeletedDescription ustgrdProductMaster, _
                                    cnAction, cnProductDescription
        
        End If
    
    End With
    
End Sub

'******************************************************************************
'* Functional Description   :   Marks a row for delete on Delete key press.
'*                              Popup edit screen on Enter Key Press
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ustgrdProductMaster_KeyDown(KeyCode As Integer, Shift As Integer)

    If fsAccessRights = FUNC_MODIFY Then
        
        With ustgrdProductMaster
            
            Select Case KeyCode
                
                Case vbKeyDelete
                    ProcessDeleteRow ustgrdProductMaster, cnAction, _
                                cnProductDescription
                                
                Case vbKeyReturn
                    ShowUpdateScreen
            
            End Select
        
        End With
    
    End If
    
End Sub

''******************************************************************************
'* Functional Description   :   Handles Esc key
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ustgrdProductMaster_KeyPress(KeyAscii As Integer)
            
    If KeyAscii = vbKeyEscape Then
        
        Unload Me
    
    End If
    
End Sub

'******************************************************************************
'* Functional Description   :   Hide the Product Desc if marked for delete
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'******************************************************************************

Private Sub ustgrdProductMaster_LeaveCell()

    With ustgrdProductMaster
        
        If .Column = cnProductDescription Then
            
            HideDeletedDescription ustgrdProductMaster, _
                                    cnAction, cnProductDescription
        
        End If
    
    End With

End Sub

'*******************************************************************************
'* Functional Description   :   Highlights row corr to error selected by user
'* Parameter Description    :   sKey - The key fields separated by " | "
'* Return Type Description  :   None.
'*******************************************************************************

Public Sub ErrorSelected(ByVal sKey As String)

Dim lRowIndex As Long
    
    On Error GoTo ErrHandler
    
    With ustgrdProductMaster
        
        For lRowIndex = 1 To .Rows - 1
            
            If .CellValue(lRowIndex, cnProductCode) = _
                            sKey Then
                
                .Row = lRowIndex
                .TopRow = lRowIndex
                .Column = cnProductCode
                .SetFocus
                Exit Sub
                
            End If
            
        Next
        
    End With
    
    Exit Sub

ErrHandler:

    'VB Error
    
    gcolErrMsg.Add Err.Description
    giErrMsg = ShowErrorMsg("ProMstr009", Me.Caption, _
                                vbOKOnly, gcolErrMsg)

End Sub

'*******************************************************************************
'* Functional Description   :   Marks a row as deleted or removes the marking.
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub mnuDelete_Click()

    ProcessDeleteRow ustgrdProductMaster, cnAction, cnProductDescription

End Sub

'*******************************************************************************
'* Functional Description   :   Handles Rightclick menu
'* Parameter Description    :   None.
'* Return Type Description  :   None.
'*******************************************************************************

Private Sub ustgrdProductMaster_MouseDown(Button As Integer, _
                                          Shift As Integer, _
                                          X As Single, _
                                          Y As Single)

    If fsAccessRights = FUNC_MODIFY Then
        
        With ustgrdProductMaster
            
            If Button = vbRightButton And .GridEditMode = Cell Then
                
                If Not fcolErrors Is Nothing Then
                    
                    mnuError.Visible = True
                
                Else: mnuError.Visible = False
                
                End If
                
                PopupMenu mnuPopup
            
            End If
        
        End With
    
    ElseIf fsAccessRights = FUNC_BROWSE_0NLY Then
        
        With ustgrdProductMaster
            
            If Button = vbRightButton And .GridEditMode = Cell Then
                
                PopupMenu mnuNoAccessPopup
            
            End If
        
        End With
    
    End If
    
End Sub


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
'Added by TCS to Get Filter String
Private Function IsFilterAvailable(ByRef SFltrval As String) As Boolean

SFltrval = ""

If Not sFltrProdCode = "" Then
    SFltrval = "PRODUCT_CODE Like '" & sFltrProdCode & "%'"
End If

If Not sFltrDivCode = "" Then
    SFltrval = IIf(SFltrval = "", "", SFltrval & " And ") & "DIVISION_CODE='" & sFltrDivCode & "'"
End If

If Not sFltrProdType = "" Then
    SFltrval = IIf(SFltrval = "", "", SFltrval & " And ") & "PRODUCT_TYPE='" & sFltrProdType & "'"
End If

If Not sFltrLabelNo = "" Then
    SFltrval = IIf(SFltrval = "", "", SFltrval & " And ") & "LABEL_NO='" & sFltrLabelNo & "'"
End If

If Not SFltrval = "" Then IsFilterAvailable = True

'Remove Parameter Values
sFltrProdCode = ""
sFltrDivCode = ""
sFltrProdType = ""
sFltrLabelNo = ""


End Function

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


