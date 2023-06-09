Attribute VB_Name = "VariableDeclarations"
    ' I101 STARTUP Flags
    Global bolSysTrayActive As Boolean
    Global bolShowMenu As Boolean
    Global bolDebug As Boolean
    Global bolNoExit As Boolean
    Global bolForceDBUpdate As Boolean
    Global bolForceLogin As Boolean
    
    Global txtHoldField As String
    Global txtOutputLine As String
        
    Global txtImageFileName As String
    Global txtImageFilePath As String
    Global txtOutputFile As String
    Global txtOutputFilePath As String
    Global txtOutputTempFilePath As String
    Global txtImageBatch As String
    Global txtActionBeforeError As String
    
    ' txtCurrentModule allows tracking which Form is currently Active]
    '  so that we can create custom features for viewer and other actions
    '  Usage:          txtCurrentModule = "frmIndex"
    Global txtCurrentModule As String
    
    Global intFieldIndex As Integer
    Global intFileIndex As Integer
    Global txtIndexPathName As String

    Global blnSocSecInvalid As Boolean
    Global blnContinueLastRun As Boolean
    Global blnFTPError As Boolean
    Global blnBatchError As Boolean
    Global blnExitBatch As Boolean
    
    Global blnBookMark As Boolean
    Global intBookMarkBegin As Integer
    Global intBookMarkEnd As Integer
    
    Global Const blnNoPrompt As Boolean = True
    
    Global txtHoldFileName    As String
    Global Const txtQuestionable As String = "*??????????*"
    Global Const txtSeparator As String = "*SEPARATOR SHEET*"
    Global Const txtDoNotFile As String = "*DO NOT FILE*"
    Global gAutoAdvanceOnSeparator As String
    Global gSetUserAsBatchOwnerOnSPLIT As String
    '
    Global gBypassBatchAutoSelect As Boolean
    Global gOpenBatchInReadOnlyMode As Boolean
    Global gNoLookupFieldsAvailable As Boolean
    
    Global strFIELDAFTERCLICK As String

    Global intIndex As Integer
    Global result As Long
    Global blnDateError As Boolean
    Global blnCommittingBatchPages As Boolean
    Global blnCommitSelectedBatches As Boolean
    Global blnCommitSelectedBatchesWait As Boolean
    Global blnBarcodeSelectedBatchesWait As Boolean
    
    Global gYesNo As Integer
    
    '*** Set Defaults for Registry Entries ***
    Global Const RegAppname As String = "Imaging101Client"
    Global Const RegSectionName As String = "Settings"
    Global RegFileName As String


    '*** Set Defaults for Registry Entries ***
    
    Global RegEcaptureBatchListConnectionType As String
    Global RegEcaptureBatchListConnectionString As String
    
    Global RegImaging101BatchListConnectionType As String
    Global RegImaging101BatchListConnectionString As String

    Global RegImaging101ConnectionType As String
    Global RegImaging101ConnectionString As String
    
    Global RegDocTypeListConnectionType As String
    Global RegDocTypeListConnectionString As String

    Global RegLookupListConnectionType As String
    Global RegLookupListConnectionString As String
    
    Global RegLookupListTableName As String
    Global RegLookupDBTableIsOnSQLServer As String
    Global RegLookupListWhereClause As String
    
'    Global RegTTCConnectionType As String
'    Global RegTTCConnectionString As String
    
    Global RegImaging101OCRConnectionType As String
    Global RegImaging101OCRConnectionString As String
    
    Global RegRootDirToStoreObjects As String
    Global RegRootDirectoryPathForImageAnnotations
    Global RegBatchRootDir As String
    
    Global RegApplicationBatchNameDelimiter As String
    
    '*** Set GLOBAL SECURITY VARIABLES ***
    Global gsecSecurityRECID As Double
    Global gsecUserID As String
    Global gsecUserName As String
    Global gsecPassword As String
    
    Global gsecDocumentGroup As String
    Global gsecUserGroups As String
    
    Global gsecRightsAdminSystem As String
    Global gsecRightsAdminApplication As String
    Global gsecRightsRetrieveImages As String
    Global gsecRightsBatchScan As String
    Global gsecRightsBatchIndex As String
    Global gsecRightsBatchAdministration As String
    Global gsecRightsBatchView As String
    Global gsecRightsBatchCommit As String
    Global gsecRightsBatchRoute As String
    Global gsecRightsBatchChangeOrder As String
    Global gsecBatchQueueNotificationFrequency As String
    Global gsecRightsImportFromFile As String
    Global gsecRightsImportFromEcapture As String
    Global gsecRightsDeleteDocuments As String
    Global gsecRightsModifyIndexes As String
    Global gsecRightsDeleteBatches As String
    Global gsecRightsSendMail As String
    Global gsecRightsLaunchDoc As String
    Global gsecRightsPrint As String
    Global gsecRightsAnnotate As String
    Global gsecRightsThumbnails As String
    Global gsecRightsScannerSettings As String
    Global gsecRightsDocPackage As String
    Global gsecRightsExport As String
    
    Global gsecBatchMode As String
    Global gsecUserSupervisor As String
    Global gsecBatchListOrder As String
    Global gsecBatchDefaultQueue As String
    Global gsecBatchDefaultApplication As String
    
    Global gsecViewResetImagesOnFind As String
    Global gsecAllowModificationOfOrigDocs As String
    
    Global gsecAdvancedSearch As String
    
    Global gsecRightsEditSearchTemplates As String
    
    '*** Set Global Site Information Variables
    Global gsecSiteInformationClientShort As String
    Global gsecSiteInformationClientLong As String
    Global gsecSiteInformationLicenseCode As String
    
    Global strRunDateField As String
    Global strRunDate As String
    
    'Global Array containing the DetailRECID of Images Displayed in the Viewer
    'this is to check if a specific DetailRECID item is already open
    ' so we don't open it again.
    Global arrDisplayedPagesRetrieve()
    Global arrDisplayedPagesIndex()
    
    ' These global Variables will track each Child Form to allow us to select and activate
    '  it such as when the document / page is already open'
    '   Essentially, they create an ARRAYS of NEW Child Forms
    Public gFormArrayRetrieve() As New ChildForm1
    Public gFormArrayIndex() As New ChildForm1

    ' These global Constants allow us to identify the Module that called the
    '  Show and Remove images in the MainMDIForm Viewer
    Global Const gI101ModuleIndex As Integer = 0
    Global Const gI101ModuleRetrieve As Integer = 1

    'Moved to here from Imaging101ScanMainPix
    Global bolCancelPendingXfers As Boolean
    Global bolRasterizingDocument As Boolean
    Global bolObjectLaunched As Boolean
    Global bolAIM_Command As Boolean
    Global bolAIM_Command_AddFile As Boolean
    Global bolAnnotationAdded As Boolean
    Global bolBatchScanningModule As Boolean

    Global gstrI101AIM_Module As String

    
    Global gProcessorID As String
    Global bolBarcodeLicenseValidated As Boolean
    
    'Moved here from frmIndex
    Global bolIndexFormLoadComplete As Boolean

    'To allow the Warning about Editing Original documents
    ' to display only the FIRST time for EACH Login.
    Global bolAllowModificationOfOrigDocsMessageDisplayed As Boolean


    ' Enhance Batch Security
    Global gsecRightsBatchFindRestricted  As String
    Global gsecRightsBatchFindRestrictToQueue  As String
    Global gsecRightsBatchFindRestrictToOwner  As String
    Global gsecRightsBatchChangeQueue  As String
    Global gsecRightsBatchChangeOwner  As String
    
    Global gsecRightsBatchAllowDocTypeEdit As String
    
    Global bolErrorOccured As Boolean
    

   'Retrieval Export Options
    Global gstrExportFormat As String
    
    'Move Documents
    Global bolWaitForMoveCommand As Boolean

    
    'Moved here from frmImaging101Search
    Global bolSearchFormLoadComplete As Boolean

    '
    Global bolEnableSearchTemplates As Boolean
    
    Global bolLogOpenedDocuments As Boolean
