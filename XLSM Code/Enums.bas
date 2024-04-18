Attribute VB_Name = "Enums"
Option Explicit

' --------------------------------------------
' Project: Attack on Europe 2022
' Made by: Matthew Borcz
' GH Page: https://github.com/MBrcz
' --------------------------------------------

' ***
' This module shall hold all the relevant enums in order to complete the project.
' They is basically mapping of the tables with Keys and Columns
' ***

' -- STARTS TABLES --
' Enums responsible for getting the positions of the tables.
Public Enum StartTables ' (Keys)
    WorksheetNames = 1
    WorkbookMode = 2
    SlicerCall = 3
    WorkbookScrollbars = 4
    OtherParameters = 5
    FamilyParameters = 6
    MilitaryBranchParameter = 7
    FunctionCaller = 8
    ItemsSearcherParameters = 9
    ItemsOtherParameters = 10
    BackendParamaters = 11
End Enum

Public Enum StartTablesParam ' (Columns)
    Reference = 2
End Enum
' -- END STARTS TABLES --

Public Enum WorkbookMode
    CINEMATIC = 1
    NORMAL = 2
    INTRO = 3
End Enum

Public Enum WorkbookModeParam
    WbModeValue = 1
End Enum

' -- SLICERS --
' Enums responsible for managing and handling slicers in the project.
Public Enum SlicerCall ' (Keys)
    backCurrentDate = 1
    backPreviousDate = 2
    milBranch = 200
    familyEquipmentType = 300
    familyFamilyName = 301
    itemsBranch = 400
    itemsEquipmentType = 401
    itemsFamilyName = 402
    itemsFullName = 403
    itemsProductionCountry = 404
    itemsFateName = 405
    itemsSideOfConflict = 406
    
    fatesBranch = 500
    fatesEquipmentType = 501
    fatesFamilyName = 502
    fatesFullName = 503
    fatesProductionCountry = 504
    fatesSideOfConflict = 505
    fatesInternalDate = 506
End Enum

Public Enum SlicerCallParam ' (Columns)
    Name = 1
    Table = 2
    Column = 3
    CurrentValue = 4
    SignatureName = 5
End Enum
' -- END SLICERS --

' -- SCROLLBARS --
' Those enums contains data about the scrollbars used in the project.
Public Enum ScrollBar ' Keys
    startTop = 100
    startMid = 101
    startBot = 102
    itemsOne = 400
End Enum

Public Enum ScrollBarParam ' Columns
    Current = 2
    Min = 3
    Max = 4
    ToShow = 5
    IncrChange = 6
    ScrName = 7
End Enum
' -- END SCROLLBARS --

' -- WORKSHEETS --
' This enums keeps all the worksheet names / parameters
Public Enum ProjectSheets ' (Key, getter only)
    Backend = 1
    Introduction = 2
    StartingPage = 100
    MilitaryOverwiev = 200
    MilitaryFamilies = 300
    ItemsSearcher = 400
    ItemsFates = 500
End Enum

Public Enum ProjectSheetsParam ' (Columns)
    SheetName = 1
End Enum
' -- END WORKSHEETS --

' -- OTHER PARAM --
Public Enum OtherParamNum ' Keys
    IntroExit = 1
    IntroRefresh = 2
    IntroModes = 3
    
    FamilySlicerTextBoxName = 300
    FamilyEqTypeToPass = 301
    FamilyEquipmentTypeStartRange = 302
    FamilyEquipmentTypeCellRange = 303
    FamilyEquipmentTypeSlicer = 304
End Enum

Public Enum OtherParamName ' Columns
    ParamValue = 1
End Enum
' -- END OTHER PARAM --

' -- MILITARY BRANCH --
' This enum gives the data about possible parameters related to the branches slicer.
' RELATES TO: Military Overwiev slicer.
Public Enum MilOverwievBranch
    Air = 1
    Land = 2
    Naval = 3
End Enum

Public Enum MilOverwievBranchParam
    MilValue = 1
End Enum

' -- END MILITARY BRANCH

' -- FAMILY PARAM --
' This enum stores all data necessary to manipulate the page with family military.
' Family object is an textbox and image with slicer in the left of the page.
Public Enum FamilyObject
    First = 1
    Second = 2
    Third = 3
    Fourth = 4
End Enum

Public Enum FamilyObjectParam
    ImageRange = 1
    TextBoxRange = 2
    Description = 3
    ImageURLPath = 4
End Enum
' -- END FAMILY PARAM --

' -- ITEMS SEARCHER PARAMETERS --
' -- Call only from items searcher.
' -- They should have in common numbers related to their slicers. --
Public Enum ItemBoxes
    ItemBoxBranch = 400
    ItemBoxType = 401
    ItemBoxFamily = 402
    ItemBoxFullName = 403
    ItemBoxCountryFullName = 404
    ItemBoxFateName = 405
    
    ItemFatesBranch = 500
    ItemFatesEquipmentType = 501
    ItemFatesFamilyNames = 502
    ItemFatesFullName = 503
    ItemFatesCountryFullName = 504
End Enum

Public Enum ItemBoxesParam
    SlicerBoxRange = 1
    SlicerToApply = 2
    SlicerSourceRangeName = 3
End Enum

Public Enum ItemSearcherBookmark
    ItemBundeswehr = 4
    ItemPolishLandForces = 5
    ItemUnitedStatesAirForce = 6
    ItemTurkishAirFoces = 7
End Enum

' -- END ITEMS SEARCHER PARAMETERS --

' -- ITEM OTHERS PARAMETERS --
Public Enum ItemOther
    ' Item Searcher
    ItemSideOfConflictRussia = 1
    ItemSideOfConflictUkraine = 2
    ItemSideOfConflictSignature = 3
    ItemSearcherFlagStartRange = 400
    ItemSearcherFlagSingature = 401
    ItemSearcherCorrectnessTable = 402
    ItemSearcherPortrairSignature = 403
    ItemSearcherPortrairRange = 404
    ItemSearcherPortrairTextBoxName = 405
    ItemSearcherPortrairPath = 406
    ItemSearcherTextBoxContent = 407
    ItemSearcherBookmarkSignature = 408
    ItemSearcherSearchButtonSignature = 409
    ItemSearcherRestartFiltersSignature = 410
    
    ' ItemFates
    ItemFatesCorrectnessTable = 500
    ItemFatesPortrairSignature = 501
    ItemFatesPortrairRange = 502
    ItemFatesPortrairTextBoxName = 503
    ItemFatesPortrairPath = 504
    ItemFatesTextBoxContent = 505
    ItemFatesBookmarkSignature = 506
    ItemFatesSearchButtonSignature = 507
    ItemFatesRestartFiltersSignature = 508
End Enum

Public Enum ItemOtherParam
    ItemOtherValue = 1
End Enum
' -- END ITEM OTHER PARAMETERS --

' -- BACKEND / CONTROL PANEL --
Public Enum Backend
    
    ' Dates
    RangeDatesAsNums = 1
    RangeDatesAsDates = 2
    CurrentDateAsText = 3
    CurrentDateAsInteger = 4
    PreviousDateAsInteger = 5
    LastPossibleDateAsInteger = 6
    
    ' Objects Signatures
    ChangeDateButtonSignature = 10
    ArrowMoveSignature = 11
    FlagSignature = 12
    BackendLogo = 13
    
    ' Flags Sources
    RussianSource = 100
    UkrainianSource = 200
End Enum

Public Enum BackendParam
    BackendValue = 1
End Enum
' -- END BACKEND / CONTROL PANEL --

' -- FUNCTION BINDER --
' This enum is responsible for binding the names of the function names
' to the convinient numbers.
' The functions here have 2 prefixes:
' a) Call - means that the function is wrapped by ExecuteProcedure function,
' b) Func - means that the function is not wrapped by ExecudeProcedure
'
' ExecuteProcedure sub changes workbook ScreenUpdate and EnableEvents properties
' before egzecution of the code.
' ------------------------------------------------------------------------------
Public Enum FunctionCaller
    ' Workbook
    FFuncZoomToVisibleCells = 1
    FFuncPreventScrollBarFromSwitching = 2
    
    ' BACKEND
    S1FuncChangeGlobalDate = 50
    S1FuncChangeGlobalDateByNum = 51
    S1FuncMoveToSheetViaArrow = 52
    S1FuncCheckTheSource = 53
    
    S1CallChangeGlobalDate = 60
    S1CallMoveToSheetViaArrow = 61
    S1CallCheckTheSource = 62
    
    S1FuncUpdateControlPanelObjects = 99
    
    ' Item Start
    SStartFuncWorksheetOpen = 299
    
    ' Military Overwiev
    S2FuncSetOverwievSlicer = 300
    S2CallSetOverwievSlicer = 301
    S2FuncWorksheetOpen = 399
    
    ' Military Families
    S3FuncSetFamilySlicer = 400
    S3CallSetFamilySlicer = 401
    S3FuncSetEquipmentTypeSlicer = 402
    S3CallSetEquipmentTypeSlicer = 403
    S3HideValidationBox = 404
    S3FuncWorksheetOpen = 499
    
    ' Item Searcher
    S4FuncInvokeBookmarks = 500
    S4CallInvokeBookmarks = 501
    S4FuncSearchQuery = 502
    S4CallSearchQuery = 503
    S4FuncRestartFilters = 504
    S4CallRestartFitlers = 505
    S4FuncSetSideOfConflictSlicer = 506
    S4CallSetSideOfConflictSlicer = 507
    S4FuncUpdateScrollBar = 508
    S4CallUpdateScrollBar = 509
    S4FuncWorksheetOpen = 599
    
    ' Item Fates
    S5FuncInvokeBookmarks = 600
    S5CallInvokeBookmarks = 601
    S5FuncSearchQuery = 602
    S5CallSearchQuery = 603
    S5FuncRestartFilters = 604
    S5CallRestartFitlers = 605
    S5FuncSetSideOfConflictSlicer = 606
    S5CallSetSideOfConflictSlicer = 607
    S5FuncWorksheetOpen = 699
    
    ' Introduction
    S6FuncWorksheetOpen = 799
    
    ' Other
    WbFuncRefreshSource = 1000
    WbCallRefreshSource = 1001
    WbfuncChangeWorkbookMode = 1002
    WbCallChangeWorkbookMode = 1003
    S7FuncExitApp = 1004
    S7CallExitApp = 1005
    
End Enum

Public Enum FunctionCallerParam
    FuncValue = 1
End Enum
' -- END FUNCTION BINDER --
