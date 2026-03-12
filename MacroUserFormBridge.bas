Attribute VB_Name = "MacroUserFormBridge"
Option Explicit

' ============================================================
' UserForm 連携ブリッジ
' ------------------------------------------------------------
' 目的:
' - UserForm 側のコードを簡潔にする
' - 4モジュールの UiOptions へ値を詰める処理を共通化する
' ============================================================

Public Function GetEscapeOnlyAValueRowFillTargetOptions() As Variant
    GetEscapeOnlyAValueRowFillTargetOptions = Array("None", "Left", "Right", "Both")
End Function

Public Sub RunBetaEvidenceFromForm( _
    ByVal sourceWorkbookPath As String, _
    ByVal inputFileName As String, _
    ByVal useSlotHeight As Boolean, _
    ByVal slotHeight As Long, _
    ByVal useOutputSheetFilter As Boolean, _
    ByVal outputSheetFilterText As String, _
    ByVal topBorderEnabled As Boolean, _
    ByVal excludeOutputSheetByPatternEnabled As Boolean, _
    ByVal excludedOutputSheetNamePatterns As String, _
    ByVal skipGrayFilledSourceCellEnabled As Boolean, _
    ByVal sourceSkipFillColorHexCodes As String, _
    ByVal rightBorderEnabled As Boolean, _
    ByVal clearOldHeaderTextEnabled As Boolean, _
    ByVal evidenceNewSideColCount As Long, _
    ByVal matchOldSideColCountToNewSide As Boolean)

    Dim options As BetaEvidenceUiOptions
    BetaEvidenceGenerator.InitializeBetaEvidenceUiOptionsForForm options

    options.sourceWorkbookPath = Trim$(sourceWorkbookPath)
    options.inputFileName = Trim$(inputFileName)

    options.useSlotHeight = useSlotHeight
    If useSlotHeight And slotHeight > 0 Then
        options.slotHeight = slotHeight
    End If

    options.useOutputSheetFilter = useOutputSheetFilter
    options.outputSheetFilterText = Trim$(outputSheetFilterText)

    options.OverrideTopBorderEnabled = True
    options.topBorderEnabled = topBorderEnabled

    options.OverrideExcludeOutputSheetByPatternEnabled = True
    options.excludeOutputSheetByPatternEnabled = excludeOutputSheetByPatternEnabled
    options.UseExcludedOutputSheetNamePatterns = True
    options.excludedOutputSheetNamePatterns = Trim$(excludedOutputSheetNamePatterns)

    options.OverrideSkipGrayFilledSourceCellEnabled = True
    options.skipGrayFilledSourceCellEnabled = skipGrayFilledSourceCellEnabled
    options.UseSourceSkipFillColorHexCodes = True
    options.sourceSkipFillColorHexCodes = Trim$(sourceSkipFillColorHexCodes)

    options.OverrideRightBorderEnabled = True
    options.rightBorderEnabled = rightBorderEnabled

    options.OverrideClearOldHeaderTextEnabled = True
    options.clearOldHeaderTextEnabled = clearOldHeaderTextEnabled

    options.UseEvidenceNewSideColCount = True
    If evidenceNewSideColCount > 0 Then
        options.evidenceNewSideColCount = evidenceNewSideColCount
    End If

    options.UseEvidenceColumnLayoutScope = True
    If matchOldSideColCountToNewSide Then
        options.evidenceColumnLayoutScope = "Both"
    Else
        options.evidenceColumnLayoutScope = "NewOnly"
    End If

    If Len(options.sourceWorkbookPath) = 0 Or Len(options.inputFileName) = 0 Then
        MsgBox "参照元ファイルパスと入力ファイル名は必須です。", vbExclamation
        Exit Sub
    End If

    BetaEvidenceGenerator.RunMainWithUiOptions options
End Sub

Public Sub RunBetaTestCaseFromForm( _
    ByVal featureId As String, _
    ByVal useOutputPath As Boolean, _
    ByVal outputPath As String)

    Dim options As BetaTestCaseUiOptions
    options = BetaTestCaseGenerator.CreateBetaTestCaseUiOptionsForForm()

    options.featureId = Trim$(featureId)
    options.useOutputPath = useOutputPath
    options.outputPath = Trim$(outputPath)

    If Len(options.featureId) = 0 Then
        MsgBox "機能連番は必須です。", vbExclamation
        Exit Sub
    End If

    BetaTestCaseGenerator.RunMainWithUiOptions options
End Sub

Public Sub RunConditionalBranchCheckerFromForm( _
    ByVal featureName As String, _
    ByVal workbookPath As String, _
    ByVal leadingFunctionStartsFromB1 As Boolean, _
    ByVal markNonFunctionLineWithDash As Boolean, _
    ByVal markFillEnabled As Boolean, _
    ByVal markFillColorHex As String)

    Dim options As ConditionalBranchCheckerUiOptions
    options = ConditionalBranchChecker.CreateConditionalBranchCheckerUiOptionsForForm()

    options.featureName = Trim$(featureName)
    options.workbookPath = Trim$(workbookPath)
    options.OverrideLeadingFunctionStartsFromB1 = True
    options.leadingFunctionStartsFromB1 = leadingFunctionStartsFromB1
    options.OverrideMarkNonFunctionLineWithDash = True
    options.markNonFunctionLineWithDash = markNonFunctionLineWithDash
    options.OverrideMarkFillEnabled = True
    options.markFillEnabled = markFillEnabled
    options.UseMarkFillColorHex = True
    options.markFillColorHex = Trim$(markFillColorHex)

    If Len(options.featureName) = 0 Or Len(options.workbookPath) = 0 Then
        MsgBox "機能名と対象ブックパスは必須です。", vbExclamation
        Exit Sub
    End If

    ConditionalBranchChecker.RunMainWithUiOptions options
End Sub

Public Sub RunEscapePartsMarkingFromForm( _
    ByVal workbookPath As String, _
    ByVal completionMessage As String, _
    ByVal escapeTargetPrefixesCsv As String, _
    ByVal onlyAValueRowFillTarget As String, _
    ByVal onlyAValueRowFillColorHex As String)

    Dim options As EscapePartsMarkingUiOptions
    options = EscapePartsMarking.CreateEscapePartsMarkingUiOptionsForForm()

    options.TargetWorkbookPath = Trim$(workbookPath)
    options.UseCompletionMessage = True
    options.completionMessage = Trim$(completionMessage)

    options.UseEscapeTargetPrefixesCsv = True
    options.escapeTargetPrefixesCsv = Trim$(escapeTargetPrefixesCsv)

    options.UseOnlyAValueRowFillTarget = True
    options.onlyAValueRowFillTarget = Trim$(onlyAValueRowFillTarget)

    options.UseOnlyAValueRowFillColorHex = True
    options.onlyAValueRowFillColorHex = Trim$(onlyAValueRowFillColorHex)

    If Len(options.TargetWorkbookPath) = 0 Then
        MsgBox "対象ブックパスは必須です。", vbExclamation
        Exit Sub
    End If

    EscapePartsMarking.RunMainWithUiOptions options
End Sub
