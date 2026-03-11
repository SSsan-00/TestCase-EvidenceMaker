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
    ByVal useRightBorderTargetCol As Boolean, _
    ByVal rightBorderTargetCol As Long)

    Dim options As BetaEvidenceUiOptions
    BetaEvidenceGenerator.InitializeBetaEvidenceUiOptionsForForm options

    options.SourceWorkbookPath = Trim$(sourceWorkbookPath)
    options.InputFileName = Trim$(inputFileName)

    options.UseSlotHeight = useSlotHeight
    If useSlotHeight And slotHeight > 0 Then
        options.SlotHeight = slotHeight
    End If

    options.UseOutputSheetFilter = useOutputSheetFilter
    options.OutputSheetFilterText = Trim$(outputSheetFilterText)

    options.OverrideTopBorderEnabled = True
    options.TopBorderEnabled = topBorderEnabled

    options.OverrideExcludeOutputSheetByPatternEnabled = True
    options.ExcludeOutputSheetByPatternEnabled = excludeOutputSheetByPatternEnabled
    options.UseExcludedOutputSheetNamePatterns = True
    options.ExcludedOutputSheetNamePatterns = Trim$(excludedOutputSheetNamePatterns)

    options.OverrideSkipGrayFilledSourceCellEnabled = True
    options.SkipGrayFilledSourceCellEnabled = skipGrayFilledSourceCellEnabled
    options.UseSourceSkipFillColorHexCodes = True
    options.SourceSkipFillColorHexCodes = Trim$(sourceSkipFillColorHexCodes)

    options.OverrideRightBorderEnabled = True
    options.RightBorderEnabled = rightBorderEnabled

    options.UseRightBorderTargetCol = useRightBorderTargetCol
    If useRightBorderTargetCol And rightBorderTargetCol >= 1 And rightBorderTargetCol <= 16384 Then
        options.RightBorderTargetCol = rightBorderTargetCol
    End If

    If Len(options.SourceWorkbookPath) = 0 Or Len(options.InputFileName) = 0 Then
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

    options.FeatureId = Trim$(featureId)
    options.UseOutputPath = useOutputPath
    options.OutputPath = Trim$(outputPath)

    If Len(options.FeatureId) = 0 Then
        MsgBox "機能連番は必須です。", vbExclamation
        Exit Sub
    End If

    BetaTestCaseGenerator.RunMainWithUiOptions options
End Sub

Public Sub RunConditionalBranchCheckerFromForm( _
    ByVal featureName As String, _
    ByVal workbookPath As String, _
    ByVal leadingFunctionStartsFromB1 As Boolean)

    Dim options As ConditionalBranchCheckerUiOptions
    options = ConditionalBranchChecker.CreateConditionalBranchCheckerUiOptionsForForm()

    options.FeatureName = Trim$(featureName)
    options.WorkbookPath = Trim$(workbookPath)
    options.OverrideLeadingFunctionStartsFromB1 = True
    options.LeadingFunctionStartsFromB1 = leadingFunctionStartsFromB1

    If Len(options.FeatureName) = 0 Or Len(options.WorkbookPath) = 0 Then
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
    options.CompletionMessage = Trim$(completionMessage)

    options.UseEscapeTargetPrefixesCsv = True
    options.EscapeTargetPrefixesCsv = Trim$(escapeTargetPrefixesCsv)

    options.UseOnlyAValueRowFillTarget = True
    options.OnlyAValueRowFillTarget = Trim$(onlyAValueRowFillTarget)

    options.UseOnlyAValueRowFillColorHex = True
    options.OnlyAValueRowFillColorHex = Trim$(onlyAValueRowFillColorHex)

    If Len(options.TargetWorkbookPath) = 0 Then
        MsgBox "対象ブックパスは必須です。", vbExclamation
        Exit Sub
    End If

    EscapePartsMarking.RunMainWithUiOptions options
End Sub