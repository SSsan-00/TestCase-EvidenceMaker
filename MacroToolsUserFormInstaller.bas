Attribute VB_Name = "MacroToolsUserFormInstaller"
Option Explicit

' ============================================================
' UserForm auto installer
' ------------------------------------------------------------
' This module creates:
' - frmMacroTools (UserForm)
' - modMacroToolsFormEntry (launcher module)
'
' Prerequisite:
' - Excel option "Trust access to the VBA project object model" must be ON.
' ============================================================

Private Const COMPONENT_TYPE_STD_MODULE As Long = 1
Private Const COMPONENT_TYPE_USERFORM As Long = 3

Public Sub InstallMacroToolsUserForm()
    On Error GoTo ErrorHandler

    Dim vbProj As Object
    Dim currentStep As String
    Dim formComponent As Object

    currentStep = "Precheck workbook writable"
    EnsureWorkbookWritable

    currentStep = "VBProject access"
    Set vbProj = ThisWorkbook.VBProject

    currentStep = "Precheck project unlocked"
    EnsureProjectUnlocked vbProj

    currentStep = "Precheck temp writable"
    EnsureTempWritable

    currentStep = "Remove existing launcher module"
    RemoveComponentIfExists vbProj, "modMacroToolsFormEntry"

    currentStep = "Find existing UserForm component"
    Set formComponent = GetComponentByName(vbProj, "frmMacroTools")
    If formComponent Is Nothing Then
        currentStep = "Create UserForm component"
        Set formComponent = TryAddUserFormComponent(vbProj)
        If formComponent Is Nothing Then
            Err.Raise vbObjectError + 3810, "InstallMacroToolsUserForm", _
                      "UserForm component could not be created." & vbCrLf & _
                      "Please manually insert one UserForm and rename it to frmMacroTools, then run this installer again."
        End If
        formComponent.Name = "frmMacroTools"
    End If

    currentStep = "Reset UserForm layout and code"
    PrepareFormComponent formComponent

    currentStep = "Build UserForm layout"
    BuildMacroToolsFormLayout formComponent.Designer

    currentStep = "Attach UserForm code"
    formComponent.CodeModule.AddFromString BuildMacroToolsFormCode()

    Dim launcherComponent As Object
    currentStep = "Create launcher module"
    Set launcherComponent = vbProj.VBComponents.Add(COMPONENT_TYPE_STD_MODULE)
    launcherComponent.Name = "modMacroToolsFormEntry"

    currentStep = "Attach launcher code"
    launcherComponent.CodeModule.AddFromString BuildMacroToolsLauncherCode()

    MsgBox "UserForm のインストールが完了しました。" & vbCrLf & _
           "実行: OpenMacroToolsForm", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "UserForm のインストールに失敗しました。" & vbCrLf & _
           "ステップ: " & currentStep & vbCrLf & _
           Err.Number & " : " & Err.Description & vbCrLf & vbCrLf & _
           "1004 エラーが出る場合は次を有効化してください:" & vbCrLf & _
           "Trust Center > Macro Settings > Trust access to the VBA project object model", _
           vbExclamation
End Sub

Private Function GetComponentByName(ByVal vbProj As Object, ByVal componentName As String) As Object
    Dim component As Object
    For Each component In vbProj.VBComponents
        If StrComp(component.Name, componentName, vbTextCompare) = 0 Then
            Set GetComponentByName = component
            Exit Function
        End If
    Next component
End Function

Private Sub PrepareFormComponent(ByVal formComponent As Object)
    ResetCodeModule formComponent.CodeModule
    ResetDesignerControls formComponent.Designer
End Sub

Private Sub ResetCodeModule(ByVal codeModule As Object)
    On Error Resume Next
    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If
    On Error GoTo 0
End Sub

Private Sub ResetDesignerControls(ByVal formDesigner As Object)
    On Error Resume Next

    Dim i As Long
    Dim controlName As String

    For i = CLng(formDesigner.Controls.Count) - 1 To 0 Step -1
        controlName = CStr(formDesigner.Controls.Item(i).Name)
        formDesigner.Controls.Remove controlName
    Next i

    On Error GoTo 0
End Sub

Private Sub EnsureWorkbookWritable()
    If ThisWorkbook.ReadOnly Then
        Err.Raise vbObjectError + 3811, "EnsureWorkbookWritable", _
                  "Workbook is read-only."
    End If
End Sub

Private Sub EnsureProjectUnlocked(ByVal vbProj As Object)
    On Error GoTo AccessError
    If CInt(vbProj.Protection) <> 0 Then
        Err.Raise vbObjectError + 3812, "EnsureProjectUnlocked", _
                  "VBA project is protected."
    End If
    Exit Sub

AccessError:
    Err.Raise vbObjectError + 3813, "EnsureProjectUnlocked", _
              "Unable to inspect VBA project protection: " & Err.Description
End Sub

Private Sub EnsureTempWritable()
    Dim tempDir As String
    Dim testPath As String
    Dim ff As Integer

    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then
        Err.Raise vbObjectError + 3814, "EnsureTempWritable", "TEMP environment variable is empty."
    End If

    If Len(Dir$(tempDir, vbDirectory)) = 0 Then
        Err.Raise vbObjectError + 3815, "EnsureTempWritable", "TEMP directory does not exist: " & tempDir
    End If

    testPath = tempDir & "\codex_vbe_write_test.tmp"
    ff = FreeFile
    Open testPath For Output As #ff
    Print #ff, "ok"
    Close #ff
    Kill testPath
End Sub

Private Function TryAddUserFormComponent(ByVal vbProj As Object) As Object
    Dim originalDir As String
    Dim candidateDirs As Variant
    Dim candidate As Variant
    Dim switched As Boolean
    Dim addedComponent As Object

    originalDir = CurDir$
    candidateDirs = Array( _
        Environ$("TEMP"), _
        ThisWorkbook.Path, _
        Environ$("SystemRoot"), _
        "C:\")

    For Each candidate In candidateDirs
        switched = False
        TrySwitchCurrentDirectory CStr(candidate), switched
        If switched Then
            If TryAddUserFormCore(vbProj, addedComponent) Then
                Set TryAddUserFormComponent = addedComponent
                Exit For
            End If
        End If
    Next candidate

    RestoreCurrentDirectory originalDir
End Function

Private Function TryAddUserFormCore(ByVal vbProj As Object, ByRef outComponent As Object) As Boolean
    On Error Resume Next
    Set outComponent = vbProj.VBComponents.Add(COMPONENT_TYPE_USERFORM)
    If outComponent Is Nothing Then
        Err.Clear
        Set outComponent = vbProj.VBComponents.Add(3)
    End If
    TryAddUserFormCore = Not (outComponent Is Nothing)
    On Error GoTo 0
End Function

Private Sub TrySwitchCurrentDirectory(ByVal targetDir As String, ByRef switched As Boolean)
    On Error Resume Next
    If Len(targetDir) = 0 Then Exit Sub
    If Len(Dir$(targetDir, vbDirectory)) = 0 Then Exit Sub

    ChDrive Left$(targetDir, 1)
    ChDir targetDir
    switched = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub RestoreCurrentDirectory(ByVal originalDir As String)
    On Error Resume Next
    If Len(originalDir) > 0 Then
        ChDrive Left$(originalDir, 1)
        ChDir originalDir
    End If
    On Error GoTo 0
End Sub

Private Sub RemoveComponentIfExists(ByVal vbProj As Object, ByVal componentName As String)
    Dim component As Object
    For Each component In vbProj.VBComponents
        If StrComp(component.Name, componentName, vbTextCompare) = 0 Then
            vbProj.VBComponents.Remove component
            Exit For
        End If
    Next component
End Sub

Private Sub BuildMacroToolsFormLayout(ByVal formDesigner As Object)
    On Error GoTo LayoutError

    Dim stageName As String
    stageName = "Form base properties"
    SafeSetProperty formDesigner, "Caption", "マクロツール"
    SafeSetProperty formDesigner, "Width", 700
    SafeSetProperty formDesigner, "Height", 700
    ' Some environments do not expose StartUpPosition on Designer.
    SafeSetProperty formDesigner, "StartUpPosition", 1
    SafeSetProperty formDesigner, "ControlBox", True
    ' レイアウトを縦方向に整理し、画面サイズ差分はスクロールで吸収する
    SafeSetProperty formDesigner, "ScrollBars", 2
    SafeSetProperty formDesigner, "ScrollHeight", 900
    SafeSetProperty formDesigner, "ScrollWidth", 670

    Dim frameEvidence As Object
    Dim frameTestCase As Object
    Dim frameConditional As Object
    Dim frameEscape As Object

    stageName = "Create frames"
    Set frameEvidence = AddFrame(formDesigner, "fraEvidence", "エビデンス生成", 12, 12, 650, 300)
    Set frameTestCase = AddFrame(formDesigner, "fraTestCase", "テストケース生成", 12, 320, 650, 90)
    Set frameConditional = AddFrame(formDesigner, "fraConditional", "条件分岐チェック", 12, 418, 650, 136)
    Set frameEscape = AddFrame(formDesigner, "fraEscape", "エスケープ箇所マーキング", 12, 562, 650, 240)

    stageName = "Build Evidence controls"
    BuildEvidenceControls frameEvidence
    stageName = "Build TestCase controls"
    BuildTestCaseControls frameTestCase
    stageName = "Build Conditional controls"
    BuildConditionalControls frameConditional
    stageName = "Build Escape controls"
    BuildEscapeControls frameEscape

    stageName = "Create close button"
    AddButton formDesigner, "btnCloseForm", "閉じる", 562, 810, 100, 24
    Exit Sub

LayoutError:
    Err.Raise vbObjectError + 3802, "BuildMacroToolsFormLayout", _
              "Stage: " & stageName & " / " & Err.Description
End Sub

Private Sub BuildEvidenceControls(ByVal parent As Object)
    AddLabel parent, "lblSourcePath", "参照元ブックパス", 12, 18, 96, 16
    AddTextBox parent, "txtSourcePath", "", 112, 16, 460, 18
    AddButton parent, "btnBrowseEvidenceSource", "...", 580, 15, 24, 20

    AddLabel parent, "lblInputFileName", "入力ファイル名", 12, 48, 96, 16
    AddTextBox parent, "txtInputFileName", "", 112, 46, 360, 18

    AddLabel parent, "lblSlotHeight", "行オフセット", 12, 78, 96, 16
    AddTextBox parent, "txtSlotHeight", "50", 112, 76, 56, 18

    AddLabel parent, "lblOutputFilter", "出力シート絞り込み", 206, 78, 96, 16
    AddTextBox parent, "txtOutputFilter", "", 306, 76, 246, 18

    AddCheckBox parent, "chkTopBorder", "上罫線を有効化", True, 12, 108, 150, 16
    AddCheckBox parent, "chkExcludePattern", "除外パターンを有効化", True, 182, 108, 160, 16

    AddLabel parent, "lblExcludePatterns", "除外パターン", 12, 136, 96, 16
    AddTextBox parent, "txtExcludePatterns", "A4,A5,A1-1,A2-3-1", 112, 134, 460, 18

    AddCheckBox parent, "chkSkipGray", "灰色塗りつぶしセルを読み飛ばす", True, 12, 164, 220, 16
    AddLabel parent, "lblSkipColors", "読み飛ばし色 (#RRGGBB)", 12, 190, 130, 16
    AddTextBox parent, "txtSkipColors", "#f2f2f2,#d9d9d9,#bfbfbf,#a6a6a6,#808080", 146, 188, 430, 18

    AddCheckBox parent, "chkRightBorder", "右罫線を有効化", True, 12, 218, 140, 16
    AddCheckBox parent, "chkUseRightBorderCol", "右罫線列名を指定", True, 170, 218, 140, 16
    AddTextBox parent, "txtRightBorderCol", "Q", 320, 216, 40, 18

    AddButton parent, "btnRunEvidence", "実行（エビデンス生成）", 470, 252, 162, 24
End Sub

Private Sub BuildTestCaseControls(ByVal parent As Object)
    AddLabel parent, "lblFeatureId", "機能ID", 12, 28, 80, 16
    AddTextBox parent, "txtFeatureId", "", 96, 26, 360, 18

    AddButton parent, "btnRunTestCase", "実行（テストケース生成）", 470, 24, 162, 24
End Sub

Private Sub BuildConditionalControls(ByVal parent As Object)
    AddLabel parent, "lblCondFeatureName", "機能名", 12, 24, 80, 16
    AddTextBox parent, "txtCondFeatureName", "", 96, 22, 440, 18

    AddLabel parent, "lblCondWorkbookPath", "対象ブックパス", 12, 52, 80, 16
    AddTextBox parent, "txtCondWorkbookPath", "", 96, 50, 440, 18
    AddButton parent, "btnBrowseConditionalWorkbook", "...", 540, 49, 24, 20

    AddCheckBox parent, "chkLeadingFunctionB1", "先頭FunctionをB1開始にする", True, 12, 80, 210, 16
    AddButton parent, "btnRunConditional", "実行（条件分岐チェック）", 470, 104, 162, 24
End Sub

Private Sub BuildEscapeControls(ByVal parent As Object)
    AddLabel parent, "lblEscapeWorkbookPath", "対象ブックパス", 12, 24, 96, 16
    AddTextBox parent, "txtEscapeWorkbookPath", "", 112, 22, 440, 18
    AddButton parent, "btnBrowseEscapeWorkbook", "...", 556, 21, 24, 20

    AddLabel parent, "lblCompletionMessage", "完了メッセージ", 12, 52, 96, 16
    AddTextBox parent, "txtCompletionMessage", "SQLインジェクション対策済み", 112, 50, 320, 18

    AddLabel parent, "lblPrefixes", "エスケープ関数一覧", 12, 80, 96, 16
    AddTextBox parent, "txtPrefixes", "sqlS,sqlN", 112, 78, 320, 18

    AddLabel parent, "lblFillTarget", "塗りつぶし対象", 12, 112, 120, 16
    AddOptionButton parent, "optFillNone", "塗りつぶしなし", False, 136, 110, 90, 16
    AddOptionButton parent, "optFillLeft", "A列のみ", False, 230, 110, 70, 16
    AddOptionButton parent, "optFillRight", "B列のみ", False, 304, 110, 70, 16
    AddOptionButton parent, "optFillBoth", "A,B列", True, 378, 110, 70, 16

    AddLabel parent, "lblFillColor", "塗りつぶし色", 12, 144, 120, 16
    AddTextBox parent, "txtFillColor", "#a6a6a6", 136, 142, 120, 18

    AddButton parent, "btnRunEscape", "実行（エスケープ箇所マーキング）", 422, 176, 210, 24
End Sub

Private Function AddFrame( _
    ByVal parent As Object, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthValue As Single, _
    ByVal heightValue As Single) As Object

    Dim ctl As Object
    Set ctl = AddControlSafe(parent, "Forms.Frame.1", controlName)
    ctl.Caption = caption
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthValue
    ctl.Height = heightValue
    Set AddFrame = ctl
End Function

Private Sub AddLabel( _
    ByVal parent As Object, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthValue As Single, _
    ByVal heightValue As Single)

    Dim ctl As Object
    Set ctl = AddControlSafe(parent, "Forms.Label.1", controlName)
    ctl.Caption = caption
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthValue
    ctl.Height = heightValue
End Sub

Private Sub AddTextBox( _
    ByVal parent As Object, _
    ByVal controlName As String, _
    ByVal initialText As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthValue As Single, _
    ByVal heightValue As Single)

    Dim ctl As Object
    Set ctl = AddControlSafe(parent, "Forms.TextBox.1", controlName)
    SafeSetProperty ctl, "Value", initialText
    SafeSetProperty ctl, "Text", initialText
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthValue
    ctl.Height = heightValue
End Sub

Private Sub AddCheckBox( _
    ByVal parent As Object, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal initialValue As Boolean, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthValue As Single, _
    ByVal heightValue As Single)

    Dim ctl As Object
    Set ctl = AddControlSafe(parent, "Forms.CheckBox.1", controlName)
    ctl.Caption = caption
    ctl.Value = initialValue
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthValue
    ctl.Height = heightValue
End Sub

Private Sub AddOptionButton( _
    ByVal parent As Object, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal initialValue As Boolean, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthValue As Single, _
    ByVal heightValue As Single)

    Dim ctl As Object
    Set ctl = AddControlSafe(parent, "Forms.OptionButton.1", controlName)
    ctl.Caption = caption
    ctl.Value = initialValue
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthValue
    ctl.Height = heightValue
End Sub

Private Sub AddButton( _
    ByVal parent As Object, _
    ByVal controlName As String, _
    ByVal caption As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthValue As Single, _
    ByVal heightValue As Single)

    Dim ctl As Object
    Set ctl = AddControlSafe(parent, "Forms.CommandButton.1", controlName)
    ctl.Caption = caption
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthValue
    ctl.Height = heightValue
End Sub

Private Sub AddComboBox( _
    ByVal parent As Object, _
    ByVal controlName As String, _
    ByVal leftPos As Single, _
    ByVal topPos As Single, _
    ByVal widthValue As Single, _
    ByVal heightValue As Single)

    Dim ctl As Object
    Set ctl = AddControlSafe(parent, "Forms.ComboBox.1", controlName)
    ctl.Left = leftPos
    ctl.Top = topPos
    ctl.Width = widthValue
    ctl.Height = heightValue
End Sub

Private Function BuildMacroToolsFormCode() As String
    Dim codeText As String

    AppendLine codeText, "Option Explicit"
    AppendLine codeText, ""
    AppendLine codeText, "Private Const DEFAULT_SLOT_HEIGHT As Long = 50"
    AppendLine codeText, "Private Const DEFAULT_RIGHT_BORDER_COL As Long = 17"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub UserForm_Initialize()"
    AppendLine codeText, "    chkUseRightBorderCol_Click"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)"
    AppendLine codeText, "    Cancel = False"
    AppendLine codeText, "    If CloseMode = 0 Then Unload Me"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub btnCloseForm_Click()"
    AppendLine codeText, "    Unload Me"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub chkUseRightBorderCol_Click()"
    AppendLine codeText, "    Me.txtRightBorderCol.Enabled = CBool(Me.chkUseRightBorderCol.Value)"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub btnBrowseEvidenceSource_Click()"
    AppendLine codeText, "    Me.txtSourcePath.Value = PickOpenWorkbookPath(CStr(Me.txtSourcePath.Value))"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub btnBrowseConditionalWorkbook_Click()"
    AppendLine codeText, "    Me.txtCondWorkbookPath.Value = PickOpenWorkbookPath(CStr(Me.txtCondWorkbookPath.Value))"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub btnBrowseEscapeWorkbook_Click()"
    AppendLine codeText, "    Me.txtEscapeWorkbookPath.Value = PickOpenWorkbookPath(CStr(Me.txtEscapeWorkbookPath.Value))"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub btnRunEvidence_Click()"
    AppendLine codeText, "    MacroUserFormBridge.RunBetaEvidenceFromForm _"
    AppendLine codeText, "        sourceWorkbookPath:=Trim$(CStr(Me.txtSourcePath.Value)), _"
    AppendLine codeText, "        inputFileName:=Trim$(CStr(Me.txtInputFileName.Value)), _"
    AppendLine codeText, "        useSlotHeight:=True, _"
    AppendLine codeText, "        slotHeight:=ReadLongOrDefault(CStr(Me.txtSlotHeight.Value), DEFAULT_SLOT_HEIGHT), _"
    AppendLine codeText, "        useOutputSheetFilter:=True, _"
    AppendLine codeText, "        outputSheetFilterText:=Trim$(CStr(Me.txtOutputFilter.Value)), _"
    AppendLine codeText, "        topBorderEnabled:=CBool(Me.chkTopBorder.Value), _"
    AppendLine codeText, "        excludeOutputSheetByPatternEnabled:=CBool(Me.chkExcludePattern.Value), _"
    AppendLine codeText, "        excludedOutputSheetNamePatterns:=Trim$(CStr(Me.txtExcludePatterns.Value)), _"
    AppendLine codeText, "        skipGrayFilledSourceCellEnabled:=CBool(Me.chkSkipGray.Value), _"
    AppendLine codeText, "        sourceSkipFillColorHexCodes:=Trim$(CStr(Me.txtSkipColors.Value)), _"
    AppendLine codeText, "        rightBorderEnabled:=CBool(Me.chkRightBorder.Value), _"
    AppendLine codeText, "        useRightBorderTargetCol:=CBool(Me.chkUseRightBorderCol.Value), _"
    AppendLine codeText, "        rightBorderTargetCol:=ReadColumnIndexOrDefault(CStr(Me.txtRightBorderCol.Value), DEFAULT_RIGHT_BORDER_COL)"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub btnRunTestCase_Click()"
    AppendLine codeText, "    MacroUserFormBridge.RunBetaTestCaseFromForm _"
    AppendLine codeText, "        featureId:=Trim$(CStr(Me.txtFeatureId.Value)), _"
    AppendLine codeText, "        useOutputPath:=False, _"
    AppendLine codeText, "        outputPath:=vbNullString"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub btnRunConditional_Click()"
    AppendLine codeText, "    MacroUserFormBridge.RunConditionalBranchCheckerFromForm _"
    AppendLine codeText, "        featureName:=Trim$(CStr(Me.txtCondFeatureName.Value)), _"
    AppendLine codeText, "        workbookPath:=Trim$(CStr(Me.txtCondWorkbookPath.Value)), _"
    AppendLine codeText, "        leadingFunctionStartsFromB1:=CBool(Me.chkLeadingFunctionB1.Value)"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Sub btnRunEscape_Click()"
    AppendLine codeText, "    MacroUserFormBridge.RunEscapePartsMarkingFromForm _"
    AppendLine codeText, "        workbookPath:=Trim$(CStr(Me.txtEscapeWorkbookPath.Value)), _"
    AppendLine codeText, "        completionMessage:=Trim$(CStr(Me.txtCompletionMessage.Value)), _"
    AppendLine codeText, "        escapeTargetPrefixesCsv:=Trim$(CStr(Me.txtPrefixes.Value)), _"
    AppendLine codeText, "        onlyAValueRowFillTarget:=ResolveOnlyAValueRowFillTarget(), _"
    AppendLine codeText, "        onlyAValueRowFillColorHex:=Trim$(CStr(Me.txtFillColor.Value))"
    AppendLine codeText, "End Sub"
    AppendLine codeText, ""
    AppendLine codeText, "Private Function PickOpenWorkbookPath(ByVal initialPath As String) As String"
    AppendLine codeText, "    On Error GoTo Failed"
    AppendLine codeText, "    Dim pickedPath As Variant"
    AppendLine codeText, "    pickedPath = Application.GetOpenFilename(""Excel Files (*.xlsx;*.xlsm;*.xls;*.xlsb),*.xlsx;*.xlsm;*.xls;*.xlsb"", 1, ""ブックを選択"")"
    AppendLine codeText, "    If VarType(pickedPath) = vbBoolean Then"
    AppendLine codeText, "        PickOpenWorkbookPath = initialPath"
    AppendLine codeText, "    Else"
    AppendLine codeText, "        PickOpenWorkbookPath = CStr(pickedPath)"
    AppendLine codeText, "    End If"
    AppendLine codeText, "    Exit Function"
    AppendLine codeText, ""
    AppendLine codeText, "Failed:"
    AppendLine codeText, "    MsgBox ""ファイル選択ダイアログを開けませんでした。"" & vbCrLf & _"
    AppendLine codeText, "           CStr(Err.Number) & "" : "" & Err.Description, vbExclamation"
    AppendLine codeText, "    PickOpenWorkbookPath = initialPath"
    AppendLine codeText, "End Function"
    AppendLine codeText, ""
    AppendLine codeText, "Private Function ResolveOnlyAValueRowFillTarget() As String"
    AppendLine codeText, "    If CBool(Me.optFillNone.Value) Then"
    AppendLine codeText, "        ResolveOnlyAValueRowFillTarget = ""None"""
    AppendLine codeText, "    ElseIf CBool(Me.optFillLeft.Value) Then"
    AppendLine codeText, "        ResolveOnlyAValueRowFillTarget = ""Left"""
    AppendLine codeText, "    ElseIf CBool(Me.optFillRight.Value) Then"
    AppendLine codeText, "        ResolveOnlyAValueRowFillTarget = ""Right"""
    AppendLine codeText, "    Else"
    AppendLine codeText, "        ResolveOnlyAValueRowFillTarget = ""Both"""
    AppendLine codeText, "    End If"
    AppendLine codeText, "End Function"
    AppendLine codeText, ""
    AppendLine codeText, ""
    AppendLine codeText, "Private Function ReadColumnIndexOrDefault(ByVal textValue As String, ByVal defaultValue As Long) As Long"
    AppendLine codeText, "    Dim s As String"
    AppendLine codeText, "    Dim i As Long"
    AppendLine codeText, "    Dim ch As Integer"
    AppendLine codeText, "    Dim v As Long"
    AppendLine codeText, ""
    AppendLine codeText, "    s = UCase$(Trim$(textValue))"
    AppendLine codeText, "    If Len(s) = 0 Then"
    AppendLine codeText, "        ReadColumnIndexOrDefault = defaultValue"
    AppendLine codeText, "        Exit Function"
    AppendLine codeText, "    End If"
    AppendLine codeText, ""
    AppendLine codeText, "    If Left$(s, 1) = ""$"" Then s = Mid$(s, 2)"
    AppendLine codeText, "    If Len(s) = 0 Then"
    AppendLine codeText, "        ReadColumnIndexOrDefault = defaultValue"
    AppendLine codeText, "        Exit Function"
    AppendLine codeText, "    End If"
    AppendLine codeText, ""
    AppendLine codeText, "    ' 互換のため数値入力も受け付ける"
    AppendLine codeText, "    If IsNumeric(s) Then"
    AppendLine codeText, "        v = CLng(Val(s))"
    AppendLine codeText, "        If v >= 1 And v <= 16384 Then"
    AppendLine codeText, "            ReadColumnIndexOrDefault = v"
    AppendLine codeText, "        Else"
    AppendLine codeText, "            ReadColumnIndexOrDefault = defaultValue"
    AppendLine codeText, "        End If"
    AppendLine codeText, "        Exit Function"
    AppendLine codeText, "    End If"
    AppendLine codeText, ""
    AppendLine codeText, "    v = 0"
    AppendLine codeText, "    For i = 1 To Len(s)"
    AppendLine codeText, "        ch = Asc(Mid$(s, i, 1))"
    AppendLine codeText, "        If ch < 65 Or ch > 90 Then"
    AppendLine codeText, "            ReadColumnIndexOrDefault = defaultValue"
    AppendLine codeText, "            Exit Function"
    AppendLine codeText, "        End If"
    AppendLine codeText, "        v = v * 26 + (ch - 64)"
    AppendLine codeText, "        If v > 16384 Then"
    AppendLine codeText, "            ReadColumnIndexOrDefault = defaultValue"
    AppendLine codeText, "            Exit Function"
    AppendLine codeText, "        End If"
    AppendLine codeText, "    Next i"
    AppendLine codeText, ""
    AppendLine codeText, "    If v <= 0 Then"
    AppendLine codeText, "        ReadColumnIndexOrDefault = defaultValue"
    AppendLine codeText, "    Else"
    AppendLine codeText, "        ReadColumnIndexOrDefault = v"
    AppendLine codeText, "    End If"
    AppendLine codeText, "End Function"
    AppendLine codeText, ""
    AppendLine codeText, "Private Function ReadLongOrDefault(ByVal textValue As String, ByVal defaultValue As Long) As Long"
    AppendLine codeText, "    Dim n As Double"
    AppendLine codeText, "    textValue = Trim$(textValue)"
    AppendLine codeText, "    If Len(textValue) = 0 Then"
    AppendLine codeText, "        ReadLongOrDefault = defaultValue"
    AppendLine codeText, "        Exit Function"
    AppendLine codeText, "    End If"
    AppendLine codeText, "    If Not IsNumeric(textValue) Then"
    AppendLine codeText, "        ReadLongOrDefault = defaultValue"
    AppendLine codeText, "        Exit Function"
    AppendLine codeText, "    End If"
    AppendLine codeText, "    n = CDbl(textValue)"
    AppendLine codeText, "    If n <= 0 Or n <> Fix(n) Then"
    AppendLine codeText, "        ReadLongOrDefault = defaultValue"
    AppendLine codeText, "    Else"
    AppendLine codeText, "        ReadLongOrDefault = CLng(n)"
    AppendLine codeText, "    End If"
    AppendLine codeText, "End Function"

    BuildMacroToolsFormCode = codeText
End Function

Private Function AddControlSafe( _
    ByVal parent As Object, _
    ByVal progId As String, _
    ByVal controlName As String) As Object

    On Error GoTo TryTwoArgs
    Set AddControlSafe = parent.Controls.Add(progId, controlName, True)
    Exit Function

TryTwoArgs:
    Err.Clear
    On Error GoTo TryOneArg
    Set AddControlSafe = parent.Controls.Add(progId, controlName)
    Exit Function

TryOneArg:
    Err.Clear
    On Error GoTo CreateFail
    Set AddControlSafe = parent.Controls.Add(progId)
    On Error Resume Next
    AddControlSafe.Name = controlName
    On Error GoTo 0
    Exit Function

CreateFail:
    Err.Raise vbObjectError + 3801, "AddControlSafe", _
              "Control create failed (" & progId & ", " & controlName & "): " & Err.Description
End Function

Private Sub SafeSetProperty(ByVal target As Object, ByVal propertyName As String, ByVal value As Variant)
    On Error Resume Next
    CallByName target, propertyName, VbLet, value
    On Error GoTo 0
End Sub

Private Function BuildMacroToolsLauncherCode() As String
    Dim codeText As String
    AppendLine codeText, "Option Explicit"
    AppendLine codeText, ""
    AppendLine codeText, "Public Sub OpenMacroToolsForm()"
    AppendLine codeText, "    frmMacroTools.Show"
    AppendLine codeText, "End Sub"
    BuildMacroToolsLauncherCode = codeText
End Function

Private Sub AppendLine(ByRef buffer As String, ByVal lineText As String)
    If Len(buffer) = 0 Then
        buffer = lineText
    Else
        buffer = buffer & vbCrLf & lineText
    End If
End Sub
