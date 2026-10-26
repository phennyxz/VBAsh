Attribute VB_Name = "modVBAshInstaller"
Option Explicit

Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_CLASSMODULE As Long = 2
Private Const VBEXT_CT_MSFORM As Long = 3
Private Const MSO_FILE_DIALOG_FOLDER_PICKER As Long = 4
Private Const MSO_PROPERTY_TYPE_STRING As Long = 4
Private Const INSTALL_MARKER_NAME As String = "VBAshInstallerVersion"
Private Const INSTALL_MARKER_VALUE As String = "1"
Private Const INSTALL_MODULE_COUNT_NAME As String = "VBAshInstalledModuleCount"
Private Const INSTALL_MODULE_PROPERTY_PREFIX As String = "VBAshInstalledModule"

Public Sub VBAsh_Install()
    Dim folderDialog As Object
    Dim project As Object
    Dim component As Object
    Dim addedComponents As Collection
    Dim sourceCode As Object
    Dim sourceFiles As Variant
    Dim fileName As Variant
    Dim formComponent As Object
    Dim moduleName As String
    Dim sourcePath As String
    Dim sourceFolder As String
    Dim failure As String
    Dim answer As VbMsgBoxResult

    Set folderDialog = Application.FileDialog(MSO_FILE_DIALOG_FOLDER_PICKER)
    folderDialog.Title = "Selecione a pasta raiz do VBAsh"
    folderDialog.AllowMultiSelect = False
    If folderDialog.Show <> -1 Then Exit Sub
    sourceFolder = CStr(folderDialog.SelectedItems(1))

    Set sourceCode = CreateObject("Scripting.Dictionary")
    sourceCode.CompareMode = vbTextCompare
    sourceFiles = VBAshSourceFiles(sourceFolder)
    If IsEmpty(sourceFiles) Then
        MsgBox "Nenhum arquivo .bas ou .cls foi encontrado na pasta selecionada.", _
               vbExclamation, "VBAsh"
        Exit Sub
    End If

    For Each fileName In sourceFiles
        sourcePath = VBAshBuildPath(sourceFolder, CStr(fileName))
        If Len(Dir$(sourcePath)) = 0 Then
            MsgBox "Arquivo necessário não encontrado:" & vbCrLf & sourcePath, vbExclamation, "VBAsh"
            Exit Sub
        End If

        moduleName = VBAshModuleName(CStr(fileName))
        If sourceCode.Exists(moduleName) Then
            MsgBox "Há nomes de módulos duplicados na lista do instalador.", vbCritical, "VBAsh"
            Exit Sub
        End If
        sourceCode.Add moduleName, VBAshReadUtf8File(sourcePath)
    Next fileName

    On Error GoTo ProjectAccessError
    Set project = ThisWorkbook.VBProject
    If VBAshHasInstallMarker() Then
        MsgBox "O instalador já registrou uma instalação neste arquivo. Execute VBAsh_Uninstall antes de reinstalar.", _
               vbExclamation, "VBAsh"
        Exit Sub
    End If

    For Each fileName In sourceFiles
        moduleName = VBAshModuleName(CStr(fileName))
        If VBAshComponentExists(project, moduleName) Then
            MsgBox "O módulo '" & moduleName & "' já existe. Use VBAsh_Uninstall antes de reinstalar.", _
                   vbExclamation, "VBAsh"
            Exit Sub
        End If
    Next fileName
    If VBAshComponentExists(project, "frmVBAsh") Then
        MsgBox "O formulário frmVBAsh já existe. Use VBAsh_Uninstall antes de reinstalar.", _
               vbExclamation, "VBAsh"
        Exit Sub
    End If

    answer = MsgBox( _
        "O instalador vai adicionar " & CStr(sourceCode.Count) & _
        " módulos e o formulário terminal ao projeto deste arquivo:" & vbCrLf & ThisWorkbook.Name & vbCrLf & vbCrLf & _
        "Somente continue se você confia nos arquivos da pasta selecionada:" & vbCrLf & _
        sourceFolder & vbCrLf & vbCrLf & _
        "Deseja continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Instalar VBAsh")
    If answer <> vbYes Then Exit Sub

    Set addedComponents = New Collection
    On Error GoTo ImportError

    For Each fileName In sourceFiles
        moduleName = VBAshModuleName(CStr(fileName))
        If LCase$(Right$(CStr(fileName), 4)) = ".bas" Then
            Set component = project.VBComponents.Add(VBEXT_CT_STDMODULE)
        Else
            Set component = project.VBComponents.Add(VBEXT_CT_CLASSMODULE)
        End If

        addedComponents.Add component
        component.Name = moduleName
        component.CodeModule.AddFromString CStr(sourceCode(moduleName))
    Next fileName

    Set formComponent = project.VBComponents.Add(VBEXT_CT_MSFORM)
    addedComponents.Add formComponent
    formComponent.Name = "frmVBAsh"
    formComponent.CodeModule.AddFromString VBAshTerminalFormCode()

    VBAshSaveInstalledModules sourceFiles
    ThisWorkbook.CustomDocumentProperties.Add _
        INSTALL_MARKER_NAME, False, MSO_PROPERTY_TYPE_STRING, INSTALL_MARKER_VALUE

    On Error GoTo 0
    MsgBox "Módulos instalados. O instalador permanece no projeto; para removê-los depois, execute VBAsh_Uninstall.", _
           vbInformation, "VBAsh"

    If MsgBox("Deseja iniciar o shell agora?", vbQuestion + vbYesNo + vbDefaultButton2, "VBAsh") = vbYes Then
        On Error GoTo StartError
        Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!VBAsh_Start"
        On Error GoTo 0
    End If
    Exit Sub

ProjectAccessError:
    MsgBox "Não foi possível acessar o projeto VBA. No Excel, habilite " & _
           "'Confiar no acesso ao modelo de objeto do projeto VBA' nas configurações da Central de Confiabilidade." & _
           vbCrLf & vbCrLf & Err.Description, vbCritical, "VBAsh"
    Exit Sub

ImportError:
    failure = Err.Description
    On Error Resume Next
    For Each component In addedComponents
        project.VBComponents.Remove component
    Next component
    VBAshDeleteInstalledModules sourceFiles
    If VBAshHasInstallMarker() Then ThisWorkbook.CustomDocumentProperties(INSTALL_MARKER_NAME).Delete
    On Error GoTo 0
    MsgBox "A importação falhou e os módulos adicionados nesta tentativa foram removidos." & _
           vbCrLf & failure, vbCritical, "VBAsh"
    Exit Sub

StartError:
    MsgBox "Os módulos foram instalados, mas não foi possível iniciar VBAsh_Start:" & _
           vbCrLf & Err.Description, vbExclamation, "VBAsh"
End Sub

Public Sub VBAsh_Uninstall()
    Dim project As Object
    Dim component As Object
    Dim moduleName As Variant
    Dim answer As VbMsgBoxResult
    Dim removedCount As Long

    If Not VBAshHasInstallMarker() Then
        MsgBox "Este instalador não encontrou uma instalação registrada neste arquivo; nada foi removido.", _
               vbInformation, "VBAsh"
        Exit Sub
    End If

    answer = MsgBox( _
        "Isso removerá do projeto deste arquivo os módulos registrados na instalação e o formulário frmVBAsh." & _
        vbCrLf & "O módulo modVBAshInstaller não será removido." & vbCrLf & vbCrLf & _
        "Deseja continuar?", vbExclamation + vbYesNo + vbDefaultButton2, "Remover VBAsh")
    If answer <> vbYes Then Exit Sub

    On Error GoTo ProjectAccessError
    Set project = ThisWorkbook.VBProject

    For Each moduleName In VBAshInstalledModuleNames()
        If VBAshComponentExists(project, CStr(moduleName)) Then
            Set component = project.VBComponents(CStr(moduleName))
            project.VBComponents.Remove component
            removedCount = removedCount + 1
        End If
    Next moduleName
    If VBAshComponentExists(project, "frmVBAsh") Then
        Set component = project.VBComponents("frmVBAsh")
        project.VBComponents.Remove component
        removedCount = removedCount + 1
    End If
    VBAshDeleteInstalledModuleProperties
    ThisWorkbook.CustomDocumentProperties(INSTALL_MARKER_NAME).Delete

    On Error GoTo 0
    MsgBox CStr(removedCount) & " módulo(s) VBAsh removido(s). O instalador foi mantido.", _
           vbInformation, "VBAsh"
    Exit Sub

ProjectAccessError:
    MsgBox "Não foi possível acessar o projeto VBA. Verifique a configuração " & _
           "'Confiar no acesso ao modelo de objeto do projeto VBA'." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "VBAsh"
End Sub

Private Function VBAshSourceFiles(ByVal sourceFolder As String) As Variant
    Dim fileSystem As Object
    Dim folder As Object
    Dim sourceFile As Object
    Dim names As Collection
    Dim files() As String
    Dim extension As String
    Dim i As Long

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(sourceFolder)
    Set names = New Collection
    For Each sourceFile In folder.Files
        extension = LCase$(fileSystem.GetExtensionName(sourceFile.Name))
        If (extension = "bas" Or extension = "cls") And _
           StrComp(sourceFile.Name, "VBAshInstaller.bas", vbTextCompare) <> 0 Then
            names.Add CStr(sourceFile.Name)
        End If
    Next sourceFile

    If names.Count = 0 Then Exit Function
    ReDim files(0 To names.Count - 1)
    For i = 1 To names.Count
        files(i - 1) = CStr(names(i))
    Next i
    VBAshSourceFiles = files
End Function

Private Sub VBAshSaveInstalledModules(ByVal sourceFiles As Variant)
    Dim i As Long

    For i = LBound(sourceFiles) To UBound(sourceFiles)
        ThisWorkbook.CustomDocumentProperties.Add _
            VBAshInstalledModulePropertyName(i - LBound(sourceFiles) + 1), _
            False, MSO_PROPERTY_TYPE_STRING, VBAshModuleName(CStr(sourceFiles(i)))
    Next i
    ThisWorkbook.CustomDocumentProperties.Add _
        INSTALL_MODULE_COUNT_NAME, False, MSO_PROPERTY_TYPE_STRING, _
        CStr(UBound(sourceFiles) - LBound(sourceFiles) + 1)
End Sub

Private Function VBAshInstalledModuleNames() As Variant
    Dim moduleCount As Long
    Dim names() As String
    Dim i As Long

    If Not VBAshHasDocumentProperty(INSTALL_MODULE_COUNT_NAME) Then
        VBAshInstalledModuleNames = VBAshLegacyModuleNames()
        Exit Function
    End If

    moduleCount = CLng(ThisWorkbook.CustomDocumentProperties(INSTALL_MODULE_COUNT_NAME).Value)
    If moduleCount < 1 Then
        Err.Raise vbObjectError + 1060, "VBAsh_Uninstall", "A lista de módulos instalados está inválida."
    End If
    ReDim names(0 To moduleCount - 1)
    For i = 1 To moduleCount
        names(i - 1) = CStr(ThisWorkbook.CustomDocumentProperties( _
            VBAshInstalledModulePropertyName(i)).Value)
    Next i
    VBAshInstalledModuleNames = names
End Function

Private Function VBAshLegacyModuleNames() As Variant
    VBAshLegacyModuleNames = Array( _
        "VBAsh", "VBAshTests", "IBuiltin", "clsBuiltinCd", "clsBuiltinEcho", _
        "clsBuiltinExit", "clsBuiltinExport", "clsBuiltinHelp", "clsBuiltinPwd", _
        "clsBuiltinRegistry", "clsBuiltinSetEnv", "clsBuiltinTest", "clsBuiltinUnset", _
        "clsCommand", "clsCommandResolver", "clsConsole", "clsFileDescriptorTable", _
        "clsFileStream", "clsParser", "clsRedirection", "clsSession", "clsStream", _
        "clsToken", "clsTokenizer", "clsBuiltinLs", "clsBuiltinMkdir")
End Function

Private Sub VBAshDeleteInstalledModules(ByVal sourceFiles As Variant)
    Dim i As Long

    On Error Resume Next
    For i = LBound(sourceFiles) To UBound(sourceFiles)
        ThisWorkbook.CustomDocumentProperties( _
            VBAshInstalledModulePropertyName(i - LBound(sourceFiles) + 1)).Delete
    Next i
    ThisWorkbook.CustomDocumentProperties(INSTALL_MODULE_COUNT_NAME).Delete
    On Error GoTo 0
End Sub

Private Sub VBAshDeleteInstalledModuleProperties()
    Dim moduleCount As Long
    Dim i As Long

    If VBAshHasDocumentProperty(INSTALL_MODULE_COUNT_NAME) Then
        moduleCount = CLng(ThisWorkbook.CustomDocumentProperties(INSTALL_MODULE_COUNT_NAME).Value)
        For i = 1 To moduleCount
            ThisWorkbook.CustomDocumentProperties(VBAshInstalledModulePropertyName(i)).Delete
        Next i
        ThisWorkbook.CustomDocumentProperties(INSTALL_MODULE_COUNT_NAME).Delete
    End If
End Sub

Private Function VBAshInstalledModulePropertyName(ByVal index As Long) As String
    VBAshInstalledModulePropertyName = INSTALL_MODULE_PROPERTY_PREFIX & Format$(index, "000")
End Function

Private Function VBAshHasDocumentProperty(ByVal propertyName As String) As Boolean
    Dim docProperty As Object

    On Error Resume Next
    Err.Clear
    Set docProperty = ThisWorkbook.CustomDocumentProperties(propertyName)
    VBAshHasDocumentProperty = (Err.Number = 0)
    If Err.Number <> 0 Then
        Err.Clear
        VBAshHasDocumentProperty = False
    End If
    On Error GoTo 0
End Function

Private Function VBAshTerminalFormCode() As String
    VBAshTerminalFormCode = VBAshFormSetupCode() & vbCrLf & _
                            VBAshFormEventsCode() & vbCrLf & _
                            VBAshFormCommandCode()
End Function

Private Function VBAshFormSetupCode() As String
    VBAshFormSetupCode = Join(Array( _
        "Option Explicit", _
        "Private WithEvents mTerminal As MSForms.TextBox", _
        "Private mInputStart As Long", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    Me.Caption = ""VBAsh terminal""", _
        "    Me.Width = 720", _
        "    Me.Height = 480", _
        "    Me.BackColor = RGB(24, 24, 24)", _
        "    Me.StartUpPosition = 1", _
        "    Set mTerminal = Me.Controls.Add(""Forms.TextBox.1"", ""txtTerminal"", True)", _
        "    With mTerminal", _
        "        .Left = 12", _
        "        .Top = 12", _
        "        .Width = Me.InsideWidth - 24", _
        "        .Height = Me.InsideHeight - 24", _
        "        .MultiLine = True", _
        "        .ScrollBars = 3", _
        "        .Locked = False", _
        "        .EnterKeyBehavior = True", _
        "        .BackColor = RGB(18, 18, 18)", _
        "        .ForeColor = RGB(220, 220, 220)", _
        "        .Font.Name = ""Consolas""", _
        "        .Font.Size = 10", _
        "        .BorderStyle = 0", _
        "    End With", _
        "    mTerminal.Value = ""VBAsh pronto. Digite help para ver os comandos."" & vbCrLf & gSession.Prompt", _
        "    mInputStart = Len(mTerminal.Value)", _
        "    mTerminal.SelStart = mInputStart", _
        "    mTerminal.SetFocus", _
        "End Sub"), vbCrLf)
End Function

Private Function VBAshFormEventsCode() As String
    VBAshFormEventsCode = Join(Array( _
        "Private Sub UserForm_Activate()", _
        "    mTerminal.SetFocus", _
        "End Sub", _
        "Private Sub mTerminal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)", _
        "    If KeyCode = vbKeyReturn Then", _
        "        KeyCode = 0", _
        "        SubmitCommand", _
        "    ElseIf mTerminal.SelStart < mInputStart Then", _
        "        mTerminal.SelStart = Len(mTerminal.Text)", _
        "        mTerminal.SelLength = 0", _
        "    ElseIf KeyCode = vbKeyBack And mTerminal.SelLength = 0 And _", _
        "           mTerminal.SelStart <= mInputStart Then", _
        "        KeyCode = 0", _
        "    ElseIf mTerminal.SelStart + mTerminal.SelLength < mInputStart Then", _
        "        mTerminal.SelLength = 0", _
        "        mTerminal.SelStart = Len(mTerminal.Text)", _
        "    End If", _
        "End Sub"), vbCrLf)
End Function

Private Function VBAshFormCommandCode() As String
    VBAshFormCommandCode = Join(Array( _
        "Private Sub SubmitCommand()", _
        "    Dim commandLine As String", _
        "    Dim result As String", _
        "    commandLine = Mid$(mTerminal.Text, mInputStart + 1)", _
        "    If Len(Trim$(commandLine)) = 0 Then Exit Sub", _
        "    mTerminal.Value = mTerminal.Text & vbCrLf", _
        "    result = gSession.ProcessInput(commandLine)", _
        "    If gSession.HasOutput Then", _
        "        mTerminal.Value = mTerminal.Text & result", _
        "        If Not gSession.SuppressNewline Then mTerminal.Value = mTerminal.Text & vbCrLf", _
        "    End If", _
        "    If gSession.ShouldExit Then", _
        "        Unload Me", _
        "    Else", _
        "        mTerminal.Value = mTerminal.Text & gSession.Prompt", _
        "        mInputStart = Len(mTerminal.Text)", _
        "        mTerminal.SelStart = mInputStart", _
        "        mTerminal.SetFocus", _
        "    End If", _
        "End Sub"), vbCrLf)
End Function

Private Function VBAshComponentExists(ByVal project As Object, ByVal moduleName As String) As Boolean
    Dim component As Object

    For Each component In project.VBComponents
        If StrComp(CStr(component.Name), moduleName, vbTextCompare) = 0 Then
            VBAshComponentExists = True
            Exit Function
        End If
    Next component
End Function

Private Function VBAshHasInstallMarker() As Boolean
    On Error Resume Next
    VBAshHasInstallMarker = _
        (CStr(ThisWorkbook.CustomDocumentProperties(INSTALL_MARKER_NAME).Value) = INSTALL_MARKER_VALUE)
    On Error GoTo 0
End Function

Private Function VBAshModuleName(ByVal fileName As String) As String
    VBAshModuleName = Left$(fileName, InStrRev(fileName, ".") - 1)
End Function

Private Function VBAshBuildPath(ByVal folder As String, ByVal fileName As String) As String
    If Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        VBAshBuildPath = folder & fileName
    Else
        VBAshBuildPath = folder & Application.PathSeparator & fileName
    End If
End Function

Private Function VBAshReadUtf8File(ByVal path As String) As String
    Dim textStream As Object

    Set textStream = CreateObject("ADODB.Stream")
    textStream.Type = 2
    textStream.Charset = "utf-8"
    textStream.Open
    textStream.LoadFromFile path
    VBAshReadUtf8File = textStream.ReadText(-1)
    textStream.Close
End Function
