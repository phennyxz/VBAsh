Attribute VB_Name = "modVBAshInstaller"
Option Explicit

Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_CLASSMODULE As Long = 2
Private Const VBEXT_CT_MSFORM As Long = 3
Private Const MSO_FILE_DIALOG_FOLDER_PICKER As Long = 4
Private Const MSO_PROPERTY_TYPE_STRING As Long = 4
Private Const INSTALL_MARKER_NAME As String = "VBAshInstallerVersion"
Private Const INSTALL_MARKER_VALUE As String = "1"

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
    sourceFiles = VBAshSourceFiles()

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
        "Isso removerá do projeto deste arquivo os módulos VBAsh listados pelo instalador e o formulário frmVBAsh." & _
        vbCrLf & "O módulo modVBAshInstaller não será removido." & vbCrLf & vbCrLf & _
        "Deseja continuar?", vbExclamation + vbYesNo + vbDefaultButton2, "Remover VBAsh")
    If answer <> vbYes Then Exit Sub

    On Error GoTo ProjectAccessError
    Set project = ThisWorkbook.VBProject

    For Each moduleName In VBAshModuleNames()
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

Private Function VBAshSourceFiles() As Variant
    Dim files(0 To 25) As String

    files(0) = "VBAsh.bas"
    files(1) = "VBAshTests.bas"
    files(2) = "IBuiltin.cls"
    files(3) = "clsBuiltinCd.cls"
    files(4) = "clsBuiltinEcho.cls"
    files(5) = "clsBuiltinExit.cls"
    files(6) = "clsBuiltinExport.cls"
    files(7) = "clsBuiltinHelp.cls"
    files(8) = "clsBuiltinLs.cls"
    files(9) = "clsBuiltinMkdir.cls"
    files(10) = "clsBuiltinPwd.cls"
    files(11) = "clsBuiltinRegistry.cls"
    files(12) = "clsBuiltinSetEnv.cls"
    files(13) = "clsBuiltinTest.cls"
    files(14) = "clsBuiltinUnset.cls"
    files(15) = "clsCommand.cls"
    files(16) = "clsCommandResolver.cls"
    files(17) = "clsConsole.cls"
    files(18) = "clsFileDescriptorTable.cls"
    files(19) = "clsFileStream.cls"
    files(20) = "clsParser.cls"
    files(21) = "clsRedirection.cls"
    files(22) = "clsSession.cls"
    files(23) = "clsStream.cls"
    files(24) = "clsToken.cls"
    files(25) = "clsTokenizer.cls"
    VBAshSourceFiles = files
End Function

Private Function VBAshModuleNames() As Variant
    Dim files As Variant
    Dim names() As String
    Dim i As Long

    files = VBAshSourceFiles()
    ReDim names(LBound(files) To UBound(files))
    For i = LBound(files) To UBound(files)
        names(i) = VBAshModuleName(CStr(files(i)))
    Next i
    VBAshModuleNames = names
End Function

Private Function VBAshTerminalFormCode() As String
    VBAshTerminalFormCode = VBAshFormSetupCode() & vbCrLf & _
                            VBAshFormOutputCode() & vbCrLf & _
                            VBAshFormPromptCode() & vbCrLf & _
                            VBAshFormInputCode() & vbCrLf & _
                            VBAshFormEventsCode() & vbCrLf & _
                            VBAshFormCommandCode()
End Function

Private Function VBAshFormSetupCode() As String
    VBAshFormSetupCode = Join(Array( _
        "Option Explicit", _
        "Private WithEvents mCommandInput As MSForms.TextBox", _
        "Private mOutput As MSForms.TextBox", _
        "Private mPrompt As MSForms.Label", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    Me.Caption = ""VBAsh terminal""", _
        "    Me.Width = 720", _
        "    Me.Height = 480", _
        "    Me.BackColor = RGB(24, 24, 24)", _
        "    Me.StartUpPosition = 1"), vbCrLf)
End Function

Private Function VBAshFormOutputCode() As String
    VBAshFormOutputCode = Join(Array( _
        "    Set mOutput = Me.Controls.Add(""Forms.TextBox.1"", ""txtOutput"", True)", _
        "    With mOutput", _
        "        .Left = 12", _
        "        .Top = 12", _
        "        .Width = Me.InsideWidth - 24", _
        "        .Height = Me.InsideHeight - 96", _
        "        .MultiLine = True", _
        "        .ScrollBars = 2", _
        "        .Locked = True", _
        "        .EnterKeyBehavior = True", _
        "        .BackColor = RGB(18, 18, 18)", _
        "        .ForeColor = RGB(220, 220, 220)", _
        "        .Font.Name = ""Consolas""", _
        "        .Font.Size = 10", _
        "        .BorderStyle = 0", _
        "    End With"), vbCrLf)
End Function

Private Function VBAshFormPromptCode() As String
    VBAshFormPromptCode = Join(Array( _
        "    Set mPrompt = Me.Controls.Add(""Forms.Label.1"", ""lblPrompt"", True)", _
        "    With mPrompt", _
        "        .Left = 12", _
        "        .Top = Me.InsideHeight - 72", _
        "        .Width = Me.InsideWidth - 24", _
        "        .Height = 20", _
        "        .Caption = gSession.Prompt", _
        "        .BackColor = RGB(24, 24, 24)", _
        "        .ForeColor = RGB(90, 220, 130)", _
        "        .Font.Name = ""Consolas""", _
        "        .Font.Size = 10", _
        "    End With"), vbCrLf)
End Function

Private Function VBAshFormInputCode() As String
    VBAshFormInputCode = Join(Array( _
        "    Set mCommandInput = Me.Controls.Add(""Forms.TextBox.1"", ""txtCommand"", True)", _
        "    With mCommandInput", _
        "        .Left = 12", _
        "        .Top = Me.InsideHeight - 46", _
        "        .Width = Me.InsideWidth - 24", _
        "        .Height = 24", _
        "        .BackColor = RGB(24, 24, 24)", _
        "        .ForeColor = RGB(245, 245, 245)", _
        "        .Font.Name = ""Consolas""", _
        "        .Font.Size = 10", _
        "        .BorderStyle = 0", _
        "    End With", _
        "    mOutput.Value = ""VBAsh pronto. Digite help para ver os comandos."" & vbCrLf", _
        "End Sub"), vbCrLf)
End Function

Private Function VBAshFormEventsCode() As String
    VBAshFormEventsCode = Join(Array( _
        "Private Sub UserForm_Activate()", _
        "    mCommandInput.SetFocus", _
        "End Sub", _
        "Private Sub mCommandInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)", _
        "    If KeyCode = vbKeyReturn Then", _
        "        KeyCode = 0", _
        "        SubmitCommand", _
        "    End If", _
        "End Sub"), vbCrLf)
End Function

Private Function VBAshFormCommandCode() As String
    VBAshFormCommandCode = Join(Array( _
        "Private Sub SubmitCommand()", _
        "    Dim commandLine As String", _
        "    Dim result As String", _
        "    commandLine = mCommandInput.Text", _
        "    mCommandInput.Text = """"", _
        "    If Len(Trim$(commandLine)) = 0 Then Exit Sub", _
        "    mOutput.Value = mOutput.Value & mPrompt.Caption & commandLine & vbCrLf", _
        "    result = gSession.ProcessInput(commandLine)", _
        "    If gSession.HasOutput Then", _
        "        mOutput.Value = mOutput.Value & result", _
        "        If Not gSession.SuppressNewline Then mOutput.Value = mOutput.Value & vbCrLf", _
        "    End If", _
        "    mPrompt.Caption = gSession.Prompt", _
        "    mOutput.SelStart = Len(mOutput.Value)", _
        "    If gSession.ShouldExit Then Unload Me", _
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
