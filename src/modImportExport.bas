Attribute VB_Name = "modImportExport"
Option Explicit

'// Updates the configuration file for the current active project.
'// * Entries for modules not yet declared in the configuration file as created.
'// * Modules listed in the configuration file which are not found are prompted
'//   to be deleted from the configuration file.
'// * The current loaded references are used to update the configuration file.
'// * References in the configuration file whic hare not loaded are prompted to
'//   be deleted from the configuration file.
Public Sub MakeConfigFile()

    Dim prjActProj          As VBProject
    Dim Config              As clsConfiguration

    Dim comModule           As VBComponent
    Dim boolDeleteModule    As Boolean
    Dim boolCreateNewEntry  As Boolean
    Dim varModuleName       As Variant
    Dim strModuleName       As String

    Dim refReference        As Reference
    Dim lngIndex            As Long
    Dim varIndex            As Variant
    Dim boolForbiddenRef    As Boolean

    Dim collDeleteList      As Collection
    Dim strDeleteListStr    As String
    Dim intUserResponse     As Integer

    On Error GoTo catchError

    Set prjActProj = Application.VBE.ActiveVBProject
    If prjActProj Is Nothing Then GoTo exitSub

    Set Config = New clsConfiguration
    Config.Project = prjActProj
    Config.ReadFromProjectConfigFile

    '// Generate entries for modules not yet listed
    For Each comModule In prjActProj.VBComponents
        boolCreateNewEntry = _
            ModuleHandler.ExportableModule(comModule) And _
            InStr(1, comModule.Name, Config.TestModuleSuffix) = 0 And _
            Not Config.ModuleDeclared(comModule.Name)

        If boolCreateNewEntry Then
            Config.ModulePath(comModule.Name) = comModule.Name & "." & ModuleHandler.FileExtension(comModule)
        End If
    Next comModule

    '// Ask user if they want to delete entries for missing modules
    Set collDeleteList = New Collection
    strDeleteListStr = ""
    For Each varModuleName In Config.ModuleNames
        strModuleName = varModuleName
        boolDeleteModule = True
        If CollectionKeyExists(prjActProj.VBComponents, strModuleName) Then
            If ModuleHandler.ExportableModule(prjActProj.VBComponents(strModuleName)) Then
                boolDeleteModule = False
            End If
        End If
        If boolDeleteModule Then
            collDeleteList.Add strModuleName
            strDeleteListStr = strDeleteListStr & strModuleName & vbNewLine
        End If
    Next varModuleName

    If collDeleteList.Count > 0 Then
        intUserResponse = MsgBox( _
            Prompt:= _
                "There are some modules listed in the configuration file which " & _
                "haven't been found in the current project. Would you like to " & _
                "remove these modules from the configuration file?" & vbNewLine & _
                vbNewLine & _
                "Missing modules:" & vbNewLine & _
                strDeleteListStr, _
            Buttons:=vbYesNo + vbDefaultButton2, _
            Title:="Missing Modules")

        If intUserResponse = vbYes Then
            For Each varModuleName In collDeleteList
                strModuleName = varModuleName
                Config.ModulePathRemove strModuleName
            Next varModuleName
        End If
    End If

    '// Generate entries for references in the current VBProject
    For Each refReference In prjActProj.References
        If Not refReference.BuiltIn Then
            boolForbiddenRef = _
                refReference.Name = "stdole" Or _
                refReference.Name = "Office" Or _
                refReference.Name = "MSForms"
            If Not boolForbiddenRef Then
                Config.ReferencesUpdateFromVBRef refReference
            End If
        End If
    Next refReference

    '// Prompt if entries for missing references should be deleted
    Set collDeleteList = New Collection
    strDeleteListStr = ""
    For lngIndex = Config.ReferencesCount To 1 Step -1
        If Not CollectionKeyExists(prjActProj.References, Config.ReferenceName(lngIndex)) Then
            collDeleteList.Add lngIndex
            strDeleteListStr = strDeleteListStr & Config.ReferenceName(lngIndex) & vbNewLine
        End If
    Next

    If collDeleteList.Count > 0 Then
        intUserResponse = MsgBox( _
            Prompt:= _
                "There are some references listed in the configuration file which " & _
                "haven't been found in the current project. Would you like to " & _
                "remove these references from the configuration file?" & vbNewLine & _
                vbNewLine & _
                "Missing references:" & vbNewLine & _
                strDeleteListStr, _
            Buttons:=vbYesNo + vbDefaultButton2, _
            Title:="Missing References")

        If intUserResponse = vbYes Then
            For Each varIndex In collDeleteList
                lngIndex = varIndex
                Config.ReferenceRemove lngIndex
            Next varIndex
        End If
    End If

    '// Write changes to config file
    Config.WriteToProjectConfigFile

    MsgBox _
        "Configuration file was successfully updated. Please review the " & _
        "file with a text editor."

exitSub:
    Exit Sub

catchError:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub

'// Exports code modules and cleans the current active VBProject as specified
'// by the project's configuration file.
'// * Any code module in the VBProject which is listed in the configuration
'//   file is exported to the configured path.
'// * code modules which were exported are deleted or cleared.
'// * References loaded in the Project which are listed in the configuration
'//   file is deleted.
Public Sub Export(Optional RemoveFromProject As Boolean = True)

    Dim prjActProj          As VBProject
    Dim Config              As clsConfiguration
    Dim comModule           As VBComponent
    Dim lngIndex            As Long
    Dim strModuleName       As String
    Dim varModuleName       As Variant
    
    On Error GoTo ErrHandler

    Set prjActProj = Application.VBE.ActiveVBProject
    If prjActProj Is Nothing Then GoTo exitSub

    Set Config = New clsConfiguration
    Config.Project = prjActProj
    Config.ReadFromProjectConfigFile

    '// Export all modules listed in the configuration
    For Each varModuleName In Config.ModuleNames
        strModuleName = varModuleName
        If CollectionKeyExists(prjActProj.VBComponents, strModuleName) Then
            Set comModule = prjActProj.VBComponents(strModuleName)
            ModuleHandler.EnsurePath Config.ModuleFullPath(strModuleName)
                
            comModule.Export Config.ModuleFullPath(strModuleName)
            
            If RemoveFromProject Then
                If comModule.Type = vbext_ct_Document Then
                    comModule.CodeModule.DeleteLines 1, comModule.CodeModule.CountOfLines
                Else
                    prjActProj.VBComponents.Remove comModule
                End If
            End If
        Else
            ' TODO Provide a warning if module listed in configuration is not found
        End If
    Next varModuleName

    '// Remove all references listed
    For lngIndex = 1 To Config.ReferencesCount
        If RemoveFromProject And CollectionKeyExists(prjActProj.References, Config.ReferenceName(lngIndex)) Then
            prjActProj.References.Remove prjActProj.References(Config.ReferenceName(lngIndex))
        End If
    Next lngIndex

exitSub:
    Exit Sub

ErrHandler:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub

'// Exports code modules without cleaning the current active
'// VBProject as specifiedby the project's configuration file.
Public Sub Save()
    Export False
End Sub

'// Imports textual data from the file system such as VBA code to build the
'// current active VBProject as specified in it's configuration file.
'// * Each code module file listed in the configuration file is imported into
'//   the VBProject. Modules with the same name are overwritten.
'// * All references declared in the configuration file are loaded into the
'//   project.
'// * The project name is set to the project name specified by the configuration
'//   file.
Public Sub Import()

    Dim prjActProj          As VBProject
    Dim Config              As clsConfiguration
    Dim strModuleName       As String
    Dim varModuleName       As Variant
    
    On Error GoTo catchError

    Set prjActProj = Application.VBE.ActiveVBProject
    If prjActProj Is Nothing Then GoTo exitSub

    Set Config = New clsConfiguration
    Config.Project = prjActProj
    Config.ReadFromProjectConfigFile

    '// Import code from listed module files
    For Each varModuleName In Config.ModuleNames
        strModuleName = varModuleName
        ModuleHandler.ImportModule prjActProj, strModuleName, Config.ModuleFullPath(strModuleName)
    Next varModuleName

    '// Add references listed in the config file
    Config.ReferencesAddToVBRefs prjActProj.References

    '// Set the VBA Project name
    If Config.VBAProjectNameDeclared Then
        prjActProj.Name = Config.VBAProjectName
    End If

exitSub:
    Exit Sub

catchError:
    If HandleCrash(Err.Number, Err.Description, Err.Source) Then
        Stop
        Resume
    End If
    GoTo exitSub

End Sub
