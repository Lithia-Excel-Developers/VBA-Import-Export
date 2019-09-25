VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModuleHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'// Import a VBA code module... how hard could it be right?
Public Sub ImportModule(ByVal Project As VBProject, ByVal moduleName As String, ByVal ModulePath As String)

    Dim comNewImport        As VBComponent
    Dim comExistingComp     As VBComponent
    Dim modCodeCopy         As CodeModule
    Dim modCodePaste        As CodeModule

    Set comNewImport = Project.VBComponents.Import(ModulePath)
    If comNewImport.Name <> moduleName Then
        If CollectionKeyExists(Project.VBComponents, moduleName) Then

            Set comExistingComp = Project.VBComponents(moduleName)
            If comExistingComp.Type = vbext_ct_Document Then

                Set modCodeCopy = comNewImport.CodeModule
                Set modCodePaste = comExistingComp.CodeModule
                modCodePaste.DeleteLines 1, modCodePaste.CountOfLines
                If modCodeCopy.CountOfLines > 0 Then
                    modCodePaste.AddFromString modCodeCopy.Lines(1, modCodeCopy.CountOfLines)
                End If
                Project.VBComponents.Remove comNewImport

            Else

                Project.VBComponents.Remove comExistingComp
                comNewImport.Name = moduleName

            End If
        Else

            comNewImport.Name = moduleName

        End If
    End If

End Sub


'// Is the given module exportable by this tool?
Public Function ExportableModule(ByVal comModule As VBComponent) As Boolean

    ExportableModule = _
        (Not ModuleEmpty(comModule)) And _
        (Not FileExtension(comModule) = vbNullString)

End Function


'// Check if a code module is effectively empty.
'// effectively empty should be functionally and semantically equivelent to
'// actually empty.
Public Function ModuleEmpty(ByVal comModule As VBComponent) As Boolean

    Dim lngNumLines As Long
    Dim lngCurLine As Long
    Dim strCurLine As String

    ModuleEmpty = True

    lngNumLines = comModule.CodeModule.CountOfLines
    For lngCurLine = 1 To lngNumLines
        strCurLine = comModule.CodeModule.Lines(lngCurLine, 1)
        If Not (strCurLine = "Option Explicit" Or strCurLine = "") Then
            ModuleEmpty = False
            Exit Function
        End If
    Next lngCurLine

End Function


'// The appropriate file extension for exporting the given module
Public Function FileExtension(ByVal comModule As VBComponent) As String

    Select Case comModule.Type
        Case vbext_ct_StdModule
            FileExtension = "bas"
        Case vbext_ct_ClassModule, vbext_ct_Document
            FileExtension = "cls"
        Case vbext_ct_MSForm
            FileExtension = "frm"
        Case Else
            FileExtension = vbNullString
    End Select

End Function


'// Ensure path to a file exists. Creates missing folders.
Public Sub EnsurePath(ByVal Path As String)

    Dim strParentPath As String

    Set FSO = New Scripting.FileSystemObject
    strParentPath = FSO.GetParentFolderName(Path)

    If Not strParentPath = "" Then
        EnsurePath strParentPath
        If Not FSO.FolderExists(strParentPath) Then
            If FSO.FileExists(strParentPath) Then
                Err.Raise vbObjectError + 1, "modImportExport:EnsurePath", "A file exists where a folder needs to be: " & strParentPath
            Else
                FSO.CreateFolder (strParentPath)
            End If
        End If
    End If

End Sub

