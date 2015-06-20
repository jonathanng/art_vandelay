Option Explicit

Private Const DOCUMENT_FOLDER = "Archive\"
Private Const VBACODE_FOLDER = "VBACode\"
Private Const TEMP_ZIP = "\temp.zip"
Private Const EXTENSION = "xlsm"

Const vbext_ct_ClassModule = 2
Const vbext_ct_Document = 100
Const vbext_ct_MSForm = 3
Const vbext_ct_StdModule = 1

Public Sub cleanUp()
    Dim VBComp As VBComponent
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        If Right(VBComp.Name, 1) = "1" Then VBComp.Name = Left(VBComp.Name, Len(VBComp.Name) - 1)
    Next

    Debug.Print "Cleanup done!"

End Sub

Public Sub export()

    'http://www.pretentiousname.com/excel_extractvba/index.html

    Dim VBComp As VBComponent
    Dim path As String

    ensureDir folder
    ensureDir folder & VBACODE_FOLDER

    For Each VBComp In ActiveWorkbook.VBProject.VBComponents

        If ext(VBComp) <> "" Then
            ensureDir folder & VBACODE_FOLDER & subfolder(VBComp)
            path = folder & VBACODE_FOLDER & subfolder(VBComp) & "\" & VBComp.Name & ext(VBComp)
            Debug.Print "Exporting " & path
            VBComp.export path
        End If

    Next

    If Replace(ActiveWorkbook.Name, EXTENSION, "") = ".xlsm" Then

        Dim fs As New FileSystemObject
        fs.copyFile source:=ActiveWorkbook.path & "\" & ActiveWorkbook.Name, destination:=ActiveWorkbook.path & TEMP_ZIP, overwritefiles:=True

        unzip destination:=folder & DOCUMENT_FOLDER, zipFileName:=ActiveWorkbook.path & TEMP_ZIP

        fs.DeleteFile filespec:=ActiveWorkbook.path & TEMP_ZIP, force:=True

    End If

    Debug.Print "Exporting done!"

End Sub

Public Sub import()

    cleanUp

    ' deletes all modules and classes

    Dim VBComp As VBComponent
    Dim path As String

    For Each VBComp In ActiveWorkbook.VBProject.VBComponents

        path = ""

        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                If VBComp.Name <> "VersionController" Then
                    path = folder & VBACODE_FOLDER & subfolder(VBComp) & "\" & VBComp.Name & ext(VBComp)
                End If
            Case vbext_ct_StdModule
                If VBComp.Name <> "VersionControl" Then
                    path = folder & VBACODE_FOLDER & subfolder(VBComp) & "\" & VBComp.Name & ext(VBComp)
                End If
        End Select

        If path <> "" Then
            Debug.Print "Importing " & VBComp.Name
            ActiveWorkbook.VBProject.VBComponents.import path
            ActiveWorkbook.VBProject.VBComponents.Remove VBComp
        End If

    Next

    Debug.Print "Importing done!"

End Sub

Private Function ext(VBComp As VBComponent) As String

    Select Case VBComp.Type
        Case vbext_ct_ClassModule:     ext = ".cls"
        Case vbext_ct_Document:        ext = ".cls"
        Case vbext_ct_MSForm:          ext = ".frm"
        Case vbext_ct_StdModule:       ext = ".bas"
        Case Else:                     ext = ""
    End Select

End Function

Private Function folder() As String

    folder = ActiveWorkbook.path & "\"

End Function

Private Function subfolder(VBComp As VBComponent) As String

    Select Case VBComp.Type
         Case vbext_ct_ClassModule:     subfolder = "Classes"
         Case vbext_ct_Document:        subfolder = "Documents"
         Case vbext_ct_MSForm:          subfolder = "Forms"
         Case vbext_ct_StdModule:       subfolder = "Modules"
         Case Else:                     subfolder = ""
    End Select

End Function

