Module libFSys

    Public Function nuFso() As Scripting.FileSystemObject
        nuFso = New Scripting.FileSystemObject
    End Function

    Public Function fdUserHome() As Scripting.Folder
        fdUserHome = nuFso.GetFolder(Environ("AppData")).ParentFolder
    End Function

    Public Function fnExt(fn As String) As String
        Dim ar As Object
        Dim mx As Long

        ar = Split(fn, ".")
        mx = UBound(ar)
        If mx < 0 Then
            fnExt = ""
        Else
            fnExt = ar(mx)
        End If
    End Function

    Public Function folderIfPresent(Path As String,
    Optional Base As Scripting.Folder = Nothing
) As Scripting.Folder
        Dim rt As Scripting.Folder

        If Base Is Nothing Then
            With nuFso()
                If .FolderExists(Path) Then
                    rt = .GetFolder(Path)
                Else
                    rt = Nothing
                End If
            End With
        Else
            rt = folderIfPresent(Base.Path & "\" & Path)
        End If

        folderIfPresent = rt
    End Function

    Public Function fileIfPresent(Path As String,
    Optional Base As Scripting.Folder = Nothing
) As Scripting.File
        Dim rt As Scripting.File

        If Base Is Nothing Then
            With nuFso()
                If .FileExists(Path) Then
                    rt = .GetFile(Path)
                Else
                    rt = Nothing
                End If
            End With
        Else
            rt = fileIfPresent(Base.Path & "\" & Path)
        End If

        fileIfPresent = rt
    End Function

    Public Function dcFilesIn(fd As Scripting.Folder) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim fl As Scripting.File

        rt = New Scripting.Dictionary
        If Not fd Is Nothing Then
            For Each fl In fd.Files
                rt.Add(fl.Name, fl)
            Next
        End If
        dcFilesIn = rt
    End Function
    'send2clipBd dcFilesIn(nuFso.GetFolder(""))
    'send2clipBd Join(dcFoldersIn(nuFso.GetFolder("C:\Doyle_Vault\Designs\doyle")).Keys, vbCrLf)
    'send2clipBd Join(dcFilesIn(nuFso.GetFolder("W:\Parts Lists")).Keys, vbCrLf)

    Public Function dcFoldersIn(
    fd As Scripting.Folder
) As Scripting.Dictionary
        Dim rt As Scripting.Dictionary
        Dim fl As Scripting.Folder

        rt = New Scripting.Dictionary
        If Not fd Is Nothing Then
            For Each fl In fd.SubFolders
                rt.Add(fl.Name, fl)
            Next
        End If
        dcFoldersIn = rt
    End Function
End Module