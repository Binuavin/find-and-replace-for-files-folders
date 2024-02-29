' Functions:
' Function to rename files in a folder and its subfolders
Sub RenameFiles(folderPath, searchWord, replaceWord, includeSubfolders, searchWordMatchCase, changeExtension, fileExtension, renameFolders)
    Dim objFso, folder
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    ' Validate folder path
    If Not objFso.FolderExists(folderPath) Then
        WScript.Echo "Folder does not exist: " & folderPath
        Exit Sub
    End If
    
    ' Get the folder object
    Set folder = objFso.GetFolder(folderPath)
    
    ' Rename files in the current folder
    RenameFilesInFolder folder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension
    
    ' If includeSubfolders is true, rename files in subfolders recursively
    If includeSubfolders Then
        RenameFilesInSubfolders folder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension, renameFolders
    End If
    
    ' Add Output Message
    DisplayMessage("Success")
    
End Sub

' Subroutine to rename files in a folder
Sub RenameFilesInFolder(folder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension)
    Dim file, fileName, newFileName, fileExt
    Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
    
    ' Initialize progress counter
    Dim fileCount, currentCount
    fileCount = folder.Files.Count
    currentCount = 0

    ' Iterate through each file in the folder
    For Each file In folder.Files
        fileName = file.Name
        fileExt = objFso.GetExtensionName(fileName)
        
        ' Check to only match the word with the Case Entered
        If searchWordMatchCase Then  
            If InStr(fileName, searchWord) > 0 Then
                If replaceWord <> "" Then ' Check if replaceWord is not empty
                    newFileName = Replace(fileName, searchWord, replaceWord)
                Else
                    newFileName = Replace(fileName, searchWord, searchWord) ' Replace with itself as the replaceWord is empty
                End If
                If changeExtension Then
                    newFileName = Left(newFileName, InStrRev(newFileName, ".")) & fileExtension
                End If
                file.Move objFso.BuildPath(folder.Path, newFileName)
            End If
        Else
            ' Check if the search word (case insensitive) is found in the filename
            If InStr(LCase(fileName), LCase(searchWord)) > 0 Then
                Dim regex, match
                Set regex = New RegExp

                ' Set the regular expression pattern to match the search word (case insensitive)
                regex.IgnoreCase = True
                regex.Global = True
                regex.Pattern = "\b" & EscapeRegExp(searchWord) & "\b"

                ' Check if the search word matches in the filename
                Set match = regex.Execute(fileName)
                If match.Count > 0 Then
                    ' Replace the search word with the replace word
                    If replaceWord <> "" Then ' Check if replaceWord is not empty
                        ' Replace the search word with the replace word
                        newFileName = regex.Replace(fileName, replaceWord)
                    Else
                        ' Replace with itself as the replaceWord is empty
                        newFileName = regex.Replace(fileName, searchWord)
                    End If

                    If changeExtension Then
                        newFileName = Left(newFileName, InStrRev(newFileName, ".")) & fileExtension
                    End If
                    ' Move the file with the new filename
                    file.Move objFso.BuildPath(folder.Path, newFileName)
                End If
            End If
        End If

        ' Update progress
        currentCount = currentCount + 1
    Next
    ' UpdateProgress currentCount, fileCount
End Sub

' Subroutine to rename files in subfolders and Folder names recursively
Sub RenameFilesInSubfolders(folder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension, renameFolders)
    Dim subfolder, objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")

    ' Iterate through each subfolder
    For Each subfolder In folder.SubFolders
        ' Rename files in the subfolder
        RenameFilesInFolder subfolder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension
        
        If renameFolders Then
            ' Rename the subfolder itself if needed
            Dim newFolderName
            If searchWordMatchCase Then
                If InStr(subfolder.Name, searchWord) > 0 Then
                    If replaceWord <> "" Then ' Check if replaceWord is not empty
                        newFolderName = Replace(subfolder.Name, searchWord, replaceWord)
                    Else
                        newFolderName = Replace(subfolder.Name, searchWord, searchWord) ' Replace with itself as the replaceWord is empty
                    End If
                    subfolder.Name = newFolderName
                End If
            Else
                Dim regex, match
                Set regex = New RegExp
                regex.IgnoreCase = True
                regex.Global = True
                regex.Pattern = "\b" & EscapeRegExp(searchWord) & "\b"
                If InStr(LCase(subfolder.Name), LCase(searchWord)) > 0 Then
                    Set match = regex.Execute(subfolder.Name)
                    If match.Count > 0 Then

                        ' Replace the search word with the replace word
                        If replaceWord <> "" Then ' Check if replaceWord is not empty
                            ' Replace the search word with the replace word
                            newFolderName = regex.Replace(subfolder.Name, replaceWord)
                        Else
                            ' Replace with itself as the replaceWord is empty
                            newFolderName = regex.Replace(subfolder.Name, searchWord)
                        End If
                        subfolder.Name = newFolderName
                    End If
                End If
            End If
        End If
        
        ' Recursive call to rename files in subfolders of subfolder
        RenameFilesInSubfolders subfolder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension, renameFolders
    Next
End Sub


' Function to regex filter the common symbol strings
Function EscapeRegExp(str)
    EscapeRegExp = Replace(str, "\\", "\\\\")
    EscapeRegExp = Replace(EscapeRegExp, "^", "\\^")
    EscapeRegExp = Replace(EscapeRegExp, "$", "\\$")
    EscapeRegExp = Replace(EscapeRegExp, ".", "\\.")
    EscapeRegExp = Replace(EscapeRegExp, "|", "\\|")
    EscapeRegExp = Replace(EscapeRegExp, "(", "\\(")
    EscapeRegExp = Replace(EscapeRegExp, ")", "\\)")
    EscapeRegExp = Replace(EscapeRegExp, "[", "\\[")
    EscapeRegExp = Replace(EscapeRegExp, "]", "\\]")
    EscapeRegExp = Replace(EscapeRegExp, "*", "\\*")
    EscapeRegExp = Replace(EscapeRegExp, "+", "\\+")
    EscapeRegExp = Replace(EscapeRegExp, "?", "\\?")
    EscapeRegExp = Replace(EscapeRegExp, "/", "\\/")
End Function

' Subroutine to update progress
Sub UpdateProgress(current, total)
    Dim percentDone
    percentDone = Int((current / total) * 100)
    WScript.StdOut.Write "Progress: " & percentDone & "% " & vbCrLf
End Sub

' Display Output Message
Function DisplayMessage(condition)
    Select Case condition 
    Case "Success"
        ' WScript.Echo "Files renamed successfully!"
        MsgBox "Files renamed successfully.", vbInformation, "Confirmation"
    Case "No File"
        ' WScript.Echo "No files found with search!"
        MsgBox "No files found with search!", vbError, "No File"
    Case "Fail"
        WScript.Echo "Files renaming failed!"
    End Select
End Function

' Main code to read input from command line arguments and execute renaming
Dim folderPath, searchWord, replaceWord, includeSubfolders, searchWordMatchCase, changeExtension, fileExtension, renameFolders

' Get command line arguments
If WScript.Arguments.Count <> 8 Then
    WScript.Echo "Invalid number of arguments!"
    WScript.Quit
End If

folderPath = WScript.Arguments(0)
searchWord = WScript.Arguments(1)
replaceWord = WScript.Arguments(2)
includeSubfolders = CBool(WScript.Arguments(3))
searchWordMatchCase = CBool(WScript.Arguments(4))
changeExtension = CBool(WScript.Arguments(5))
fileExtension = WScript.Arguments(6)
renameFolders = CBool(WScript.Arguments(7))

' Execute renaming
RenameFiles folderPath, searchWord, replaceWord, includeSubfolders, searchWordMatchCase, changeExtension, fileExtension, renameFolders




' Error Handling
' On Error Resume Next
' subfolder.Name = newFolderName
' If Err.Number <> 0 Then
'     MsgBox "Error renaming folder: " & Err.Description, vbExclamation, "Error"
' End If
' On Error GoTo 0
