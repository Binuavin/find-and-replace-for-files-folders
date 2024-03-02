' Functions:
' Function to rename files in a folder and its subfolders
Sub RenameFiles(folderPath, searchWord, replaceWord, includeSubfolders, searchWordMatchCase, changeExtension, fileExtension, renameFolders)
    Dim objFso, folder
    Set objFso = CreateObject("Scripting.FileSystemObject")

    Dim dt : dt = Now()
    Dim datetime : datetime = Year(dt) & "-" & Month(dt) & "-" & Day(dt) & "_" & Hour(dt) & "-" & Minute(dt) & "-" & Second(dt)

    Dim outputFile : outputFile = folderpath & "\output-" & datetime & ".txt"

    ' Validate folder path
    If Not objFso.FolderExists(folderPath) Then
        MsgBox "Folder does not exist: " & folderPath, vbExclamation, "Error"
        Exit Sub
    End If

    ' Get the folder object
    Set folder = objFso.GetFolder(folderPath)

    LogToFile outputFile, "Info || Starting Search at: " & dt & vbCrLf

    ' Rename files in the current folder
    RenameFilesInFolder outputFile, folder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension

    ' If includeSubfolders is true, rename files in subfolders recursively
    If includeSubfolders Then
        RenameFilesInSubfolders outputFile, folder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension, renameFolders
    End If

    ' Add Output Message
    DisplayMessage "Success", outputFile

End Sub

' Subroutine to rename files in a folder
Sub RenameFilesInFolder(outputFile, folder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension)
    Dim file, fileName, newFileName, fileExt
    Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
    
    ' Initialize progress counter
    Dim currentCount
    ' Dim fileCount : fileCount = folder.Files.Count
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

                ' Log the change to the output file
                LogToFile outputFile, "File || """ & folder.Path & "\" & fileName & """ || " & newFileName
                currentCount = currentCount + 1
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

                    ' Log the change to the output file
                    LogToFile outputFile, "File || """ & folder.Path & "\" & fileName & """ || " & newFileName
                    currentCount = currentCount + 1 
                End If
            End If
        End If

        ' Update progress
        ' currentCount = currentCount + 1
        ' UpdateProgress currentCount, fileCount
    Next
    If currentCount = 0 Then
        LogToFile outputFile, "Note || """ & folder.Path & """ || No files found with the search."
    End If
End Sub

' Subroutine to rename files in subfolders and Folder names recursively
Sub RenameFilesInSubfolders(outputFile, folder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension, renameFolders)
    Dim subfolder, objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")

    ' Initialize progress counter
    Dim currentSubCount
    ' Dim fileCount : fileCount = folder.Files.Count
    currentSubCount = 0

    ' Iterate through each subfolder
    For Each subfolder In folder.SubFolders

        Dim originalFolderName : originalFolderName = subfolder.Name
        Dim originalFolderPath : originalFolderPath = subfolder.Path
        ' Rename files in the subfolder
        RenameFilesInFolder outputFile, subfolder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension

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

                    ' Log the change to the output file  
                    LogToFile outputFile, "Folder || """ & originalFolderPath & """ || " & subfolder.Name
                    currentSubCount = currentSubCount + 1
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

                        ' Log the change to the output file 
                        LogToFile outputFile, "Folder || """ & originalFolderPath & "\" & originalFolderName & """ || " & newFolderName
                        currentSubCount = currentSubCount + 1
                    End If
                End If
            End If
        End If

        ' Recursive call to rename files in subfolders of subfolder
        RenameFilesInSubfolders outputFile, subfolder, searchWord, replaceWord, searchWordMatchCase, changeExtension, fileExtension, renameFolders

        If currentSubCount = 0 Then
            LogToFile outputFile, "Note || No files found with the search in " & originalFolderName
        End If
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

' Function to log changes to an output file
Sub LogToFile(filePath, logEntry)
    Dim fs, outFile
    Set fs = CreateObject("Scripting.FileSystemObject")

    ' Check if the file already exists, if not, create a new one
    If Not fs.FileExists(filePath) Then
        Set outFile = fs.CreateTextFile(filePath, True)
    Else
        Set outFile = fs.OpenTextFile(filePath, 8, True) ' 8 = FileAppend
    End If

    ' Write the log entry to the file
    outFile.Write logEntry & vbCrLf

    ' Close the file
    outFile.Close
End Sub

' Display Output Message
Sub DisplayMessage(condition, outputFile)
    Select Case condition 
    Case "Success"
        ' WScript.Echo "Files renamed successfully!"
        MsgBox "Files renamed successfully. Log saved to: " & outputFile, vbInformation, "Confirmation"
    Case "No File"
        ' WScript.Echo "No files found with search!"
        MsgBox "No files found with search!", vbExclamation, "No File"
    Case "Fail"
        ' WScript.Echo "Files renaming failed!"
        MsgBox "File Renaming Failed", vbError, "Error"
    End Select
End Sub

' Main code to read input from command line arguments and execute renaming
Dim folderPath, searchWord, replaceWord, includeSubfolders, searchWordMatchCase, changeExtension, fileExtension, renameFolders

' Get command line arguments
If WScript.Arguments.Count <> 8 Then
    ' WScript.Echo "Invalid number of arguments!"
    MsgBox "Invalid number of arguments!", vbError, "Error"
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
