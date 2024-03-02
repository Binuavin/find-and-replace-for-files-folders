''''''''''''''''''''''''''
''' Global State Variables
''''''''''''''''''''''''''
' Window Load
Sub Window_Onload
    self.Focus()

    Dim posX, posY
    ' Center the Window on screen
    posX = CInt( ( window.screen.width  - document.body.offsetWidth  ) / 2 )
    posY = CInt( ( window.screen.height - document.body.offsetHeight ) / 2 )
    If posX < 0 Then posX = 0
    If posY < 0 Then posY = 0

    ' self.moveTo posX, posY
    ' self.ResizeTo CInt( window.screen.width - 400 ), CInt( window.screen.height - 150 )

    ' Disable the context menu to prevent users from right-clicking
    ' self.body.contextMenu = "return false" 
    ' This has been replaced with CONTEXTMENU="no" in header meta

    ' Set Custom Properties in Hidden Input Boxes
    ' document.getElementById("propToCheckSubFoldersUserClick").value = "user_not_checked"

End Sub

Sub CheckKeyPress
    ' Check if Esc key (key code 27) is pressed
    If window.event.keyCode = 27 Then 
        window.close ' Close the window
    End If
End Sub


Sub OpenFolderSelector
    sBFF = PickFolder()
    If not sBFF = "" Then 
    document.PROCESS_FILE_NAMES.folderPath.value = sBFF
    End If
End Sub 

' Check and Disable Subfolder Renaming option according to Checbox Conditon
Sub subFolderCheckValue 
    Dim subFolderCheckBtn, propToCheckSubFoldersUserClick, renameFoldersCheckBox, submitButton
    Set subFolderCheckBtn = document.getElementById("includeSubfolders")
    Set propToCheckSubFoldersUserClick = document.getElementById("propToCheckSubFoldersUserClick")
    Set renameFoldersCheckBox = document.getElementById("renameFolders")
    Set submitButton = document.getElementById("runbutton")

    ' Check if the checkbox is checked
    If subFolderCheckBtn.checked Then
        propToCheckSubFoldersUserClick.value = "user_checked"
    Else 
        propToCheckSubFoldersUserClick.value = "user_not_checked"
        If renameFoldersCheckBox.checked Then 
            renameFoldersCheckBox.checked = False
            submitButton.innerText = "Rename Files"
        End If
    End If

End Sub

' Update Button Name accordign to checkbox click
Sub updateBtnNameOnClick
    Dim subFolderCheckBox, subFolderCheckHiddenValue, renameFoldersCheckBox, submitButton, subFolderSearch
    ' Get the checkbox and button elements
    Set subFolderCheckBox = document.getElementById("includeSubfolders")
    Set subFolderCheckHiddenValue = document.getElementById("propToCheckSubFoldersUserClick")

    Set renameFoldersCheckBox = document.getElementById("renameFolders")
    Set submitButton = document.getElementById("runbutton")

    ' Check to see if the SubFolderSearch checkbox is already checked
    subFolderSearch = subFolderCheckBox.checked

    ' Check if the checkbox is checked
    If renameFoldersCheckBox.checked Then
        subFolderCheckBox.checked = True
        ' Update button text when checkbox is checked
        submitButton.innerText = "Rename Files and Folders"
    Else
        ' If the user has not already check subfolderSearch then disable the renameCheckbox
        If subFolderCheckHiddenValue.value <> "user_not_checked" Then
            subFolderCheckBox.checked = True
        Else
            subFolderCheckBox.checked = False
        End If 
        ' Update button text when checkbox is unchecked
        submitButton.innerText = "Rename Files"
    End If
End Sub


' Function to Pic a folder
Function PickFolder()
    Dim shell, Folder
    Set shell = CreateObject("Shell.Application")
    Set Folder = shell.BrowseForFolder(0, "Choose a folder:" _
        , &H0001 + &H0004 + &H0010 + &H0020)
    'See MSDN "BROWSEINFO structure" for constants
    If (Not Folder Is Nothing) Then
        PickFolder = Folder.Self.Path
    Else
        PickFolder = ""
    End If
    Set shell = Nothing
    Set Folder = Nothing
End Function

' Function to validate folder path
Function ValidateFolderPath(folderPath)
    Dim fso, folder
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        Set folder = fso.GetFolder(folderPath)
        If (Not folder Is Nothing) Then
            ValidateFolderPath = True
        Else
            ValidateFolderPath = False
        End If
    Else
        ValidateFolderPath = False
    End If
End Function

' Function to display error message and exit script
Sub DisplayErrorMessage(errorMessage)
    MsgBox errorMessage, vbExclamation, "Error"
    ' WScript.Quit
    ' window.close()
    Exit Sub
End Sub

' Main subroutine to rename files
Sub RenameFiles()
    Dim folderPath, searchWord, replaceWord, includeSubfolders, searchWordMatchCase, fileExtension, changeExtension, renameFolders
    Dim fso, folder, shell, command

    ' Retrieve user inputs
    folderPath = Trim(document.getElementById("folderPath").value)
    searchWord = Trim(document.getElementById("searchWord").value)
    replaceWord = Trim(document.getElementById("replaceWord").value)
    includeSubfolders = document.getElementById("includeSubfolders").checked
    searchWordMatchCase = document.getElementById("searchWordMatchCase").checked
    fileExtension = Trim(document.getElementById("fileExtension").value)
    changeExtension = document.getElementById("changeExtension").checked
    renameFolders = document.getElementById("renameFolders").checked

    ' Validate folder input
    If folderPath = "" Then
        DisplayErrorMessage "Folder path cannot be empty."
        document.getElementById("folderPath").focus()
        Exit Sub
    End If

    ' Validate search input
    If searchWord = "" Then
        DisplayErrorMessage "Enter a search term."
        document.getElementById("searchWord").focus()
        Exit Sub
    End If

    ' Validate folder
    If Not ValidateFolderPath(folderPath) Then
        DisplayErrorMessage "Please enter a valid folder path."
        document.getElementById("folderPath").focus()
        Exit Sub
    End If

    ' construct the command to execute the VBScript
    command = "cscript //NoLogo rename_files_folders.vbs  """ & folderPath & """ """ & searchWord & """ """ & replaceWord & """ " & includeSubfolders & " " & searchWordMatchCase & " " & changeExtension & " """ & fileExtension & """ " & renameFolders

    ' Create shell object
    Set shell = CreateObject("WScript.Shell")

    ' Execute the command
    shell.Run command, 1, True

    ' Clean up
    Set shell = Nothing
End Sub

