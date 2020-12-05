Attribute VB_Name = "Module1"
Dim rootFolderPath As String
Dim rootFolder As Outlook.folder
Dim duplicateRootFolderPath As String

Public Sub Start()
  Dim folder As Outlook.MAPIFolder
  Dim EditSubfoldersOnly As Boolean

  'Select start folder
  Set folder = Application.Session.PickFolder
  
  If Not folder Is Nothing Then
  
    rootFolderPath = "\\" & folder.Store.GetRootFolder
    Set rootFolder = GetFolder(rootFolderPath)
    
    duplicateRootFolderPath = rootFolderPath & "\Duplicates"
    CreateFolder (duplicateRootFolderPath)
    duplicateRootFolder = GetFolder(duplicateRootFolderPath)
    
    LoopFolders folder, True
  End If
  
End Sub

Sub LoopFolders(CurrentFolder As Outlook.MAPIFolder, ByVal Recursive As Boolean)

  Dim folder As Outlook.MAPIFolder

  If CurrentFolder.FolderPath = duplicateRootFolderPath Then
    Debug.Print "Skipped " & CurrentFolder.FolderPath
    Exit Sub
  End If
  
  DoFolderActions CurrentFolder

  For Each folder In CurrentFolder.Folders

    If Recursive Then
      LoopFolders folder, Recursive
    End If
  Next
End Sub

Private Sub DoFolderActions(folder As Outlook.MAPIFolder)

  Dim duplicateTargetFolderPath As String
  Dim duplicateTagertFolder As Outlook.folder
  
  duplicateTargetFolderPath = Replace(folder.FolderPath, rootFolderPath, duplicateRootFolderPath)
  Debug.Print "Deduplicating: " & folder.FolderPath
  CreateFolder (duplicateTargetFolderPath)
  Set duplicateTagertFolder = GetFolder(duplicateTargetFolderPath)
  RemoveDuplicateItems folder, duplicateTagertFolder
    
End Sub


Sub RemoveDuplicateItems(objFolder As Outlook.folder, objTargetFolder As Outlook.folder)
    Dim objDictionary As Object
    Dim i As Long
    Dim objItem As Object
    Dim strKey As String

    Set objDictionary = CreateObject("scripting.dictionary")

    If Not (objFolder Is Nothing) Then
        objFolder.Items.Sort "[CreationTime]", False
        For i = objFolder.Items.Count To 1 Step -1
           Set objItem = objFolder.Items.Item(i)
           strKey = ""
            
            Select Case True
               'Check email subject, body and sent time
               Case TypeOf objItem Is Outlook.MailItem
                 Dim currentMailItem As Outlook.MailItem
                 Set currentMailItem = objItem
                 strKey = "MailItem" & currentMailItem.Subject & "," & currentMailItem.Body & "," & currentMailItem.SentOn
               'Check appointment subject, start time, duration, location and body
               Case TypeOf objItem Is Outlook.MeetingItem
                strKey = "MeetingItem" & objItem.Subject & "," & objItem.Body & "," & objItem.SentOn
               Case TypeOf objItem Is Outlook.ReportItem
                strKey = "ReportItem" & objItem.Subject & "," & objItem.Body
               Case TypeOf objItem Is Outlook.AppointmentItem
                 strKey = "AppointmentItem" & objItem.Subject & "," & objItem.Start & "," & objItem.Duration & "," & objItem.Location & "," & objItem.Body
               'Check contact full name and email address
               Case TypeOf objItem Is Outlook.ContactItem
                 strKey = "ContactItem" & objItem.FullName & "," & objItem.Email1Address & "," & objItem.Email2Address & "," & objItem.Email3Address
               'Check task subject, start date, due date and body
               Case TypeOf objItem Is Outlook.TaskItem
                 strKey = "TaskItem" & objItem.Subject & "," & objItem.StartDate & "," & objItem.DueDate & "," & objItem.Body
            End Select
    
            If Not strKey = "" Then
              strKey = Replace(strKey, ", ", Chr(32))
    
              'Remove the duplicate items
              If objDictionary.Exists(strKey) = True Then
              objItem.Move objTargetFolder
              Else
                 objDictionary.Add strKey, True
              End If
            End If
       Next i
    End If
End Sub


Function GetFolder(ByVal FolderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
 
    On Error GoTo GetFolder_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    
    On Error GoTo 0
    
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set TestFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            Set TestFolder = SubFolders.Item(FoldersArray(i))
            If TestFolder Is Nothing Then
                Set GetFolder = Nothing
            End If
        Next
    End If
     
   'Return the TestFolder
    Set GetFolder = TestFolder
    Exit Function
 
GetFolder_Error:
    Set GetFolder = Nothing
    Exit Function
End Function


Function CreateFolder(ByVal FolderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
 
    On Error GoTo GetFolder_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set TestFolder = Application.Session.Folders.Item(FoldersArray(0))
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            Set TestFolder = Nothing
            
            On Error Resume Next
            Set TestFolder = SubFolders.Item(FoldersArray(i))
            If TestFolder Is Nothing Then
                SubFolders.Add (FoldersArray(i))
                Set TestFolder = SubFolders.Item(FoldersArray(i))
            End If
        Next
    End If
     
    Exit Function
 
GetFolder_Error:
    Exit Function
End Function
