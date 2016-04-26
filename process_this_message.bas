'<declarations>
Public save_folder As String
Public base_folder As String
Public msg_subj As String
Public job_number As String
Public wo_number As String
Public job_name As String
Public desktop_path As String
'</declarations>


Function GetCurrentItem() As Object
' Call this function to get the current item, works for calendar events and mail items.
    Dim objApp As Outlook.Application
            
    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
            Set objApp = Nothing
End Function


Sub process_this_message()
    'MsgBox "Started" '<---Uncomment for testing
    
    base_folder = "C:\Users\smyers\Google Drive\Lutron\Jobs" '<----------CHANGE THIS LINE!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'MsgBox "base_folder: " & base_folder '<---Uncomment for testing
  
    Set new_msg = GetCurrentItem()
    Dim msg_att As Outlook.Attachment
    Dim base_folder, save_folder, msg_subj As String
    
    msg_subj = new_msg.Subject
    'MsgBox "msg_subj: " & msg_subj '<---Uncomment for testing
   
    save_folder = msg_subj
    
     job_number = Mid(msg_subj, (InStr(msg_subj, ",") + 2), 7)
     wo_number = Mid(msg_subj, (InStr(msg_subj, "WO") + 3), 8)
     job_name = Left(msg_subj, (InStr(msg_subj, ",")) - 1)
     save_folder = "JN " & job_number & " WO# " & wo_number & " " & job_name
    'MsgBox "JN " & job_number & " WO# " & wo_number & " " & job_name
    
    save_folder = Replace(save_folder, "Lutron Service Confirmation: ", "") 'remove  Lutron Service Confirmation from SAVE_FOLDER
    save_folder = Replace(save_folder, ":", "") 'remove colons from SAVE_FOLDER
    save_folder = Replace(save_folder, ".", "") 'remove periods from SAVE_FOLDER
    save_folder = Replace(save_folder, "/", "") 'remove slashes from SAVE_FOLDER
    save_folder = Replace(save_folder, ";", "") 'remove semi-colons from save_folder
    save_folder = Replace(save_folder, ",", "-") 'remove semi-colons from save_folder
    

    On Error Resume Next
    MkDir (base_folder & "\" & save_folder) 'create the save directory
    'MsgBox "Folder Created: " & save_folder '<---Uncomment for testing
    On Error GoTo 0

    For Each msg_att In new_msg.Attachments 'do it for each attachment
      

        msg_att.SaveAsFile (base_folder & "\" & save_folder & "\" & msg_att.DisplayName) 'save attachment
        'MsgBox msg_att.DisplayName & " saved to: " & save_folder '<---Uncomment for testing
        Set msg_att = Nothing
    Next
    '<--Create SDrive Shortcut-->
    
    
    Set WshShell = CreateObject("WScript.Shell")
    Set oShellLink = WshShell.CreateShortcut(base_folder & "\" & save_folder & "\SDrive_" & job_number & ".lnk")
    'MsgBox oShellLink '<---Uncomment for testing
    oShellLink.TargetPath = "\\intra.lutron.com\dfs01\cb\jobs\DOM\Ltg_Projects\Projects" & Left(job_number, 4) & "\" & job_number
    'MsgBox "\\intra.lutron.com\dfs01\cb\jobs\DOM\Ltg_Projects\Projects" & Left(job_number, 4) & "\" & job_number '<---Uncomment for testing
    oShellLink.WindowStyle = 1
    'MsgBox oShellLink.WindowStyle '<---Uncomment for testing
    oShellLink.Description = job_number & " S Drive Shortcut"
    'MsgBox oShellLink.Description '<---Uncomment for testing
    oShellLink.WorkingDirectory = "\\intra.lutron.com\dfs01\cb\jobs\DOM\Ltg_Projects\Projects" & Left(job_number, 4)
    'MsgBox oShellLink.WorkingDirectory '<---Uncomment for testing
    oShellLink.Save
    '</--Create SDrive Shortcut-->
    
End Sub




