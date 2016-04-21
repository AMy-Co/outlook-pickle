'<declarations>
Public save_folder As String
Public base_folder As String
Public msg_subj As String
Public job_number As String
Public wo_number As String
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

Sub process_message(new_msg As Outlook.MailItem)

    'Set process_this_message = objFSO.OpenTextFile("C:\logs\process_this_message.log", For_Writing, True)
    MsgBox "Started" '<---Uncomment for testing
    Set new_msg = GetCurrentItem()
    Dim msg_att As Outlook.Attachment
    Dim base_folder, save_folder, msg_subj As String
    

    msg_subj = new_msg.Subject
    'MsgBox "msg_subj: " & msg_subj '<---Uncomment for testing
    base_folder = "C:\Users\smyers\Google Drive\Lutron\Jobs" 'Change this line, you could even change it to the S drive
    'MsgBox "base_folder: " & base_folder '<---Uncomment for testing
  

    save_folder = msg_subj
    save_folder = Replace(save_folder, ":", "") 'remove colons from SAVE_FOLDER
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
    job_number = Mid(msg_subj, (InStr(msg_subj, ",") + 2), 7)
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
    
    'Call CopyLSC
   
End Sub

Sub process_this_message()
'Set process_this_message = objFSO.OpenTextFile("C:\logs\process_this_message.log", For_Writing, True)
    'MsgBox "Started" '<---Uncomment for testing
    Set new_msg = GetCurrentItem()
    Dim msg_att As Outlook.Attachment
    Dim base_folder, save_folder, msg_subj As String
    

    msg_subj = new_msg.Subject
    'MsgBox "msg_subj: " & msg_subj '<---Uncomment for testing
    base_folder = "C:\Users\smyers\Google Drive\Lutron\Jobs" 'Change this line, you could even change it to the S drive
    'MsgBox "base_folder: " & base_folder '<---Uncomment for testing
  

    save_folder = msg_subj
    
    save_folder = Replace(save_folder, "Lutron Service Confirmation: ", "") 'remove  Lutron Service Confirmation from SAVE_FOLDER
    save_folder = Replace(save_folder, ":", "") 'remove colons from SAVE_FOLDER
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
    job_number = Mid(msg_subj, (InStr(msg_subj, ",") + 2), 7)
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
    
    'Call CopyLSC
    
End Sub

Sub CopyLSC()

' <----Date---->
Dim SO_date As Variant 'initialize variable
SO_date = Format((Year(Now() + 1) Mod 100), _
        "20##") & "-" & _
        Format((Month(Now() + 1) Mod 100), "0#") & "-" & _
        Format((Day(Now()) Mod 100), "0#")
' </----Date---->

    Dim LSC_SO_master As String
    Dim LSC_SO As String
    
    'MsgBox "Started" '<---Uncomment for testing
    Set new_msg = GetCurrentItem()
    msg_subj = new_msg.Subject
    msg_subj_len = Len(msg_subj)
    
    base_folder = "C:\Users\smyers\Google Drive\Lutron\Jobs"
    save_folder = Replace(msg_subj, ":", "") 'remove colons from subject

    job_number = Mid(msg_subj, (InStr(msg_subj, "JN") + 3), 7)
    wo_number = Mid(msg_subj, (InStr(msg_subj, "WO#") + 4), 8)
    job_desc = Mid(msg_subj, (InStr(msg_subj, "WO#") + 15), msg_subj_len)

    'MsgBox job_number & "_" & wo_number '<---Uncomment for testing
    
    LSC_SO_master = base_folder & "\" & "LSC_SO_master.pdf"
    LSC_SO = save_folder & "\" & job_number & "_" & job_desc & "_SO.pdf"
    LSC_SO = Replace(LSC_SO, ":", "") 'remove colons from LSC_SO

    'MsgBox LSC_SO_master '<---Uncomment for testing
    'MsgBox LSC_SO '<---Uncomment for testing

    FileCopy LSC_SO_master, base_folder & "\" & LSC_SO

    'Name OldFile As NewFile

    'MsgBox "Ended" '<---Uncomment for testing
    
End Sub


