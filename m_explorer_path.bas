Attribute VB_Name = "m_explorer_path"
Option Explicit

'creator: Julian Jung (julian.jung@tws-partners.com)
'change-log:
    '2022-09-05: review


'Input:  path (e.g., ActiveWorkbook)
'Output: "explorer" path

'"SaveAs" works with sharepoint, but "SaveCopyAs" requires "explorer" paths.
'This procedure transforms Sharepoint paths from the personal and operative drives to "explorer" paths.
'directory structure
    '  i) computer               C:\Users\JungJ\
    ' ii) Sharepoint - personal  "SharePoint"/personal/julian_jung_tws-partners_com/Documents/Persoenliche_Ordner/
    'iii) Sharepoint - operative "SharePoint"/sites/operative/Documents/
    ' iv) Sharepoint - admin     "SharePoint"/sites/administrative/Documents/

Private Function Explorer_path(str_original_path As String) As String

    Const str_explorer As String = "explorer"
    Const str_operative As String = "operative"
    Const str_admin As String = "administrative"
    Const str_personal As String = "personal"

    Dim str_saving_system As String
    Dim str_file As String
    
    'determine where the link comes from
    str_saving_system = Saving_system(str_original_path, str_explorer, str_operative, str_admin, str_personal)
    
    'transformer for all path cases
    If str_saving_system = str_explorer Then
        'no actions needed
    ElseIf str_saving_system = str_operative Then
        str_file = Operative_transformer(str_original_path, str_explorer, str_operative, str_admin, str_personal)
    ElseIf str_saving_system = str_admin Then 'haven't got and not planning to bring admin-Sharepoint to my local computer
        MsgBox "This script does not work with files from the 'Administrative - Sharepoint'."
        End
    ElseIf str_saving_system = str_personal Then
        str_file = Personal_transformer(str_original_path, str_explorer, str_operative, str_admin, str_personal)
    End If
    
    Explorer_path = str_file

End Function


Private Function Saving_system(str_original_path As String, _
                               str_explorer As String, str_operative As String, _
                               str_admin As String, str_personal As String) As String
    'input:  original path
    'output: {explorer, operative, admin, personal}
    
    Dim str_system As String
    Dim int_slash As Integer
    Dim str_shrinking_path As String
    Dim str_path_block As String
    
    If VBA.InStr(str_original_path, ":") = 2 Then
        str_system = str_explorer
    Else
        str_shrinking_path = str_original_path
        str_path_block = ""
        While str_path_block <> str_operative _
          And str_path_block <> str_admin _
          And str_path_block <> str_personal
          
            'cut path down by slashs until file-system is found
            int_slash = VBA.InStr(str_shrinking_path, "/")
            str_path_block = VBA.Left(str_shrinking_path, int_slash - 1)
            str_shrinking_path = VBA.Mid(str_shrinking_path, int_slash + 1)
            
            'error catcher for cases that I haven't considered
            If int_slash = 0 Then
                MsgBox "Neither" & str_operative & "nor" & str_admin & "nor" & str_personal & "was found in the file's path."
                End
            End If
        Wend
        
        str_system = str_path_block
    End If
    
    Saving_system = str_system
    
End Function


Private Function Operative_transformer(str_original_path As String, _
                                       str_explorer As String, str_operative As String, _
                                       str_admin As String, str_personal As String) As String
                                       
    Const str_drive As String = "C:\Users"
    Const str_tws As String = "TWS Partners"
    
    Dim str_file As String
    Dim str_user_vba As String, str_user_path As String
    Dim int_slash As Integer
    Dim int_split As Integer
    Dim str_shrinking_path As String
    Dim str_path_block As String
    Dim i As Integer
    
    'drive prefix
    str_file = str_drive
   
    'user name
    str_user_vba = Application.UserName
    int_split = VBA.InStr(str_user_vba, ",")
    str_user_path = VBA.Left(str_user_vba, int_split - 1) _
                  & VBA.Mid(str_user_vba, int_split + 2, 1)
    str_file = str_file & "\" & str_user_path
    
    'tws
    str_file = str_file & "\" & str_tws
    
    'extract project
    str_shrinking_path = str_original_path
    str_path_block = ""
    While str_path_block <> str_operative
        int_slash = VBA.InStr(str_shrinking_path, "/")
        str_path_block = VBA.Left(str_shrinking_path, int_slash - 1)
        str_shrinking_path = VBA.Mid(str_shrinking_path, int_slash + 1)
    Wend
    
    For i = 1 To 2
        int_slash = VBA.InStr(str_shrinking_path, "/")
        str_path_block = VBA.Left(str_shrinking_path, int_slash - 1)
        str_shrinking_path = VBA.Mid(str_shrinking_path, int_slash + 1)
    Next i 'project is now isolated in str_path_block & the leftover path is captured in str_shrinking_path
    
    str_file = str_file & "\" & _
               Application.WorksheetFunction.Proper(str_operative) & " - " & _
               str_path_block
     
    'add path but replace slashes
    str_shrinking_path = VBA.Replace(str_shrinking_path, "/", "\")
    str_file = str_file & "\" & str_shrinking_path

    Operative_transformer = str_file

End Function


Private Function Personal_transformer(str_original_path As String, _
                                      str_explorer As String, str_operative As String, _
                                      str_admin As String, str_personal As String) As String

    Const str_drive As String = "C:\Users"
    Const str_one_drive As String = "OneDrive"
    Const str_tws As String = "TWS Partners"
    Const str_personal As String = "Persoenliche_Ordner"

    Dim str_file As String
    Dim str_user_vba As String, str_user_path As String
    Dim str_tail As String
    Dim int_split As Integer, int_len_path As Integer, int_len_personal As Integer
    
    'drive prefix
    str_file = str_drive
   
    'user name
    str_user_vba = Application.UserName
    int_split = VBA.InStr(str_user_vba, ",")
    str_user_path = VBA.Left(str_user_vba, int_split - 1) _
                  & VBA.Mid(str_user_vba, int_split + 2, 1)
    str_file = str_file & "\" & str_user_path
    
    'tws
    str_file = str_file & "\" & _
               str_one_drive & " - " & _
               str_tws
        
    'personal folder
    str_file = str_file & "\" & str_personal
    
    'shrink path down to tail & replace slashes
    int_split = VBA.InStr(str_original_path, str_personal)
    int_len_path = VBA.Len(str_original_path)
    int_len_personal = VBA.Len(str_personal)
    str_tail = VBA.Right(str_original_path, int_len_path - int_split - int_len_personal)
    str_tail = VBA.Replace(str_tail, "/", "\")
    str_file = str_file & "\" & str_tail
    
    Personal_transformer = str_file

End Function


