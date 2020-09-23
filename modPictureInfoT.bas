Attribute VB_Name = "modPictureInfoGP"
'Option Compare Database
Option Explicit
Option Compare Text
' Public enumeration.
Public Enum opgParsePath
    FILE_ONLY
    PATH_ONLY
    DRIVE_ONLY
    FILEEXT_ONLY
End Enum
Public Function ParsePath(strPath As String, _
                          lngPart As opgParsePath) As String
    ' This procedure takes a file path and returns
    ' either the path, file, drive, or file extension portion,
    ' depending on which constant was passed in.
    
    Dim lngPos          As Long
    Dim strPart         As String
    Dim blnIncludesFile As Boolean
    
    ' Check that this is a file path.
    ' Find the last path separator.
    lngPos = InStrRev(strPath, "\")
    ' Determine whether portion of string after last backslash
    ' contains a period.
    blnIncludesFile = InStrRev(strPath, ".") > lngPos
    
    If lngPos > 0 Then
        Select Case lngPart
            ' Return file name.
            Case opgParsePath.FILE_ONLY
                If blnIncludesFile Then
                    strPart = right$(strPath, Len(strPath) - lngPos)
                Else
                    strPart = ""
                End If
            ' Return path.
            Case opgParsePath.PATH_ONLY
                If blnIncludesFile Then
                    'if you need with backslash "\"
                    '''strPart = left$(strPath, lngPos)
                    'or without
                    strPart = left$(strPath, lngPos - 1)
                Else
                    strPart = strPath
                End If
            ' Return drive.
            Case opgParsePath.DRIVE_ONLY
                strPart = left$(strPath, 3)
            ' Return file extension.
            Case opgParsePath.FILEEXT_ONLY
                If blnIncludesFile Then
                    ' Take three characters after period.
                    strPart = Mid(strPath, InStrRev(strPath, ".") + 1, 3)
                Else
                    strPart = ""
                End If
            Case Else
                strPart = ""
        End Select
    End If
    ParsePath = strPart

ParsePath_End:
    Exit Function

End Function


'Function FillDictionary(strPath As String) As Scripting.Dictionary
'    ' Looks at all files in folder and adds path and file name
'    ' of graphics files to a Dictionary object.
'
'    Dim fsoSysObj As Scripting.FileSystemObject
'    Dim fdrFolder As Scripting.Folder
'    Dim filFile As Scripting.File
'    Dim dctImages As Scripting.Dictionary
'
'    ' Get folder for passed-in path.
'    Set fsoSysObj = New FileSystemObject
'    Set fdrFolder = fsoSysObj.GetFolder(strPath)
'    ' Create a new Dictionary object.
'    Set dctImages = New Scripting.Dictionary
'
'    For Each filFile In fdrFolder.Files
'        ' Check file extension.
'        Select Case ParsePath(filFile.Path, FILEEXT_ONLY)
'            ' Add if file is bitmap, Windows metafile,
'            ' GIF, or JPEG.
'            Case "bmp", "wmf", "gif", "jpg"
'                dctImages.Add filFile.Path, filFile.Name
'        End Select
'    Next
'    ' Return filled Dictionary object.
'    Set FillDictionary = dctImages
'End Function


'errorhandler:
'    MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & _
'        "module: modPictureInfo,  Sub: SavePictureDataInfo"
'    Resume EXIT_SUB

