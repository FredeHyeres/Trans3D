Attribute VB_Name = "DiagAttachment"
Option Explicit

Sub DiagReferences()
    Dim oAttachments As Attachments
    Set oAttachments = ActiveModelReference.Attachments

    If oAttachments.Count = 0 Then
        MsgBox "0 references attachees.", vbInformation, "Diag"
        Exit Sub
    End If

    Dim sReport As String
    sReport = "=== DIAGNOSTIC ===" & vbCrLf
    sReport = sReport & "Nb refs : " & oAttachments.Count & vbCrLf & vbCrLf

    Dim i As Long
    For i = 1 To oAttachments.Count
        Dim oAtt As Attachment
        Set oAtt = oAttachments(i)

        sReport = sReport & "--- Ref #" & i & " ---" & vbCrLf
        sReport = sReport & "  TypeName = " & TypeName(oAtt) & vbCrLf

        On Error Resume Next

        sReport = sReport & "  .AttachName = " & oAtt.AttachName & vbCrLf

        Dim v As Variant

        Err.Clear
        v = oAtt.DisplayFlag
        If Err.Number = 0 Then
            sReport = sReport & "  .DisplayFlag = " & CStr(v) & vbCrLf
        Else
            sReport = sReport & "  .DisplayFlag = ERR " & Err.Number & vbCrLf
            Err.Clear
        End If

        Err.Clear
        v = oAtt.IsAttachment
        If Err.Number = 0 Then
            sReport = sReport & "  .IsAttachment = " & CStr(v) & vbCrLf
        Else
            sReport = sReport & "  .IsAttachment = ERR " & Err.Number & vbCrLf
            Err.Clear
        End If

        Err.Clear
        v = oAtt.DisplayPriority
        If Err.Number = 0 Then
            sReport = sReport & "  .DisplayPriority = " & CStr(v) & vbCrLf
        Else
            sReport = sReport & "  .DisplayPriority = ERR " & Err.Number & vbCrLf
            Err.Clear
        End If

        Err.Clear
        v = oAtt.IsActive
        If Err.Number = 0 Then
            sReport = sReport & "  .IsActive = " & CStr(v) & vbCrLf
        Else
            sReport = sReport & "  .IsActive = ERR " & Err.Number & vbCrLf
            Err.Clear
        End If

        Err.Clear
        v = oAtt.DisplayAsNested
        If Err.Number = 0 Then
            sReport = sReport & "  .DisplayAsNested = " & CStr(v) & vbCrLf
        Else
            sReport = sReport & "  .DisplayAsNested = ERR " & Err.Number & vbCrLf
            Err.Clear
        End If

        On Error GoTo 0
        sReport = sReport & vbCrLf
    Next i

    MsgBox sReport, vbInformation, "Diag Attachments"

    ' Ecrire dans un fichier texte a cote du DGN
    Dim sPath As String
    sPath = Left$(ActiveDesignFile.FullName, InStrRev(ActiveDesignFile.FullName, "\"))
    sPath = sPath & "DiagAttachment.txt"

    Dim nFile As Integer
    nFile = FreeFile
    Open sPath For Output As #nFile
    Print #nFile, sReport
    Close #nFile

    MsgBox "Diagnostic ecrit dans : " & sPath, vbInformation, "Diag"
End Sub
