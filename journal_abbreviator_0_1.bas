Attribute VB_Name = "Abbreviator"
Sub JournalAbbreviator0_1()
    Dim findAndReplaceList As Variant
    Dim filePath As String
    Dim fileContent As String
    Dim fileNumber As Integer
    Dim lines() As String
    Dim i As Integer
    Dim parts() As String
    Dim j As Integer
    Dim temp As Variant
    Dim r As Range
    Dim arrayLength As Integer
    Dim originalText As String
    Dim replacementText As String
    Dim cleanedText As String
    
    ' Define the path to the text file
    filePath = "/Users/.../termlist.txt"
    
    ' Open the file and read its content
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileContent = Input$(LOF(fileNumber), fileNumber)
    Close #fileNumber

    ' Remove BOM if present (e.g., UTF-8 BOM)
    If Left(fileContent, 3) = ChrW(&HFEFF) Then
        fileContent = Mid(fileContent, 4)
    End If

    ' Replace different newline characters
    fileContent = Replace(fileContent, vbCrLf, vbLf)  ' Convert Windows line endings to Unix
    fileContent = Replace(fileContent, vbCr, vbLf)    ' Convert Mac line endings to Unix

    ' Replace Unicode placeholders with actual characters
    fileContent = Replace(fileContent, "\u2013", ChrW(&H2013)) ' Replace placeholder for en-dash
    fileContent = Replace(fileContent, "\u00E9", ChrW(&HE9))   ' Replace placeholder for Ž
    fileContent = Replace(fileContent, "\u00FC", ChrW(&HFC))   ' Replace placeholder for Ÿ

    ' Split the content into lines
    lines = Split(fileContent, vbLf)
    
    ' Initialize the findAndReplaceList array
    ReDim findAndReplaceList(LBound(lines) To UBound(lines))

    ' Process each line
    For i = LBound(lines) To UBound(lines)
        ' Split the line into parts using tab as the delimiter
        parts = Split(lines(i), vbTab)
        ' Add the parts to the array if it contains exactly 2 elements
        If UBound(parts) = 1 Then
            findAndReplaceList(i) = Array(parts(0), parts(1))
        End If
    Next i

    ' Sort findAndReplaceList by the length of the full title in descending order
    For i = LBound(findAndReplaceList) To UBound(findAndReplaceList) - 1
        For j = i + 1 To UBound(findAndReplaceList)
            ' Check if both elements are arrays and compare their lengths
            If IsArray(findAndReplaceList(i)) And IsArray(findAndReplaceList(j)) Then
                If Len(findAndReplaceList(i)(0)) < Len(findAndReplaceList(j)(0)) Then
                    ' Swap elements
                    temp = findAndReplaceList(i)
                    findAndReplaceList(i) = findAndReplaceList(j)
                    findAndReplaceList(j) = temp
                End If
            End If
        Next j
    Next i

    ' Calculate the length of findAndReplaceList array
    arrayLength = UBound(findAndReplaceList) - LBound(findAndReplaceList) + 1
    
    ' Confirm continuation and display array length
    If MsgBox("Continue with the find and replace operations? Array length: " & arrayLength, vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    ' Check if there is a selection
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select the text you want to process.", vbExclamation
        Exit Sub
    End If
    
    ' Perform find and replace operations with error handling
    Set r = Selection.Range
    
    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False
    
    ' Perform find and replace operations
    For j = LBound(findAndReplaceList) To UBound(findAndReplaceList)
        If IsArray(findAndReplaceList(j)) Then
            On Error Resume Next ' Ignore errors during find and replace
            With r.Find
                .Text = findAndReplaceList(j)(0)
                .Replacement.Text = findAndReplaceList(j)(1)
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = True
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            r.Find.Execute Replace:=wdReplaceAll
            On Error GoTo 0 ' Resume normal error handling
        End If
    Next j

    ' Clean up trailing dots
    r.Select ' Select the entire range for cleaning
    Application.ScreenUpdating = False
    
    ' Find and remove extra dots
    Do While r.Find.Execute(FindText:="..", Forward:=True, _
                            Wrap:=wdFindStop, Format:=False, _
                            MatchCase:=False, MatchWholeWord:=False, _
                            MatchWildcards:=False, MatchSoundsLike:=False, _
                            MatchAllWordForms:=False)
        r.Text = Replace(r.Text, "..", ".")
    Loop
    
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    
    ' Notify user that the process is complete
    MsgBox "The find and replace process is complete.", vbInformation
End Sub
