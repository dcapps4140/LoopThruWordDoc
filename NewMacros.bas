Attribute VB_Name = "NewMacros"
Sub LoopingText()
'Will move to the end of each line in the document and move the text to match

'Declare variables
    Dim outputStr As String, currLine As String, endChar As String
    Dim numOfLines As Integer

'Count the number of non blank lines in current document
    numOfLines = ActiveDocument.BuiltInDocumentProperties("NUMBER OF LINES")

'Move to start of document
    Selection.HomeKey Unit:=wdStory

'Start the loop - looping once for each line
    For x1 = 1 To numOfLines
        'Move to start of line
        Selection.HomeKey Unit:=wdLine
        'Select entire line and copy into variable currLine
        Selection.EndKey Unit:=wdLine, Extend:=wdExtend
        currLine = Selection.Range.Text
        'Check to see if character currently at begining of currLine is "'" (apostrophe)
        endChar = Left(currLine, 1)
        If endChar = "'" Then
            'If preceding line included an apostrophe at the begining
            Selection.Range.Font.ColorIndex = wdDarkRed
            Selection.Range.Font.Bold = True
            'Debug.Print x1, endChar
        End If
        'Move down one line
        Selection.MoveDown Unit:=wdLine, Count:=1
    'Move to the next part of the loop
    Next x1
End Sub

