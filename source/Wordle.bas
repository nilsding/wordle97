Attribute VB_Name = "Wordle"
Global WordleWord As String

Private Function gameOver(correctGuess As Boolean)
    Dim res As Integer
    
    Dim text As String
    If correctGuess Then
        text = "Great job!!!" & vbNewLine & _
            "You won this WORDle!"
    Else
        text = "Too many guesses - game over!" & vbNewLine & _
            "The word was: " & WordleWord
    End If
    
    res = MsgBox(text & vbNewLine & _
            "Play again?", vbYesNo, "Game Over")

    If res = vbYes Then
        ThisDocument.Content.Select
        Selection.Delete
        WordleWord = WordleWordList(Int(Rnd() * WordleWordListLength))
    End If
End Function

Private Sub ClearParagraph(paragraph)
    ' reset the colour of the characters, just in case
    For i = 1 To paragraph.Characters.Count
        paragraph.Characters(i).HighlightColorIndex = WdColorIndex.wdAuto
    Next
    
    ' delete the paragraph, including the newline
    paragraph.Select
    Selection.Delete
End Sub

Private Sub NewLine()
    ' insert a new line and jump to it
    ThisDocument.Content.InsertAfter vbNewLine
    Selection.EndKey wdStory
End Sub

Sub WordleGuess()
    Dim paragraph
    Dim lastParagraph As Long
    Dim guess As String
    lastParagraph = ThisDocument.Paragraphs.Count
    
    If WordleWord = "" Then
        ' initialise the game with the daily word
        WordleWord = GetDailyWord
    End If

    paragraph = ThisDocument.Paragraphs(lastParagraph)
    
    If paragraph.Characters.Count <> 6 Then ' it's 6 because of extra characters at the end?  bit strange this, but okay
        MsgBox "Your guess must be 5 characters long."
        
        ClearParagraph paragraph
        If lastParagraph <> 1 Then NewLine
        Exit Sub
    End If
    
    If Not IsKnownWord(UCase(Left(paragraph.text, 5))) Then
        MsgBox "Unknown word!  Try again."
        
        ClearParagraph paragraph
        If lastParagraph <> 1 Then NewLine
        Exit Sub
    End If
    
    For i = 1 To paragraph.Characters.Count - 1 ' -1 because of the same reason as it's 6 above
        With paragraph.Characters(i)
            .text = UCase(.text)
            guess = guess & .text
            
            If .text = Mid(WordleWord, i, 1) Then
                ' we have a match
                .HighlightColorIndex = WdColorIndex.wdBrightGreen
            Else
                ' reset the colour, just in case
                .HighlightColorIndex = WdColorIndex.wdAuto

                ' this character exists in the word somewhere else
                For j = 1 To Len(WordleWord)
                    If .text = Mid(WordleWord, j, 1) Then
                        .HighlightColorIndex = WdColorIndex.wdYellow
                    End If
                Next
            End If
        End With
    Next
    
    NewLine
    
    ' check if the guess was correct or we ran out of guesses
    If guess = WordleWord Then
        gameOver True
        Exit Sub
    ElseIf lastParagraph = 6 Then
        gameOver False
        Exit Sub
    End If
End Sub

