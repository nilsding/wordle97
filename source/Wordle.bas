Attribute VB_Name = "Wordle"
Global WordleWord As String

Private Function gameOver(correctGuess As Boolean)
    Dim res As Integer
    
    Dim Text As String
    If correctGuess Then
        Text = "Great job!!!" & vbNewLine & _
            "You won this WORDle!"
    Else
        Text = "Too many guesses - game over!" & vbNewLine & _
            "The word was: " & WordleWord
    End If
    
    res = MsgBox(Text & vbNewLine & _
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

Public Function TextContainsChar(Text As String, Char As String) As Boolean
    TextContainsChar = False
    
    For i = 1 To Len(Text)
        If Char = Mid(Text, i, 1) Then
            TextContainsChar = True
            Exit Function
        End If
    Next
End Function

Public Function TextCharCount(Text As String, Char As String) As Integer
    TextCharCount = 0
    
    For i = 1 To Len(Text)
        If Char = Mid(Text, i, 1) Then
            TextCharCount = TextCharCount + 1
        End If
    Next
End Function

Public Function TextWithoutFirstChar(Text As String, Char As String) As String
    Dim CharFound As Boolean
    Dim TestChar As String
    CharFound = False
    
    TextWithoutFirstChar = ""
    
    For i = 1 To Len(Text)
        TestChar = Mid(Text, i, 1)
        If Not CharFound And Char = TestChar Then
            CharFound = True
        Else
            TextWithoutFirstChar = TextWithoutFirstChar & TestChar
        End If
    Next
End Function

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
    
    If Not IsKnownWord(UCase(Left(paragraph.Text, 5))) Then
        MsgBox "Unknown word!  Try again."
        
        ClearParagraph paragraph
        If lastParagraph <> 1 Then NewLine
        Exit Sub
    End If
    
    ' mark correct letters as green and record incorrect guesses
    Dim IncorrectLetters As String
    IncorrectLetters = ""
    For i = 1 To paragraph.Characters.Count - 1 ' -1 because of the same reason as it's 6 above
        With paragraph.Characters(i)
            .Text = UCase(.Text)

            If .Text = Mid(WordleWord, i, 1) Then ' current character is in correct place
                .HighlightColorIndex = WdColorIndex.wdBrightGreen
            Else
                ' reset the colour, just in case
                .HighlightColorIndex = WdColorIndex.wdAuto

                ' remember the incorrect guess
                IncorrectLetters = IncorrectLetters & .Text
            End If
        End With
    Next
    
    ' mark the other letters
    For i = 1 To paragraph.Characters.Count - 1 ' -1 because of the same reason as it's 6 above
        With paragraph.Characters(i)
            .Text = UCase(.Text)
            guess = guess & .Text
            
            If .Text <> Mid(WordleWord, i, 1) Then ' current character is not correct
                ' reset the colour, just in case
                .HighlightColorIndex = WdColorIndex.wdAuto

                If TextContainsChar(WordleWord, .Text) Then ' this character exists in the word somewhere else
                    If TextCharCount(IncorrectLetters, .Text) - TextCharCount(guess, .Text) >= 0 Then ' and this character was not already guessed in this round
                        IncorrectLetters = TextWithoutFirstChar(IncorrectLetters, .Text)
                        .HighlightColorIndex = WdColorIndex.wdYellow
                    End If
                End If
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


