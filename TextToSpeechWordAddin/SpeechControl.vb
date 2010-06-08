Imports System.Speech.Synthesis 'speech tools import
Imports System.Collections.ObjectModel 'Use for speech collections

Public Class SpeechControl

    Private WithEvents mySynth As New SpeechSynthesizer 'reads text
    Dim isPasused As Boolean 'true if paused
    Dim nextVoice As String 'holds what voice will be used to read text
    Dim highlight As Boolean 'true = use highlighting while reading text
    Dim index As Integer 'index into selection
    Dim rng As Word.Range 'range of selection
    Dim readMe As String 'text of selection
    Dim lastIndex As Integer 'last index of word read
    Dim isInt As Boolean 'used to indicate "word" will generate multiple SpeakProgress events
    Dim count As Integer 'used when isInt is true
    Dim continuous As Boolean 'if true, read continuously
    Dim steps As Boolean 'if true, update cursor after reading
    Dim singles As Boolean 'if true, only read selection once
    Dim stopClick As Boolean 'if true, stop button generated SpeakCompleted event
    Dim document As Word.Document 'holds document to read

    'Retrieves all the installed voices, run on startup
    Private Sub GetInstalledVoices(ByVal synth As Speech.Synthesis.SpeechSynthesizer)
        'gets collection of InstalledVoice class objects, each InstalledVoice is a different 
        'speech voice
        Dim voices As ReadOnlyCollection(Of InstalledVoice) = _
          synth.GetInstalledVoices(Globalization.CultureInfo.CurrentCulture)
        If voices.Count = 0 Then
            'no voices installed, so disable controls, print error
            playimg.Enabled = False
            volumeTrackBar.Enabled = False
            speedTrackBar.Enabled = False
            useHighlight.Enabled = False
            stopimg.Enabled = False
            pauseimg.Enabled = False
            errorLabel.Visible = True
            continuousR.Enabled = False
            stepR.Enabled = False
            singleR.Enabled = False
            MsgBox("Error: No voices installed!", 0, "Error Popup")
        Else 'all good, show buttons
            pauseimg.Visible = False
            playimg.Visible = True
            stopimg.Visible = False
            stopOff.Visible = True
        End If

        'populate comboBox with voices
        Try
            Dim voiceInformation As VoiceInfo = voices(0).VoiceInfo
            For Each v As InstalledVoice In voices
                voiceInformation = v.VoiceInfo
                ComboBox1.Items.Add(voiceInformation.Name.ToString) 'combobox1 == "Speech voice"
            Next
            ComboBox1.SelectedIndex = 0 'select first option on load
        Catch ex As Exception
            'display error if something goes wrong
            MsgBox("Error: Could not populate voice menu!", 0, "Error Popup")
        End Try

    End Sub

    Private Sub SpeechControl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'on toolbar load, populate voice menu, amount menu
        isPasused = False
        highlight = False
        continuous = False
        steps = False
        singles = True
        isInt = False
        'Get the instance of the active Microsoft Word 2007 document
        document = Globals.ThisAddIn.Application.ActiveDocument
        count = 0
        ComboBox2.Items.Add("Document") 'combobox2 == "Speech Amount"
        ComboBox2.Items.Add("Page")
        ComboBox2.Items.Add("Paragraph")
        ComboBox2.Items.Add("Sentence")
        ComboBox2.Items.Add("Selection")
        ComboBox2.SelectedIndex = 0 'select first option on load
        GetInstalledVoices(mySynth)
    End Sub


    Private Sub playimg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles playimg.Click
        'if paused, resume it!
        If isPasused Then
            mySynth.Resume()
            isPasused = False
            playimg.Visible = False
            pauseimg.Visible = True
            stopimg.Visible = True
            stopOff.Visible = False
        Else
            'If no voice is selected, no action is taken
            If String.IsNullOrEmpty(nextVoice) = True Then Exit Sub
            'Select the specified voice
            mySynth.SelectVoice(nextVoice)
            document = Globals.ThisAddIn.Application.ActiveDocument
            'select reading amount
            Dim txt As String = ComboBox2.Text
            If txt.ToLower = "paragraph" Then
                readMe = Globals.ThisAddIn.Application.Selection.Paragraphs.First.Range.Text
                rng = Globals.ThisAddIn.Application.Selection.Paragraphs.First.Range
            ElseIf txt.ToLower = "selection" Then
                readMe = Globals.ThisAddIn.Application.Selection.Text
                rng = Globals.ThisAddIn.Application.Selection.Range
            ElseIf txt.ToLower = "sentence" Then
                readMe = Globals.ThisAddIn.Application.Selection.Sentences.First.Text
                rng = Globals.ThisAddIn.Application.Selection.Sentences.First
            ElseIf txt.ToLower = "page" Then
                rng = document.Bookmarks("\page").Range
                readMe = rng.Text
            Else 'read entire document
                readMe = document.Content.Text
                rng = document.Range
            End If
            'Let it speak! set buttons, booleans
            index = 1 'index of current word about to be read, text starts at 1!
            lastIndex = 0
            If highlight Then
                rng.HighlightColorIndex = Word.WdColorIndex.wdYellow
            End If

            'set booleans
            stopimg.Visible = True
            stopOff.Visible = False
            stopClick = False
            pauseimg.Visible = True
            playimg.Visible = False
            speedTrackBar.Enabled = False
            volumeTrackBar.Enabled = False
            useHighlight.Enabled = False
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            continuousR.Enabled = False
            singleR.Enabled = False
            stepR.Enabled = False
            isPasused = False
            mySynth.SpeakAsync(readMe) 'speak text

        End If
    End Sub

    Private Sub pauseimg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pauseimg.Click
        'pause!
        mySynth.Pause()
        'set booleans
        isPasused = True
        playimg.Visible = True
        pauseimg.Visible = False
        stopimg.Visible = True
    End Sub

    Private Sub stopimg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopimg.Click
        'stop speaking! set button visibilities/enables, remove highlight

        If isPasused Then 'if paused, resume first or it will get stuck
            mySynth.Resume()
            isPasused = False
        End If
        stopClick = True
        mySynth.SpeakAsyncCancelAll() 'stop speaking
        'set all booleans, remove highlighting

        'playimg.Visible = True
        'pauseimg.Visible = False
        'stopimg.Visible = False
        'speedTrackBar.Enabled = True
        'volumeTrackBar.Enabled = True
        'useHighlight.Enabled = True
        'If highlight Then
        '    rng.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight
        'End If

    End Sub

    Private Sub speedTrackBar_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles speedTrackBar.Scroll
        'Set speed
        mySynth.Rate = speedTrackBar.Value

    End Sub

    Private Sub volumeTrackBar_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles volumeTrackBar.Scroll
        'Set volume
        mySynth.Volume = volumeTrackBar.Value
    End Sub

    'set voice
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        nextVoice = ComboBox1.Text
    End Sub

    'toggles highlighting or not
    Private Sub useHighlight_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles useHighlight.CheckedChanged
        If highlight Then
            highlight = False
        Else
            highlight = True
        End If
    End Sub

    'occurs when speech done
    Private Sub mySynth_SpeakCompleted(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mySynth.SpeakCompleted

        'keep going: remove old highlight, find new range, highlight and speak it!
        If continuous And stopClick = False Then
            If highlight Then
                rng.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight
            End If
            Dim txt As String = ComboBox2.Text
            Try
                'try to move to next selection to read, update cursor
                If txt.ToLower = "paragraph" Then
                    Globals.ThisAddIn.Application.Selection.EndOf(Unit:=Word.WdUnits.wdParagraph) 'move to end of current paragraph
                    Globals.ThisAddIn.Application.Selection.Move(Unit:=Word.WdUnits.wdCharacter) 'plus 1 char
                    rng = rng.Next(Word.WdUnits.wdParagraph, 1) 'set range to next paragraph
                    readMe = rng.Text
                ElseIf txt.ToLower = "sentence" Then
                    Globals.ThisAddIn.Application.Selection.EndOf(Unit:=Word.WdUnits.wdSentence) 'move to end of current sentence
                    Globals.ThisAddIn.Application.Selection.Move(Unit:=Word.WdUnits.wdCharacter) 'plus 1 char
                    rng = rng.Next(Word.WdUnits.wdSentence, 1) 'set range to next sentance
                    readMe = rng.Text
                End If
            Catch ex As Exception 'no more text to read, return to default state
                playimg.Visible = True
                stopimg.Visible = False
                pauseimg.Visible = False
                speedTrackBar.Enabled = True
                volumeTrackBar.Enabled = True
                useHighlight.Enabled = True
                If txt.ToLower = "paragraph" Then
                    continuousR.Enabled = True
                    stepR.Enabled = True
                ElseIf txt.ToLower = "sentence" Then
                    continuousR.Enabled = True
                    stepR.Enabled = True
                Else
                    continuousR.Enabled = False
                    stepR.Enabled = False
                End If
                singleR.Enabled = True
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                stopClick = False
                stopOff.Visible = True
                Exit Sub
            End Try

            'Let it speak!
            index = 1 'index of current word about to be read, text starts at 1
            lastIndex = 0
            If highlight Then
                rng.HighlightColorIndex = Word.WdColorIndex.wdYellow
            End If
            mySynth.SpeakAsync(readMe) 'speak new text

        Else   'done speaking, set all booleans, remove highlighting
            Dim txt As String = ComboBox2.Text
            If txt.ToLower = "paragraph" Then
                continuousR.Enabled = True
                stepR.Enabled = True
            ElseIf txt.ToLower = "sentence" Then
                continuousR.Enabled = True
                stepR.Enabled = True
            Else
                continuousR.Enabled = False
                stepR.Enabled = False
            End If
            'check if cursor should move (Step selected)
            If steps And stopClick = False Then
                If txt.ToLower = "paragraph" Then
                    Globals.ThisAddIn.Application.Selection.EndOf(Unit:=Word.WdUnits.wdParagraph)
                    Globals.ThisAddIn.Application.Selection.Move(Unit:=Word.WdUnits.wdCharacter)
                ElseIf txt.ToLower = "sentence" Then
                    Globals.ThisAddIn.Application.Selection.EndOf(Unit:=Word.WdUnits.wdSentence)
                    Globals.ThisAddIn.Application.Selection.Move(Unit:=Word.WdUnits.wdCharacter)
                End If
            End If
            'turn on options again
            singleR.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            stopClick = False
            playimg.Visible = True
            stopimg.Visible = False
            pauseimg.Visible = False
            stopOff.Visible = True
            speedTrackBar.Enabled = True
            volumeTrackBar.Enabled = True
            useHighlight.Enabled = True
            If highlight Then
                rng.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight
            End If
        End If

    End Sub

    'occurs when about to speak a word, used for highlighting
    Private Sub mySynth_SpeakProgress(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mySynth.SpeakProgress
        If highlight Then
            If count > 0 Then
                count = count - 1
            Else
                isInt = False 'assume not a number
                Dim wrdrng As Word.Range
                Dim wrdtxt As String
                Try
                    wrdrng = rng.Words.Item(index)
                    wrdtxt = wrdrng.Text
                    Dim s As Char = " "
                    'loop until find next "Word" that will be read aloud by TTS.  Could be punctuation inside a word!
                    'if some types of punctuation, will read if last char in word is not a space (end of sentence)!
                    Do Until (Char.IsPunctuation(wrdtxt, 0) = False) Or (wrdtxt.Last().Equals(s) = False And Test(wrdtxt) = False)
                        index = index + 1
                        wrdrng = rng.Words.Item(index)
                        wrdtxt = wrdrng.Text
                    Loop
                    'once out of loop wrdrng points to next "Word" that will be read by TTS
                    'need to check if "word" will be generate more than 1 SPEAKPROGRESS event, ex: "1234.56" = 7 events, 3 "words"
                    Dim value As Double

                    isInt = Double.TryParse(wrdtxt, value)
                    If isInt Then '"Word" is an integer, will have an event per digit/decimal
                        count = wrdtxt.Length() - 1 'how many times to stay on this word
                        If wrdtxt.Last().Equals(s) Then
                            count = count - 1
                        End If
                    End If

                    wrdrng.HighlightColorIndex = Word.WdColorIndex.wdBrightGreen 'highlight word
                    If lastIndex > 0 Then ' unhighlight word read previously
                        rng.Words.Item(lastIndex).HighlightColorIndex = Word.WdColorIndex.wdYellow
                    End If

                Catch ex As Exception 'Index out of bounds due to unknown chars, just continue on, but highlighting will be incorrect
                End Try
                lastIndex = index 'update lastIndex for next SpeakProgress
                index = index + 1
            End If
        End If
    End Sub

    'characters that return True should not be highlighted as they will not be read aloud. Look up ASCII codes to see chars skipped.
    'there are probably many many more of these, I just did as many important ones I could find
    Public Function Test(ByVal input As String) As Boolean
        Dim c As Char = input.Chars(0)
        If c.Equals(Chr(133)) Then
            Return True
        ElseIf c.Equals(Chr(44)) Then
            Return True
        ElseIf c.Equals(Chr(34)) Then
            Return True
        ElseIf c.Equals(Chr(39)) Then
            Return True
        ElseIf c.Equals(Chr(40)) Then
            Return True
        ElseIf c.Equals(Chr(41)) Then
            Return True
        ElseIf c.Equals(Chr(45)) Then
            Return True
        ElseIf c.Equals(Chr(91)) Then
            Return True
        ElseIf c.Equals(Chr(93)) Then
            Return True
        ElseIf c.Equals(Chr(123)) Then
            Return True
        ElseIf c.Equals(Chr(125)) Then
            Return True
        ElseIf c.Equals(Chr(145)) Then
            Return True
        ElseIf c.Equals(Chr(146)) Then
            Return True
        ElseIf c.Equals(Chr(147)) Then
            Return True
        ElseIf c.Equals(Chr(96)) Then
            Return True
        ElseIf c.Equals(Chr(148)) Then
            Return True
        ElseIf c.Equals(Chr(150)) Then
            Return True
        ElseIf c.Equals(Chr(151)) Then
            Return True
        ElseIf c.Equals(Chr(191)) Then
            Return True
        ElseIf c.Equals(Chr(161)) Then
            Return True
        Else
            Return False
        End If
    End Function


    'used to enable continuous reading and step reading only for paragraphs and sentences
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim txt As String = ComboBox2.Text
        If txt.ToLower = "paragraph" Then
            continuousR.Enabled = True
            stepR.Enabled = True
        ElseIf txt.ToLower = "sentence" Then
            continuousR.Enabled = True
            stepR.Enabled = True
        Else
            continuousR.Enabled = False
            stepR.Enabled = False
            If continuous Or steps Then
                continuous = False
                singleR.Checked = True
                singles = True
                steps = False
            End If
        End If
    End Sub

    Private Sub continuousR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles continuousR.CheckedChanged
        If continuousR.Checked Then
            continuous = True
        Else
            continuous = False
        End If
    End Sub

    Private Sub stepR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stepR.CheckedChanged
        If stepR.Checked Then
            steps = True
        Else
            steps = False
        End If
    End Sub

    Private Sub singleR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles singleR.CheckedChanged
        If singleR.Checked Then
            singles = True
        Else
            singles = False
        End If
    End Sub

End Class

