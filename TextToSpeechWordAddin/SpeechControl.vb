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

    'Retrieves all the installed voices
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
            MsgBox("Error: No voices installed!", 0, "Error Popup")
        Else
            pauseimg.Visible = False
            playimg.Visible = True
            stopimg.Visible = False
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
        ComboBox2.Items.Add("Document") 'combobox2 == "Speech Amount"
        ComboBox2.Items.Add("Page")
        ComboBox2.Items.Add("Paragraph")
        ComboBox2.Items.Add("Selection")
        ComboBox2.Items.Add("Sentence")
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
        Else
            'If no voice is selected, no action is taken
            If String.IsNullOrEmpty(nextVoice) = True Then Exit Sub
            'Select the specified voice
            mySynth.SelectVoice(nextVoice)
            'Get the instance of the active Microsoft Word 2007 document
            Dim document As Word.Document = Globals.ThisAddIn.Application.ActiveDocument
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
            'Let it speak! show pause/stop buttons
            index = 1 'index of current word about to be read, text starts at 1!
            lastIndex = 0
            If highlight Then
                rng.HighlightColorIndex = Word.WdColorIndex.wdYellow
            End If

            mySynth.SpeakAsync(readMe)
            'set booleans
            stopimg.Visible = True
            pauseimg.Visible = True
            playimg.Visible = False
            speedTrackBar.Enabled = False
            volumeTrackBar.Enabled = False
            useHighlight.Enabled = False
            isPasused = False
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
        End If
        mySynth.SpeakAsyncCancelAll() 'stop speaking
        'set all booleans, remove highlighting
        isPasused = False
        playimg.Visible = True
        pauseimg.Visible = False
        stopimg.Visible = False
        speedTrackBar.Enabled = True
        volumeTrackBar.Enabled = True
        useHighlight.Enabled = True
        If highlight Then
            rng.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight
        End If
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
        'set all booleans, remove highlighting
        playimg.Visible = True
        stopimg.Visible = False
        pauseimg.Visible = False
        speedTrackBar.Enabled = True
        volumeTrackBar.Enabled = True
        useHighlight.Enabled = True
        If highlight Then
            rng.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight
        End If
    End Sub

    'occurs when about to speak a word, used for highlighting
    Private Sub mySynth_SpeakProgress(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mySynth.SpeakProgress
        If highlight Then
            Dim wrdrng As Word.Range = rng.Words.Item(index)
            Dim wrdtxt As String = wrdrng.Text
            Dim strWords() As String = Splits(wrdtxt, ".!?,;:'()[]{}" & Chr(34)) 'removes desired characters
            'find next word, NOT punctuation, to highlight (punctuation part of range in "Words")
            While strWords(0).Length = 0
                index = index + 1 'get next word
                wrdrng = rng.Words.Item(index)
                wrdtxt = wrdrng.Text
                strWords = Splits(wrdtxt, ".!?,;:'()[]{}" & Chr(34)) 'parse next word
            End While

            wrdrng.HighlightColorIndex = Word.WdColorIndex.wdBrightGreen 'highlight word
            If lastIndex > 0 Then ' unhighlight word read previously
                rng.Words.Item(lastIndex).HighlightColorIndex = Word.WdColorIndex.wdYellow
            End If
            lastIndex = index 'update last index
            index = index + 1
        End If
    End Sub


    Public Function Splits(ByVal InputText As String, _
         ByVal Delimiter As String) As Object

        ' This function splits the sentence in InputText into
        ' words and returns a string array of the words. Each
        ' element of the array contains one word.

        ' This constant contains punctuation and characters
        ' that should be filtered from the input string.
        Const CHARS As String = ".!?,;:'()[]{}" & Chr(34)
        Dim strReplacedText As String
        Dim intIndex As Integer

        ' Replace tab characters with space characters.
        strReplacedText = Trim(Replace(InputText, _
             vbTab, " "))

        ' Filter all specified characters from the string.
        For intIndex = 1 To Len(CHARS)
            strReplacedText = Trim(Replace(strReplacedText, _
                Mid(CHARS, intIndex, 1), " "))
        Next intIndex

        ' Loop until all consecutive space characters are
        ' replaced by a single space character.
        Do While InStr(strReplacedText, "  ")
            strReplacedText = Replace(strReplacedText, _
                "  ", " ")
        Loop

        ' Split the sentence into an array of words and return
        ' the array. If a delimiter is specified, use it.
        'MsgBox "String:" & strReplacedText
        Splits = Split(strReplacedText, Delimiter)

    End Function
End Class

