Imports System.Collections.ObjectModel 'Use for speech collections
Imports Word = Microsoft.Office.Interop.Word
Imports System.Speech.Synthesis 'speech tools import

Public Class SpeechControl
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Toolbar button names
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'playimg - play button
    'stopimg - stop button when TTS playing
    'stopOff - stop button when TTS not playing (greyed out STOP)
    'volumeTrackBar - volume adjustor
    'speedTrackBar - speed adjustor
    'useHighlight - Highlighting checkbox
    'ComboBox1 - drop-down box for Speech Voice
    'ComboBox2 - drop-down box for Speech Amount
    'SingleR - Radio buttor for Single reading
    'StepR - Radio button for Step reading
    'ContinuousR - Radio button for continuous reading
    'PrimaryBox - drop-down box for primary highlight color (words)
    'SecondaryBox - drop-down box for secondary highlight color (selection)
    'ReadmeButton - button to show readme file

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Variables
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private WithEvents mySynth As New SpeechSynthesizer 'Object that reads text. generates progress events
    Dim isPasused As Boolean 'true if paused
    Dim nextVoice As String 'holds what voice will be used to read text
    Dim highlight As Boolean 'true = use highlighting while reading text
    Dim index As Integer 'index into selection
    Dim rng As Word.Range 'range of selection
    Dim readMe As String 'text of selection
    Dim lastIndex As Integer 'last index of word read
    Dim isInt As Boolean 'used to indicate "word" will generate multiple SpeakProgress events
    Dim count As Integer 'used when isInt is true, how long to pause highlight on number
    Dim continuous As Boolean 'if true, read continuously
    Dim steps As Boolean 'if true, update cursor after reading
    Dim singles As Boolean 'if true, only read selection once
    Dim stopClick As Boolean 'if true, stop button generated SpeakCompleted event
    Dim document As Word.Document 'holds document to read
    Dim PrimaryHighlight As Word.WdColorIndex 'color for word highlight
    Dim SecondaryHighlight As Word.WdColorIndex 'color for selection highlight

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Startup Code
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Sub SpeechControl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'first code run by office (automatically) on toolbar startup
        'on toolbar load, populate voice menu, amount menu, set booleans to defaults
        isPasused = False
        highlight = False
        continuous = False
        steps = False
        singles = True
        isInt = False
        PrimaryHighlight = Word.WdColorIndex.wdGreen
        SecondaryHighlight = Word.WdColorIndex.wdYellow
        'Get the instance of the active Microsoft Word 2007 document
        document = Globals.ThisAddIn.Application.ActiveDocument
        count = 0
        ComboBox2.Items.Add("Document") 'combobox2 == "Speech Amount"
        ComboBox2.Items.Add("Page")
        ComboBox2.Items.Add("Paragraph")
        ComboBox2.Items.Add("Sentence")
        ComboBox2.Items.Add("Selection")
        ComboBox2.SelectedIndex = 0 'select first option on load (Document)

        PrimaryBox.Items.Add("Yellow") 'primary highlight
        PrimaryBox.Items.Add("Bright Green")
        PrimaryBox.Items.Add("Turquoise")
        PrimaryBox.Items.Add("Pink")
        PrimaryBox.Items.Add("Blue")
        PrimaryBox.Items.Add("Red")
        PrimaryBox.Items.Add("Dark Blue")
        PrimaryBox.Items.Add("Teal")
        PrimaryBox.Items.Add("Green")
        PrimaryBox.Items.Add("Violet")
        PrimaryBox.Items.Add("Dark Red")
        PrimaryBox.Items.Add("Dark Yellow")
        PrimaryBox.Items.Add("Gray")
        PrimaryBox.Items.Add("Black")
        PrimaryBox.Items.Add("None")
        PrimaryBox.SelectedIndex = 1 'Select bright green

        SecondaryBox.Items.Add("Yellow") 'secondary highlight
        SecondaryBox.Items.Add("Bright Green")
        SecondaryBox.Items.Add("Turquoise")
        SecondaryBox.Items.Add("Pink")
        SecondaryBox.Items.Add("Blue")
        SecondaryBox.Items.Add("Red")
        SecondaryBox.Items.Add("Dark Blue")
        SecondaryBox.Items.Add("Teal")
        SecondaryBox.Items.Add("Green")
        SecondaryBox.Items.Add("Violet")
        SecondaryBox.Items.Add("Dark Red")
        SecondaryBox.Items.Add("Dark Yellow")
        SecondaryBox.Items.Add("Gray")
        SecondaryBox.Items.Add("Black")
        SecondaryBox.Items.Add("None")
        SecondaryBox.SelectedIndex = 0 'select yellow

        GetInstalledVoices(mySynth) 'get voices installed for TTS
    End Sub


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
            errorLabel.Visible = True 'make error msg visible
            continuousR.Enabled = False
            stepR.Enabled = False
            singleR.Enabled = False
            PrimaryBox.Enabled = False
            SecondaryBox.Enabled = False
            MsgBox("Error: No voices installed!", 0, "Error Popup") 'create error popup
        Else 'all good, show play buttons
            pauseimg.Visible = False
            playimg.Visible = True
            stopimg.Visible = False 'stop is 2 buttons: one for when TTS is going
            stopOff.Visible = True 'one for when it's not going
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
            'display error if something goes wrong, should never occur, even if no voices installed
            MsgBox("Error: Could not populate voice menu!", 0, "Error Popup")
        End Try

    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Button Code
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
            PrimaryBox.Enabled = False
            SecondaryBox.Enabled = False
        Else
            highlight = True
            PrimaryBox.Enabled = True
            SecondaryBox.Enabled = True
        End If
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
                rng.HighlightColorIndex = SecondaryHighlight
            End If

            'set booleans as for TTS playing (disables most options)
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
            PrimaryBox.Enabled = False
            SecondaryBox.Enabled = False
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
        mySynth.SpeakAsyncCancelAll() 'stop speaking, will generate SpeakCompleted event!
    End Sub

    'used to enable continuous reading and step reading only for paragraphs and sentences
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim txt As String = ComboBox2.Text
        If txt.ToLower = "paragraph" Then 'cont, step reading OK
            continuousR.Enabled = True
            stepR.Enabled = True
        ElseIf txt.ToLower = "sentence" Then 'cont, step reading OK
            continuousR.Enabled = True
            stepR.Enabled = True
        Else
            continuousR.Enabled = False 'else, disable them, pick valid options for radio buttons (ex Document mode)
            stepR.Enabled = False
            If continuous Or steps Then
                continuous = False
                singleR.Checked = True
                singles = True
                steps = False
            End If
        End If
    End Sub

    'Sets boolean based on what radio button selected
    Private Sub continuousR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles continuousR.CheckedChanged
        If continuousR.Checked Then
            continuous = True
        Else
            continuous = False
        End If
    End Sub

    'Sets boolean based on what radio button selected
    Private Sub stepR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stepR.CheckedChanged
        If stepR.Checked Then
            steps = True
        Else
            steps = False
        End If
    End Sub

    'Sets boolean based on what radio button selected
    Private Sub singleR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles singleR.CheckedChanged
        If singleR.Checked Then
            singles = True
        Else
            singles = False
        End If
    End Sub

    'primary highlight changed, update boolean, see index order for Case statement in SpeechControl_Load Sub
    Private Sub PrimaryBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrimaryBox.SelectedIndexChanged
        Select Case PrimaryBox.SelectedIndex
            Case 0
                PrimaryHighlight = Word.WdColorIndex.wdYellow
            Case 1
                PrimaryHighlight = Word.WdColorIndex.wdBrightGreen
            Case 2
                PrimaryHighlight = Word.WdColorIndex.wdTurquoise
            Case 3
                PrimaryHighlight = Word.WdColorIndex.wdPink
            Case 4
                PrimaryHighlight = Word.WdColorIndex.wdBlue
            Case 5
                PrimaryHighlight = Word.WdColorIndex.wdRed
            Case 6
                PrimaryHighlight = Word.WdColorIndex.wdDarkBlue
            Case 7
                PrimaryHighlight = Word.WdColorIndex.wdTeal
            Case 8
                PrimaryHighlight = Word.WdColorIndex.wdGreen
            Case 9
                PrimaryHighlight = Word.WdColorIndex.wdViolet
            Case 10
                PrimaryHighlight = Word.WdColorIndex.wdDarkRed
            Case 11
                PrimaryHighlight = Word.WdColorIndex.wdDarkYellow
            Case 12
                PrimaryHighlight = Word.WdColorIndex.wdGray50
            Case 13
                PrimaryHighlight = Word.WdColorIndex.wdBlack
            Case Else
                PrimaryHighlight = SecondaryHighlight
        End Select
    End Sub

    'secondary highlight changed, update boolean
    Private Sub SecondaryBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SecondaryBox.SelectedIndexChanged
        Dim rstPrim As Boolean = False 'true if primary highlight must change
        If PrimaryBox.SelectedIndex = 14 Then 'Primary Highlight set to none, set Prim = Secondary once secondary changed
            rstPrim = True
        End If
        Select Case SecondaryBox.SelectedIndex
            Case 0
                SecondaryHighlight = Word.WdColorIndex.wdYellow
            Case 1
                SecondaryHighlight = Word.WdColorIndex.wdBrightGreen
            Case 2
                SecondaryHighlight = Word.WdColorIndex.wdTurquoise
            Case 3
                SecondaryHighlight = Word.WdColorIndex.wdPink
            Case 4
                SecondaryHighlight = Word.WdColorIndex.wdBlue
            Case 5
                SecondaryHighlight = Word.WdColorIndex.wdRed
            Case 6
                SecondaryHighlight = Word.WdColorIndex.wdDarkBlue
            Case 7
                SecondaryHighlight = Word.WdColorIndex.wdTeal
            Case 8
                SecondaryHighlight = Word.WdColorIndex.wdGreen
            Case 9
                SecondaryHighlight = Word.WdColorIndex.wdViolet
            Case 10
                SecondaryHighlight = Word.WdColorIndex.wdDarkRed
            Case 11
                SecondaryHighlight = Word.WdColorIndex.wdDarkYellow
            Case 12
                SecondaryHighlight = Word.WdColorIndex.wdGray50
            Case 13
                SecondaryHighlight = Word.WdColorIndex.wdBlack
            Case Else
                SecondaryHighlight = Word.WdColorIndex.wdNoHighlight
        End Select
        If rstPrim Then
            PrimaryHighlight = SecondaryHighlight
        End If
    End Sub

    'button creates new word document, populates with readme
    Private Sub ReadmeButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReadmeButton.Click
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oPara1 As Word.Paragraph, oPara2 As Word.Paragraph, oPara3 As Word.Paragraph
        Dim oRng As Word.Range, iRng As Word.Range, jRng As Word.Range

        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a heading paragraph at the beginning of the document.
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "Microsoft Word 2007/2010 Text-to-Speech (TTS) toolbar"
        oPara1.Range.Font.Bold = True
        oPara1.Range.Font.Size = 20
        oPara1.Format.SpaceAfter = 24    '24 pt spacing after paragraph.
        oPara1.Range.InsertParagraphAfter()

        'Insert a text paragraph.
        oPara2 = oDoc.Content.Paragraphs.Add
        oPara2.Range.Text = "How to Use"
        oPara2.Range.Font.Bold = True
        oPara2.Range.Font.Size = 16
        oPara2.Format.SpaceAfter = 6
        oPara2.Range.InsertParagraphAfter()

        'go to end of document, insert everything else manually
        oRng = oDoc.Bookmarks.Item("\endofdoc").Range
        oRng.ParagraphFormat.SpaceAfter = 6
        oRng.InsertAfter("Open up the document you wish to read with TTS.  ")
        oRng.InsertParagraphAfter()
        oRng.InsertAfter("Select your desired reading voice.")
        oRng.InsertParagraphAfter()
        oRng.InsertAfter("In 'Speech Amount' drop box select how much of the document you want read every time you press play:")
        oRng.InsertParagraphAfter()
        oRng.Font.Bold = False
        oRng.Font.Size = 12

        iRng = oDoc.Bookmarks.Item("\endofdoc").Range
        'create a list
        iRng.InsertAfter("Document – read entire document")
        iRng.InsertParagraphAfter()
        iRng.InsertAfter("Page – read current page")
        iRng.InsertParagraphAfter()
        iRng.InsertAfter("Paragraph – read paragraph cursor is on")
        iRng.InsertParagraphAfter()
        iRng.InsertAfter("Sentence – read sentence cursor is on")
        iRng.InsertParagraphAfter()
        iRng.InsertAfter("Selection – read text selection you highlighted")
        iRng.InsertParagraphAfter()
        iRng.Font.Bold = False
        iRng.Font.Size = 12
        iRng.ListFormat.ApplyBulletDefault() 'makes bullet list

        jRng = oDoc.Bookmarks.Item("\endofdoc").Range 'go to new end of file
        jRng.InsertAfter("Underneath that drop box, select radio button corresponding to what you want Word to do after it finishes reading the amount specified by 'Speech Amount'. Some options may not be available based on 'Speech Amount' selection:")
        jRng.InsertParagraphAfter()

        iRng = oDoc.Bookmarks.Item("\endofdoc").Range
        'create another list
        iRng.InsertAfter("Single – Do nothing after reading selection")
        iRng.InsertParagraphAfter()
        iRng.InsertAfter("Step – Advance cursor to next selection.  For example, if in 'Sentence' mode, the cursor will move to the next sentence after reading finishes so that pressing play again will read the next sentence aloud.")
        iRng.InsertParagraphAfter()
        iRng.InsertAfter("Continuous – Advances to next selection and reads it aloud.  For example, if in 'Sentence' mode, the cursor will move to the next sentence after reading finishes and that next sentence will be read aloud.  This mode will end up reading the entire document unless stopped manually.")
        iRng.InsertParagraphAfter()
        iRng.ListFormat.ApplyBulletDefault() 'makes bullet list

        oRng = oDoc.Bookmarks.Item("\endofdoc").Range 'go to new end of file
        oRng.InsertAfter("Check the 'Enable Highlighting' box if you want Word to highlight the word it is currently reading, as well as highlight the current 'Speech Amount' in a different color.  There are known bugs that can occur in this mode, see 'Known Issues' section below.  ")
        oRng.InsertParagraphAfter()
        oRng.InsertAfter("Select how fast you want the 'Reading Speed' to be as well as setting the 'Volume'. These settings modify your TTS settings globally, not just for Word!")
        oRng.InsertParagraphAfter()
        oRng.InsertAfter("Press green play button to read selection.  Play button will become a pause button when reading.")
        oRng.InsertParagraphAfter()
        oRng.InsertAfter("If using highlighting, select your desired highlight colors in the drop boxes.")
        oRng.InsertParagraphAfter()
        oRng.InsertAfter("NOTE: You cannot change any settings unless playback is stopped.  Use the stop button to cancel playback and allow settings to be changed.")
        oRng.InsertParagraphAfter()
        oRng.InsertParagraphAfter()

        oRng.InsertAfter("To uninstall the toolbar, go to 'Programs and Features' in the Windows Vista/7 Control Panel and uninstall 'Text-to-Speech Word Addin'.")

        oRng.InsertParagraphAfter()

        'must set fonts after each section due to resused vars!
        oRng.Font.Bold = False
        oRng.Font.Size = 12
        jRng.Font.Bold = False
        iRng.Font.Size = 12
        jRng.Font.Size = 12
        iRng.Font.Bold = False

        'Insert another text paragraph.
        oPara3 = oDoc.Content.Paragraphs.Add
        oPara3.Range.Text = "Known Issues"
        oPara3.Range.Font.Bold = True
        oPara3.Range.Font.Size = 16
        oPara3.Format.SpaceAfter = 6
        oPara3.Range.InsertParagraphAfter()

        iRng = oDoc.Bookmarks.Item("\endofdoc").Range
        iRng.InsertAfter("If a Word document is open and you open a new blank document by double-clicking on a Microsoft Word shortcut in Windows, not through Word itself, the toolbar will not appear. Open new blank documents via the Office Orb/File Menu as a workaround.")
        iRng.InsertParagraphAfter()
        iRng.InsertAfter("Numbers can cause highlighting to become un-synchronized with audio. It is primarily caused by large numbers and phone numbers. Highlighting fixes itself upon moving onto next selection (sentence, page, etc), use continuous sentence reading for fastest correction of this bug when reading large amounts of text.")
        iRng.InsertParagraphAfter()
        iRng.InsertAfter("Special characters and words with unusual pronunciations may also cause highlighting to become un-synchronized with audio. See above bullet for suggested workaround.")
        iRng.InsertParagraphAfter()
        iRng.Font.Size = 12
        iRng.Font.Bold = False
        iRng.ListFormat.ApplyBulletDefault() 'makes bullet list
        iRng.InsertParagraphAfter()

        'Insert another text paragraph.
        oPara3 = oDoc.Content.Paragraphs.Add
        oPara3.Range.Text = "All rights held by the McBurney Center at the University of Wisconsin - Madison. Coded by Jacob Friedman, University of Wisconsin - Madison. 2010."
        oPara3.Range.Font.Bold = True
        oPara3.Range.Font.Size = 14
        oPara3.Format.SpaceAfter = 6
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Primary TTS Code for mySynth events
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'event occurs when speech done
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
                If highlight Then
                    PrimaryBox.Enabled = True
                    SecondaryBox.Enabled = True
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
                rng.HighlightColorIndex = SecondaryHighlight
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
            If highlight Then
                PrimaryBox.Enabled = True
                SecondaryBox.Enabled = True
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
        If highlight Then 'do anything? only if supposed to highlight
            If count > 0 Then 'decriment count for numbers, see below
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
                    Do Until ((Char.IsPunctuation(wrdtxt, 0) = False) Or (wrdtxt.Last().Equals(s) = False And Test(wrdtxt) = False))
                        index = index + 1
                        wrdrng = rng.Words.Item(index) 'get word at our new index
                        wrdtxt = wrdrng.Text
                    Loop
                    'once out of loop wrdrng points to next "Word" that will be read by TTS
                    'need to check if "word" will be generate more than 1 SPEAKPROGRESS event, ex: "1234.56" = 7 events, 3 "words"
                    Dim value As Double

                    isInt = Double.TryParse(wrdtxt, value)
                    If isInt Then '"Word" is an integer, will have an event per nonzero digit/decimal
                        'many special cases like dates and phone numbers also exist

                        count = wrdtxt.Length() 'how many times to stay on this word
                        'cases to test when to shorten count for number (when a digit is 0), works up to 100million
                        If (value Mod 10 = 0) Then
                            count = count - 1
                        ElseIf (value < 10 And count > 0) Then 'shorten count if only 1 digit
                            count = count - 1
                        End If
                        If (value Mod 100 < 10 And value > 99) Then 'each If tests if a digit is a 0, shorten count if so
                            count = count - 1
                        End If
                        If (value Mod 1000 < 100 And value > 999) Then
                            count = count - 1
                        End If
                        If (value Mod 10000 < 1000 And value > 9999) Then
                            count = count - 1
                        End If
                        If (value Mod 100000 < 10000 And value > 99999) Then
                            count = count - 1
                        End If
                        If (value Mod 1000000 < 100000 And value > 999999) Then
                            count = count - 1
                        End If
                        If (value Mod 10000000 < 1000000 And value > 9999999) Then
                            count = count - 1
                        End If
                        If (value Mod 100000000 < 10000000 And value > 99999999) Then
                            count = count - 1
                        End If
                        If wrdtxt.Last().Equals(s) Then
                            count = count - 1
                        End If
                        'special case: if 4 digits, no comma, and > 1000 it will read as 2 two digit words
                        If (value > 1000 And value < 10000) Then
                            count = 2
                        End If
                    End If

                    If Carriage(wrdtxt) = False Then 'fixes case of Carriage return "word" causing highlight desync
                        wrdrng.HighlightColorIndex = PrimaryHighlight 'highlight word
                        If lastIndex > 0 Then ' unhighlight word read previously
                            rng.Words.Item(lastIndex).HighlightColorIndex = SecondaryHighlight
                        End If
                    Else 'carriage return found, highlight next word
                        If lastIndex > 0 Then ' unhighlight word read previously
                            rng.Words.Item(lastIndex).HighlightColorIndex = SecondaryHighlight
                        End If
                        index = index + 1 'skip carriage return
                        wrdrng = rng.Words.Item(index)
                        wrdrng.HighlightColorIndex = PrimaryHighlight 'highlight next word
                    End If

                Catch ex As Exception 'Index out of bounds due to unknown chars, just continue on, but highlighting will be incorrect
                End Try
                lastIndex = index 'update lastIndex for next SpeakProgress
                index = index + 1 'update index
            End If
        End If
    End Sub

    'tests if word is a carriage return or other new line punctuation. if so, skip it!
    Public Function Carriage(ByVal input As String) As Boolean
        Dim c As Char = input.Chars(0)
        If c.Equals(Chr(15)) Then
            Return True
        ElseIf c.Equals(Chr(12)) Then
            Return True
        ElseIf c.Equals(Chr(13)) Then
            Return (True)
        ElseIf c.Equals(Chr(11)) Then
            Return True
        Else
            Return False
        End If
    End Function

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

End Class