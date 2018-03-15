Imports System.IO
Imports System.IO.File
Public Class Form1
    Dim MyFiles() As String = IO.Directory.GetFiles("Favourites\")
    Dim playlist As Integer
    Dim duration As Boolean
    Dim play As Boolean
    Dim autoplay As Boolean
    Dim singlefile As String
    Dim audioformat As String
    Dim mute As Boolean = False
    Private Sub AxWindowsMediaPlayer1_Enter(sender As Object, e As EventArgs)

    End Sub
    Public Sub ChecAssocation()
        Try
            My.Computer.Registry.ClassesRoot.CreateSubKey(".MP3").SetValue("", "MinoxPlayer", Microsoft.Win32.RegistryValueKind.String)
            My.Computer.Registry.ClassesRoot.CreateSubKey("MinoxPlayer\shell\open\comand").SetValue("", Application.ExecutablePath & " ""%1"" ", Microsoft.Win32.RegistryValueKind.String)
            My.Computer.Registry.ClassesRoot.CreateSubKey("MinoxPlayerDefaultIcon").SetValue("", Application.StartupPath & "\MP3.ico")
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Dim strLoad2 As String
        ListBox2.Items.Clear()

        strLoad2 = Dir("Favourites\" & audioformat)


        Do While strLoad2 > vbNullString
            ListBox2.Items.Add(strLoad2)
            strLoad2 = Dir()
        Loop

        duration = False
        OpenFileDialog1.Filter = "Audio files|*.mp3;*.wav;*.mpeg;*.mpeg3;*.aiff;*.aif;"
        If FlatComboBox1.SelectedItem = "" Then
            FlatComboBox1.SelectedItem = "MP3 (*.mp3)"
            audioformat = "\*.mp3"
        Else

        End If

        If FlatComboBox1.SelectedItem = "MP3 (*.mp3)" Then
            audioformat = "\*.mp3"
        ElseIf FlatComboBox1.SelectedItem = "WAV (*.wav)" Then
            audioformat = "\*.wav"
        ElseIf FlatComboBox1.SelectedItem = "AIFF (*.aiff)" Then
            audioformat = "\*.aiff"
        ElseIf FlatComboBox1.SelectedItem = "MPEG (*.mpeg)" Then
            audioformat = "\*.wav"
        Else
            FlatComboBox1.SelectedItem = "MP3 (*.mp3)"
            audioformat = "\*.mp3"
        End If
        Timer1.Start()
        Timer2.Start()
        Dim strLoad As String
        ListBox1.Items.Clear()
        If My.Settings.folder_location = "" Then
            strLoad = Dir(System.Environment.GetFolderPath(Environment.SpecialFolder.MyMusic))
        Else
            strLoad = Dir(My.Settings.folder_location & audioformat)
        End If

        Do While strLoad > vbNullString
            ListBox1.Items.Add(strLoad)
            strLoad = Dir()
        Loop

        AxWindowsMediaPlayer1.settings.setMode("shuffle", True)
        TrackBar1.Minimum = 0
        Timer1.Interval = 100

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Me.ListBox1.SelectedIndex >= 0 Then
            playlist = 1
        Else
            playlist = 2
        End If
        If AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsStopped And autoplay = True Then
            duration = False
            If playlist = 1 Then
                Dim end2 = ListBox1.SelectedIndex + 1
                If end2 = ListBox1.Items.Count Then
                    ListBox1.SelectedIndex = 0
                Else
                    ListBox1.SelectedIndex = Me.ListBox1.SelectedIndex + 1
                    Dim curMusic As String = ListBox1.SelectedItem.ToString()
                    AxWindowsMediaPlayer1.URL = My.Settings.folder_location & curMusic
                    duration = True
                End If

            Else
                Dim end1 = ListBox2.SelectedIndex + 1
                If end1 = ListBox2.Items.Count Then
                    ListBox2.SelectedIndex = 0
                Else
                    ListBox2.SelectedIndex = Me.ListBox2.SelectedIndex + 1
                    Dim curMusic As String = ListBox2.SelectedItem.ToString()
                    AxWindowsMediaPlayer1.URL = My.Settings.folder_location & curMusic
                    duration = True
                End If
            End If
        End If

        If AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsPlaying Then
            play = True
        Else
            play = False

        End If
        If play = True Then
            Button1.Hide()
            Button2.Show()
        Else
            Button1.Show()
            Button2.Hide()
        End If

        If FlatCheckBox1.Checked = True Then
            autoplay = True
        Else
            autoplay = False
        End If


    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        FlatLabel3.Text = My.Settings.folder_location
        AxWindowsMediaPlayer1.settings.volume = TrackBar2.Value
        If duration = True Then
            FlatStatusBar1.Text = "Now playing:  " & AxWindowsMediaPlayer1.Ctlcontrols.currentItem.name
            FlatLabel2.Text = AxWindowsMediaPlayer1.Ctlcontrols.currentPositionString
            TrackBar1.Maximum = Int(AxWindowsMediaPlayer1.currentMedia.duration + 1)

            If AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsStopped Or AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsMediaEnded Or AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsUndefined Or AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsWaiting Then
                FlatLabel2.Text = "0:00"
                TrackBar1.Value = 0
            Else

            End If

            If AxWindowsMediaPlayer1.playState = WMPLib.WMPPlayState.wmppsPlaying Then
                Me.TrackBar1.Value = Me.AxWindowsMediaPlayer1.Ctlcontrols.currentPosition
                FlatLabel1.Text = AxWindowsMediaPlayer1.Ctlcontrols.currentItem.durationString
                FlatLabel19.Text = AxWindowsMediaPlayer1.Ctlcontrols.currentItem.durationString
                FlatLabel20.Text = AxWindowsMediaPlayer1.Ctlcontrols.currentItem.durationString
                Dim zArtist As String = AxWindowsMediaPlayer1.Ctlcontrols.currentItem.getItemInfo("Artist")
                Dim zTrackTitle As String = AxWindowsMediaPlayer1.Ctlcontrols.currentItem.getItemInfo("Title")
                Dim zAlbum As String = AxWindowsMediaPlayer1.Ctlcontrols.currentItem.getItemInfo("Album")
                FlatLabel16.Text = zTrackTitle
                FlatLabel17.Text = zArtist
                FlatLabel18.Text = zAlbum
                FlatLabel23.Text = zTrackTitle
                FlatLabel22.Text = zArtist
                FlatLabel21.Text = zAlbum
            Else

            End If
        End If

        If File.Exists("Favourites\" & ListBox1.SelectedItem) Then
            Button7.Enabled = False
        Else
            Button7.Enabled = True
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim curMusic As String = ListBox1.SelectedItem.ToString()
        AxWindowsMediaPlayer1.URL = My.Settings.folder_location & curMusic
        duration = True

    End Sub

    Private Sub TrackBar1_Scroll_1(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        AxWindowsMediaPlayer1.Ctlcontrols.pause()
        AxWindowsMediaPlayer1.Ctlcontrols.currentPosition = CDbl(TrackBar1.Value)
        System.Threading.Thread.Sleep(150)
        AxWindowsMediaPlayer1.Ctlcontrols.play()
    End Sub

    Private Sub AxWindowsMediaPlayer1_Enter_1(sender As Object, e As EventArgs) Handles AxWindowsMediaPlayer1.Enter

    End Sub

    Private Sub FormSkin1_Click(sender As Object, e As EventArgs) Handles FormSkin1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        AxWindowsMediaPlayer1.Ctlcontrols.play()
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs)

    End Sub

    Private Sub Panel4_Paint(sender As Object, e As PaintEventArgs) Handles Panel4.Paint

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        duration = False
        ListBox1.SelectedIndex = Me.ListBox1.SelectedIndex + 1
        Dim curMusic As String = ListBox1.SelectedItem.ToString()
        AxWindowsMediaPlayer1.URL = My.Settings.folder_location & curMusic
        duration = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        duration = False
        ListBox1.SelectedIndex = Me.ListBox1.SelectedIndex - 1
        Dim curMusic As String = ListBox1.SelectedItem.ToString()
        AxWindowsMediaPlayer1.URL = My.Settings.folder_location & curMusic
        duration = True
    End Sub

    Private Sub FlatComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim folderDlg As New FolderBrowserDialog

        folderDlg.ShowNewFolderButton = False

        If (folderDlg.ShowDialog() = DialogResult.OK) Then

            My.Settings.folder_location = folderDlg.SelectedPath & "\"
            My.Settings.Save()
            Dim strLoad As String
            ListBox1.Items.Clear()
            strLoad = Dir(My.Settings.folder_location & audioformat)
            Do While strLoad > vbNullString
                ListBox1.Items.Add(strLoad)
                strLoad = Dir()
            Loop
            Dim root As Environment.SpecialFolder = folderDlg.RootFolder

        End If
    End Sub

    Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs) Handles FolderBrowserDialog1.HelpRequest

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        singlefile = OpenFileDialog1.FileName
        AxWindowsMediaPlayer1.URL = singlefile
    End Sub

    Private Sub FlatComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles FlatComboBox1.SelectedIndexChanged
        If FlatComboBox1.SelectedItem = "MP3 (*.mp3)" Then
            audioformat = "\*.mp3"
        ElseIf FlatComboBox1.SelectedItem = "WAV (*.wav)" Then
            audioformat = "\*.wav"
        ElseIf FlatComboBox1.SelectedItem = "AIFF (*.aiff)" Then
            audioformat = "\*.aiff"
        Else
            FlatComboBox1.SelectedItem = "MP3 (*.mp3)"
            audioformat = "\*.mp3"
        End If
        Dim strLoad As String
        ListBox1.Items.Clear()
        strLoad = Dir(My.Settings.folder_location & audioformat)
        Do While strLoad > vbNullString
            ListBox1.Items.Add(strLoad)
            strLoad = Dir()
        Loop
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        AxWindowsMediaPlayer1.Ctlcontrols.pause()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        mute = False
        If mute = True Then
            PictureBox2.Show()
            TrackBar2.Value = 0
            PictureBox1.Hide()
        Else
            PictureBox2.Hide()
            TrackBar2.Value = 50
            PictureBox1.Show()
        End If
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        mute = True
        If mute = True Then
            PictureBox2.Show()
            TrackBar2.Value = 0
            PictureBox1.Hide()
        Else
            PictureBox2.Hide()
            TrackBar2.Value = 50
            PictureBox1.Show()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Do While AxWindowsMediaPlayer1.playState <> WMPLib.WMPPlayState.wmppsPlaying
            Application.DoEvents()
        Loop
        AxWindowsMediaPlayer1.fullScreen = True
    End Sub

    Private Declare Function record Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

    Private Sub FlatButton1_Click(sender As Object, e As EventArgs)
        record("open new Type waveaudio Alias recsound", "", 0, 0)
        record("record recsound", "", 0, 0)
    End Sub

    Private Sub FlatButton2_Click(sender As Object, e As EventArgs)
        record("save recsound C:\Users\piotr\Desktop\twdwdwd.wav", "", 0, 0)
        record("close recsound", "", 0, 0)
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs)
        record("open new Type waveaudio Alias recsound", "", 0, 0)
        record("record recsound", "", 0, 0)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        record("save recsound C:\Users\piotr\Desktop\twdwdwd.wav", "", 0, 0)
        record("close recsound", "", 0, 0)
    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Button7_Click_2(sender As Object, e As EventArgs) Handles Button7.Click
        File.Copy(My.Settings.folder_location & ListBox1.SelectedItem, "Favourites\" & ListBox1.SelectedItem)
        FlatAlertBox2.Show()
        Dim strLoad2 As String
        ListBox2.Items.Clear()
        If My.Settings.folder_location = "" Then
            strLoad2 = Dir(System.Environment.GetFolderPath(Environment.SpecialFolder.MyMusic))
        Else
            strLoad2 = Dir("Favourites\" & audioformat)
        End If

        Do While strLoad2 > vbNullString
            ListBox2.Items.Add(strLoad2)
            strLoad2 = Dir()
        Loop
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        Dim curMusic As String = ListBox2.SelectedItem.ToString()
        AxWindowsMediaPlayer1.URL = My.Settings.folder_location & curMusic
        duration = True
    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        AxWindowsMediaPlayer1.Ctlcontrols.stop()
        File.Delete("Favourites\" & ListBox2.SelectedItem)
        ListBox2.Items.Clear()
        Dim strLoad2 = Dir("Favourites\" & audioformat)


        Do While strLoad2 > vbNullString
            ListBox2.Items.Add(strLoad2)
            strLoad2 = Dir()
        Loop
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        AxWindowsMediaPlayer1.Ctlcontrols.stop()
        File.Delete(My.Settings.folder_location & ListBox1.SelectedItem)
        ListBox1.Items.Clear()

        Dim strLoad = Dir(My.Settings.folder_location & audioformat)


        Do While strLoad > vbNullString
            ListBox1.Items.Add(strLoad)
            strLoad = Dir()
        Loop
    End Sub
End Class
