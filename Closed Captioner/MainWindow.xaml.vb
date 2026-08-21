Imports NAudio
Imports NAudio.Wave
Imports System.Windows.Threading
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Class MainWindow
    Private samplecount As Integer = 4000
    Private timer As DispatcherTimer
    Private isDragging As Boolean = False
    Private isTicking As Boolean = False
    Private isPlaying As Boolean = False
    Private lines As Dictionary(Of String, Line)
    Private lastline As String
    Private cCloseCaption As ObservableCollection(Of ocCloseCaption)
    Property TtoX As vcTimestampToX
    Private CaptionID As Integer = 0
    Private VideoTitle As String = ""

    Private Sub LoadCCLines()
        On Error GoTo sError

        Dim sSql As String = "Select * From Training.dbo.tblVideoClosedCaptionLines Where CaptionID = " & CaptionID & " --Order By StartTime"
        Dim rSelect As New ADODB.Recordset
        cCloseCaption.Clear()
        With rSelect
            .Open(sSql, MyCn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Do While Not .EOF
                cCloseCaption.Add(New ocCloseCaption(.Fields!StartTime.Value, .Fields!EndTime.Value, .Fields!CaptionText.Value, .Fields!linkUrl.Value, .Fields!linkText.Value))
                .MoveNext()
            Loop
            .Close()
        End With

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdLoad_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdLoad.Click
        On Error GoTo sError

        Dim openfile As New Microsoft.Win32.OpenFileDialog
        openfile.ShowDialog()
        If System.IO.File.Exists(openfile.FileName) Then
            VideoWindow.Source = New Uri(openfile.FileName)
            Me.Title = "Closed Captioner - " & openfile.SafeFileName
            VideoTitle = openfile.SafeFileName

            Dim seconds As Integer = 0
            lines = New Dictionary(Of String, Line)

            Using Reader As New AudioFileReader(openfile.FileName)
                Dim samples = Reader.Length / (Reader.WaveFormat.Channels * Reader.WaveFormat.BitsPerSample / 8)
                Dim f = 0.0F
                Dim max = 0.0F

                Dim batch As Integer = Math.Max(10, samples / samplecount)
                Dim mid = 100
                Dim yScale = 100

                Dim buffer(batch) As Single

                Dim read As Integer

                Dim xPos = 0
                read = Reader.Read(buffer, 0, batch)
                While read = batch
                    For n As Integer = 0 To read
                        max = Math.Max(Math.Abs(buffer(n)), max)
                    Next

                    Dim line As New Line
                    line.X1 = xPos
                    line.X2 = xPos
                    line.Y1 = mid + (max * yScale)
                    line.Y2 = mid - (max * yScale)
                    line.Tag = Reader.CurrentTime
                    'line.Name = Reader.CurrentTime.ToString

                    line.StrokeThickness = 1
                    line.Stroke = Brushes.DarkGray

                    AddHandler line.MouseDown, AddressOf Line_MouseDown

                    waveform.Children.Add(line)

                    If Not lines.ContainsKey(Reader.CurrentTime.Hours.ToString().PadLeft(2, "0") & Reader.CurrentTime.Minutes.ToString().PadLeft(2, "0") & Reader.CurrentTime.Seconds.ToString().PadLeft(2, "0") & Reader.CurrentTime.Milliseconds.ToString().PadLeft(3, "0").Substring(0, 1)) Then
                        lines.Add(Reader.CurrentTime.Hours.ToString().PadLeft(2, "0") & Reader.CurrentTime.Minutes.ToString().PadLeft(2, "0") & Reader.CurrentTime.Seconds.ToString().PadLeft(2, "0") & Reader.CurrentTime.Milliseconds.ToString().PadLeft(3, "0").Substring(0, 1), line)
                    End If

                    If Reader.CurrentTime.TotalSeconds > (seconds + 10) Then
                        seconds = Reader.CurrentTime.TotalSeconds
                        line = New Line
                        line.X1 = xPos
                        line.X2 = xPos
                        line.Y1 = mid + yScale
                        line.Y2 = mid - yScale

                        line.StrokeThickness = 1
                        line.Stroke = Brushes.Black
                        waveform.Children.Add(line)

                        Dim textblock As New TextBlock
                        textblock.Text = Reader.CurrentTime.Hours.ToString().PadLeft(2, "0") & ":" & Reader.CurrentTime.Minutes.ToString().PadLeft(2, "0") & ":" & Reader.CurrentTime.Seconds.ToString().PadLeft(2, "0") & "," & Reader.CurrentTime.Milliseconds.ToString().PadLeft(3, "0")
                        textblock.Foreground = Brushes.Black
                        Canvas.SetLeft(textblock, xPos)
                        Canvas.SetTop(textblock, yScale)
                        waveform.Children.Add(textblock)
                    End If

                    max = 0
                    xPos += 1
                    read = Reader.Read(buffer, 0, batch)
                End While
                waveform.Width = xPos
            End Using
            TtoX = New vcTimestampToX(lines)
            ButtonsEnabled(True)
        End If

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Sub CreateRectangle(Left As Double, Width As Double)
        On Error GoTo sError

        'System.Threading.Thread.Sleep(100)
        'Dim left, right As Double
        'For Each l As String In lines.Keys
        '    If l = sender.Tag.Hours.ToString().PadLeft(2, "0") & sender.Tag.Minutes.ToString().PadLeft(2, "0") & sender.Tag.Seconds.ToString().PadLeft(2, "0") & sender.Tag.Milliseconds.ToString().PadLeft(3, "0").Substring(0, 1) Then
        '        left = lines(l).X1
        '        Exit For
        '    End If
        'Next

        Dim rect As Rectangle
        'Right = lines("0001500").X1
        rect = New Rectangle
        rect.Stroke = Brushes.Yellow
        rect.Fill = Brushes.Yellow
        rect.Opacity = 0.2
        rect.Height = 200
        rect.Width = Width 'Right - Left
        Canvas.SetLeft(rect, Left)
        waveform.Children.Add(rect)

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Sub Line_MouseDown(sender As Object, e As MouseButtonEventArgs)
        On Error GoTo sError

        VideoWindow.Position = sender.Tag
        VideoWindow.Play()
        System.Threading.Thread.Sleep(100)
        VideoWindow.Pause()
        isPlaying = False

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdPlay_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdPlay.Click
        On Error GoTo sError

        If isPlaying Then
            VideoWindow.Pause()
            cmdPlay.Content = "Play"
            isPlaying = False
        Else
            VideoWindow.Play()
            cmdPlay.Content = "Pause"
            isPlaying = True
        End If

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdPause_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) 'Handles cmdPause.Click
        On Error GoTo sError

        VideoWindow.Pause()
        isPlaying = False

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdStop_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        On Error GoTo sError

        VideoWindow.Stop()
        isPlaying = False

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdBack_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdBack.Click
        On Error GoTo sError

        VideoWindow.Position = VideoWindow.Position - TimeSpan.FromSeconds(5)

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdForward_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdForward.Click
        On Error GoTo sError

        VideoWindow.Position = VideoWindow.Position + TimeSpan.FromSeconds(5)

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdIn_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdIn.Click
        On Error GoTo sError

        VideoWindow.Pause()
        isPlaying = False
        Dim starttime As String = VideoWindow.Position.Hours.ToString().PadLeft(2, "0") & ":" & VideoWindow.Position.Minutes.ToString().PadLeft(2, "0") & ":" & VideoWindow.Position.Seconds.ToString().PadLeft(2, "0") & "," & VideoWindow.Position.Milliseconds.ToString().PadLeft(3, "0")
        Dim oc As ocCloseCaption = New ocCloseCaption(starttime)
        cCloseCaption.Add(oc)
        dgCaptions.SelectedItem = oc

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdOut_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdOut.Click
        On Error GoTo sError

        VideoWindow.Pause()
        isPlaying = False
        Dim endtime As String = VideoWindow.Position.Hours.ToString().PadLeft(2, "0") & ":" & VideoWindow.Position.Minutes.ToString().PadLeft(2, "0") & ":" & VideoWindow.Position.Seconds.ToString().PadLeft(2, "0") & "," & VideoWindow.Position.Milliseconds.ToString().PadLeft(3, "0")
        'cCloseCaption.Add(New ocCloseCaption(starttime))
        If Not dgCaptions.SelectedItem Is Nothing Then
            dgCaptions.SelectedItem.EndTime = endtime
        Else
            'dgCaptions.Items
        End If

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdOpenCC_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdOpenCC.Click
        On Error GoTo sError

        Dim frm As New LoadCCList
        frm.ShowDialog()
        If Not frm.CaptionID = 0 Then
            CaptionID = frm.CaptionID
            LoadCCLines()
        End If

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub cmdExport_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdExport.Click
        Try
            Dim i As Integer = 1
            Dim output As String = ""
            'MsgBox(output)
            Dim sSql As String = "Select * From Training.dbo.tblVideoClosedCaptions Where CaptionID = " & CaptionID
            Dim rSelect As New ADODB.Recordset
            With rSelect
                .Open(sSql, MyCn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
                If .EOF Then
                    .AddNew()
                    .Fields!Title.Value = VideoTitle
                    .Update()
                    .Close()
                    sSql = "Select Top 1 * From Training.dbo.tblVideoClosedCaptions Order By CaptionID Desc"
                    .Open(sSql, MyCn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockOptimistic)
                    CaptionID = .Fields!CaptionID.Value
                End If

                MyCn.Execute("Delete from Training.dbo.tblVideoClosedCaptionLines Where CaptionID = " & CaptionID)
                For Each oc As ocCloseCaption In dgCaptions.ItemsSource
                    MyCn.Execute("Insert Into Training.dbo.tblVideoClosedCaptionLines (CaptionID, StartTime, EndTime, CaptionText, linkURL, linkText) Values (" & CaptionID & ",'" & oc.StartTime.Replace("'", "''") & "','" & oc.EndTime.Replace("'", "''") & "','" & oc.CaptionText.Replace("'", "''") & "','" & oc.linkUrl.Replace("'", "''") & "','" & oc.linkText.Replace("'", "''") & "')")
                    output &= i & vbCrLf & oc.Output & vbCrLf & vbCrLf
                    i += 1
                Next

                .Fields!CaptionText.Value = output
                .Fields!CCDate.Value = Now.ToShortDateString
                .Update()
                MsgBox("Export successful!")
                .Close()
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        On Error GoTo sError

        ButtonsEnabled(False)
        timer = New DispatcherTimer
        timer.Interval = TimeSpan.FromMilliseconds(200)
        AddHandler timer.Tick, AddressOf timer_Tick
        cCloseCaption = New ObservableCollection(Of ocCloseCaption)

        dgCaptions.ItemsSource = cCloseCaption
        'RectArea.ItemsSource = cCloseCaption

        Me.DataContext = Me

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub VideoWindow_MediaEnded(sender As Object, e As System.Windows.RoutedEventArgs) Handles VideoWindow.MediaEnded
        On Error GoTo sError

        isPlaying = False
        cmdPlay.Content = "Play"

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub VideoWindow_MediaOpened(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles VideoWindow.MediaOpened
        On Error GoTo sError

        If VideoWindow.NaturalDuration.HasTimeSpan Then
            Dim ts As TimeSpan = VideoWindow.NaturalDuration.TimeSpan
            SeekBar.Maximum = ts.TotalSeconds
            SeekBar.SmallChange = 1
            SeekBar.LargeChange = Math.Min(10, ts.TotalSeconds / 10)
        End If
        timer.Start()

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Sub timer_Tick(sender As Object, e As EventArgs)
        If Not isDragging Then
            Dim currentposition As String = VideoWindow.Position.Hours.ToString().PadLeft(2, "0") & VideoWindow.Position.Minutes.ToString().PadLeft(2, "0") & VideoWindow.Position.Seconds.ToString().PadLeft(2, "0") & VideoWindow.Position.Milliseconds.ToString().PadLeft(3, "0").Substring(0, 1)
            SeekBar.Value = VideoWindow.Position.TotalSeconds
            txtPosition.Text = VideoWindow.Position.Hours.ToString().PadLeft(2, "0") & ":" & VideoWindow.Position.Minutes.ToString().PadLeft(2, "0") & ":" & VideoWindow.Position.Seconds.ToString().PadLeft(2, "0") & "," & VideoWindow.Position.Milliseconds.ToString().PadLeft(3, "0")
            Try
                lines(lastline).Stroke = Brushes.DarkGray
            Catch ex As Exception
            End Try
            Try
                lines(currentposition).Stroke = Brushes.Red
            Catch ex As Exception
            End Try
            lastline = currentposition
        End If
    End Sub

    Private Sub ButtonsEnabled(value As Boolean)
        On Error GoTo sError

        SeekBar.IsEnabled = value
        'cmdStop.IsEnabled = value
        'cmdPause.IsEnabled = value
        cmdPlay.IsEnabled = value
        cmdForward.IsEnabled = value
        cmdBack.IsEnabled = value
        cmdIn.IsEnabled = value
        cmdOut.IsEnabled = value
        cmdOpenCC.IsEnabled = value
        cmdExport.IsEnabled = value

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub SeekBar_DragCompleted(sender As System.Object, e As System.Windows.Controls.Primitives.DragCompletedEventArgs)
        On Error GoTo sError

        On Error GoTo sError

        isDragging = False
        VideoWindow.Play()
        isPlaying = True
        VideoWindow.Position = TimeSpan.FromSeconds(SeekBar.Value)

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub SeekBar_DragStarted(sender As System.Object, e As System.Windows.Controls.Primitives.DragStartedEventArgs)
        isDragging = True
    End Sub

    Private Sub mnuDeleteLine_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles mnuDeleteLine.Click
        On Error GoTo sError

        Dim oc As ocCloseCaption = dgCaptions.SelectedItem
        'Dim ssql As String = "Delete From JLProcessing.dbo.SchedRecROLineInfo Where SchedRecLineID = " & dgCaptions.SelectedItem.SchedRecLineID
        'MyCn.Execute(ssql)
        cCloseCaption.Remove(oc)

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If MsgBox("Are you sure you want to exit?", MsgBoxStyle.YesNo, "Are you just clicking on stuff randomly check") = MsgBoxResult.No Then
            e.Cancel = True
        End If
    End Sub
End Class
