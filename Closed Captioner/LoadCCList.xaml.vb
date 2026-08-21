Imports System.Collections.ObjectModel
Public Class LoadCCList
    Private cCCList As ObservableCollection(Of ocCCList)
    Public CaptionID As Integer = 0

    Sub loadList()
        Try
            Dim rSelect As New ADODB.Recordset
            Dim sSql As String = "Select * From Training.dbo.tblVideoClosedCaptions Order By Title"

            With rSelect
                .Open(sSql, MyCn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                Do While Not .EOF
                    Try
                        cCCList.Add(New ocCCList(.Fields!CaptionID.Value, NZ(.Fields!Title.Value, ""), NZ(.Fields!CCDate.Value, "")))
                    Catch ex As Exception
                        MsgBox("The video " & NZ(.Fields!Title.Value, "") & " could not be added to the list.")
                    End Try
                    .MoveNext()
                Loop
                .Close()
            End With
        Catch ex As Exception
            MsgBox("Well, that wasn't supposed to happen. You can retry, but it probably won't help. This probably needs to be debugged.")
        End Try
    End Sub

    Private Sub LoadCCList_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        cCCList = New ObservableCollection(Of ocCCList)
        dgCCList.ItemsSource = cCCList
        loadList()
    End Sub

    Private Sub cmdLoad_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdLoad.Click
        If Not dgCCList.SelectedItem Is Nothing Then
            CaptionID = dgCCList.SelectedItem.CaptionID
            Me.Close()
        End If
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cmdCancel.Click
        CaptionID = 0
        Me.Close()
    End Sub
End Class
