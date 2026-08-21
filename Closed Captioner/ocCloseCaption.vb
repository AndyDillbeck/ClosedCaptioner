Imports System.ComponentModel
Imports System.Collections.ObjectModel
Public Class ocCloseCaption
    Implements INotifyPropertyChanged
    Private _Changed As Boolean
    Public Event PropertyChanged(sender As Object, e As System.ComponentModel.PropertyChangedEventArgs) Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

    Protected Overridable Sub OnPropertyChanged(ByVal Propertyname As String)
        On Error GoTo sError

        If Not Propertyname.Contains("Changed") Then
            Changed = True
        End If
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(Propertyname))

        Exit Sub
sError:
        MsgBox(ErrorToString)
    End Sub

    Public Property Changed() As Boolean
        Get
            Return _Changed
        End Get
        Set(ByVal value As Boolean)
            If _Changed <> value Then
                _Changed = value
                OnPropertyChanged("Changed")
            End If
        End Set
    End Property

    Private _startTime As String
    Private _endTime As String
    Private _captionText As String
    Private _linkUrl As String
    Private _linkText As String

    Public Sub New()
        _startTime = ""
        _endTime = ""
        _captionText = ""
        _linkUrl = ""
        _linkText = ""
    End Sub

    Public Sub New(startTime As String)
        _startTime = startTime
        _endTime = ""
        _captionText = ""
        _linkUrl = ""
        _linkText = ""
    End Sub

    Public Sub New(startTime As String, endTime As String)
        _startTime = startTime
        _endTime = endTime
        _captionText = ""
        _linkUrl = ""
        _linkText = ""
    End Sub

    Public Sub New(startTime As String, endTime As String, captionText As String)
        _startTime = startTime
        _endTime = endTime
        _captionText = captionText
        _linkUrl = ""
        _linkText = ""
    End Sub

    Public Sub New(startTime As String, endTime As String, captionText As String, linkUrl As String, linkText As String)
        _startTime = startTime
        _endTime = endTime
        _captionText = captionText
        _linkUrl = linkUrl
        _linkText = linkText
    End Sub

    Public Property StartTime As String
        Get
            Return _startTime '.Hours & ":" & _startTime.Minutes & ":" & _startTime.Seconds & "," & _startTime.Milliseconds
        End Get
        Set(value As String)
            If value <> _startTime Then '.Hours & ":" & _startTime.Minutes & ":" & _startTime.Seconds & "," & _startTime.Milliseconds Then
                _startTime = value
                OnPropertyChanged("StartTime")
            End If
        End Set
    End Property

    Public Property EndTime As String
        Get
            Return _endTime '.Hours & ":" & _endTime.Minutes & ":" & _endTime.Seconds & "," & _endTime.Milliseconds
        End Get
        Set(value As String)
            If value <> _endTime Then '.Hours & ":" & _endTime.Minutes & ":" & _endTime.Seconds & "," & _endTime.Milliseconds Then
                _endTime = value 'TimeSpan.Parse()
                OnPropertyChanged("EndTime")
            End If
        End Set
    End Property

    Public Property CaptionText As String
        Get
            Return _captionText
        End Get
        Set(value As String)
            If value <> _captionText Then
                _captionText = value
                OnPropertyChanged("CaptionText")
            End If
        End Set
    End Property

    Public ReadOnly Property Output As String
        Get
            Return _startTime & " --> " & _endTime & vbCrLf & _captionText
        End Get
    End Property

    Public Property linkUrl As String
        Get
            Return _linkUrl
        End Get
        Set(value As String)
            If value <> _linkUrl Then
                _linkUrl = value
                OnPropertyChanged("linkUrl")
            End If
        End Set
    End Property
    Public Property linkText As String
        Get
            Return _linkText
        End Get
        Set(value As String)
            If value <> _linkText Then
                _linkText = value
                OnPropertyChanged("linkText")
            End If
        End Set
    End Property

End Class
