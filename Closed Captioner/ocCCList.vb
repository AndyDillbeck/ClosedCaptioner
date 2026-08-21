Imports System.ComponentModel
Imports System.Collections.ObjectModel
Public Class ocCCList
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

    Private _captionid As Integer
    Private _title As String
    Private _ccdate As Date

    Public Sub New(CaptionID As Integer, Title As String, CCDate As Date)
        _captionid = CaptionID
        _title = Title
        _ccdate = CCDate
    End Sub

    Public Property CaptionID As Integer
        Get
            Return _captionid
        End Get
        Set(value As Integer)
            If _captionid <> value Then
                _captionid = value
                OnPropertyChanged("CaptionID")
            End If
        End Set
    End Property

    Public Property Title As String
        Get
            Return _title
        End Get
        Set(value As String)
            If value <> _title Then
                _title = value
                OnPropertyChanged("Title")
            End If
        End Set
    End Property

    Public Property CCDate As Date
        Get
            Return _ccdate
        End Get
        Set(value As Date)
            If value <> _ccdate Then
                _ccdate = value
                OnPropertyChanged("CCDate")
            End If
        End Set
    End Property

End Class
