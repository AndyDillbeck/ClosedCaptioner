Public Class vcTimestampToX
    Implements IValueConverter

    Private _lines As Dictionary(Of String, Line)

    Public Sub New(lines As Dictionary(Of String, Line))
        _lines = lines
    End Sub

    Public Function Convert(value As Object, targetType As System.Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements System.Windows.Data.IValueConverter.Convert
        Dim x As Double
        x = _lines(value).X1
        Return x
    End Function

    Public Function ConvertBack(value As Object, targetType As System.Type, parameter As Object, culture As System.Globalization.CultureInfo) As Object Implements System.Windows.Data.IValueConverter.ConvertBack
        Return Nothing
    End Function
End Class
