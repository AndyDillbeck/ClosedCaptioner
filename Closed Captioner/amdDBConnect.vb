Imports System
Imports System.ComponentModel
Imports System.Xml.Serialization
Imports System.Data.SqlClient

Module amdDBConnect
    Private sConnectionString As String = "Data Source=jlw-awsrds-sqldb-02.creqsvmtsjmc.us-east-2.rds.amazonaws.com,1433;User ID=webuser;Password=i5gMysxc8XXEYViZO&yOCBX;Connect Timeout=30;Encrypt=True;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False; Initial Catalog=JL"
    Public Function MyCn() As ADODB.Connection
        MyCn = New ADODB.Connection
        MyCn.Provider = "SQLOLEDB.1"
        MyCn.ConnectionString = sConnectionString
        MyCn.CommandTimeout = 60
        MyCn.Open()
    End Function
    Public Function amdDBLookup(ByVal Table As String, ByVal Field As String, ByVal Where As String) As Object
        Dim sSql As String
        Dim rSelect As New ADODB.Recordset

        sSql = "Select " & Trim(Field) & " FROM " & Trim(Table) & " WHERE " & Trim(Where)

        With rSelect
            .Open(sSql, MyCn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
            If Not .EOF Then
                amdDBLookup = rSelect.Fields(0).Value ' NZ(, "")
            Else
                amdDBLookup = DBNull.Value
            End If
            .Close()
        End With
    End Function


    Public Function NZ(ByVal varValue As Object, Optional ByVal vIfNull As Object = "") As Object
        On Error GoTo NZError

        If IsDBNull(varValue) Then
            NZ = vIfNull
        Else
            NZ = varValue
        End If
        Exit Function

NZError:
        'lblError.Text = ErrorToString()
        'lblError.Visible = True
    End Function

    Public Function EZ(ByVal varValue As Object, Optional ByVal vIfEmpty As Object = "")
        On Error GoTo EZError
        If Not varValue.ToString = "" And Not varValue.ToString.ToLower() = "&nbsp;" Then
            EZ = Trim(varValue)
        Else
            EZ = Trim(vIfEmpty)
        End If
        Exit Function

EZError:
        'lblError.Text = ErrorToString()
        'lblError.Visible = True
    End Function
End Module
