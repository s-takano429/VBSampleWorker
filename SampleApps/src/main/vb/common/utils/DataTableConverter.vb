Module DataTableConverter

    Public Function dataTableConvertStringArray(ByVal dt As DataTable, ByVal colIdx As String, Optional ByVal addBlank As Boolean = False) As String()
        Dim rowIdx As Integer = 0
        Dim list As New List(Of String)

        If dt.Rows.Count > 0 Then
            If addBlank Then
                list.Add("")
            End If
            If colIdx Is Nothing Then
                For rowIdx = 0 To dt.Rows.Count - 1
                    If Not list.Contains(dt.Rows(rowIdx)(0).ToString) Then
                        list.Add(dt.Rows(rowIdx)(0).ToString)
                    End If
                Next
            Else
                For rowIdx = 0 To dt.Rows.Count - 1
                    If Not list.Contains(dt.Rows(rowIdx)(colIdx).ToString) Then
                        list.Add(dt.Rows(rowIdx)(colIdx).ToString)
                    End If
                Next
            End If

            dataTableConvertStringArray = list.ToArray
        Else
            dataTableConvertStringArray = New String() {""}
        End If
    End Function
End Module
