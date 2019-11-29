Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DesignDatagrid(DataGridView1)
    End Sub

    Private Sub LoadCombo()
        tampil("SELECT DATENAME(MONTH, date) As Month, MONTH(date) As MonthNumber FROM detailorder_views GROUP BY DATENAME(MONTH, date), MONTH(date) ORDER BY MONTH(date) asc", "combo", datagrid, combo, "", "")
    End Sub

    Public Sub LoadDatagrid()
        tampil("SELECT DATENAME(MONTH, date) As Month, SUM(price*qty) As Income FROM detailorder_veiw WHERE MONTH(date) between 1 and 2 GROUP BY DATENAME(MONTH, date), MONTH(date) ORDER BY MONTH(date) asc", "datagrid", datagrid, combo, "", "")
    End Sub

    Private Sub LoadChart()

        Dim rowCount As Integer = DataGridView1.Rows.Count()
        Dim row As Integer = DataGridView1.CurrentCell.RowIndex
        Chart1.Series(0).Points.Clear()

        For i = 0 To rowCount - 1
            Chart1.Series(0).Points.AddXY(DataGridView1.Item(0, i).Value, DataGridView1.Item(1, i).Value)
        Next

    End Sub

End Class
