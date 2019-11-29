Imports System.Data.SqlClient
Imports System.IO
Imports System.Text.RegularExpressions

Module Module1

    Public dt As New DataTable
    Public dr As SqlDataReader
    Public ds As New DataSet
    Public adp As SqlDataAdapter
    Public cmd As SqlCommand
    Public con As New SqlConnection("")

    Public datagrid As DataGridView
    Public combo As ComboBox

    Public Sub koneksi()
        If ConnectionState.Closed Then
            con.Open()
        Else
            con.Close()
            con.Open()
        End If
    End Sub

    Public Sub tampil(ByVal query As String, ByVal mode As String, ByVal datagrid As DataGridView, ByVal combo As ComboBox, ByVal isi As String, ByVal tampil As String, Optional ByVal imgByte As Byte() = Nothing, Optional ByVal pictureBox As PictureBox = Nothing)
        koneksi()
        adp = New SqlDataAdapter(query, con)
        dt = New DataTable
        adp.Fill(dt)

        Select Case mode
            Case "datagrid"
                datagrid.DataSource = dt
            Case "combo"
                combo.DataSource = dt
                combo.ValueMember = isi
                combo.DisplayMember = tampil
            Case "gambar"
                Dim ms As MemoryStream = New MemoryStream(imgByte, False)
                Dim img As Image = Image.FromStream(ms)

                pictureBox.Image = img

                'Dim imgByte As Byte() = TryCast(Value, Byte())
        End Select
    End Sub

    Public Sub isi(ByVal query As String, ByVal mode As String, ByVal pesan As String, Optional ByVal field As String = "", Optional ByVal imgPath As String = "")
        koneksi()
        Select Case mode
            Case "simpan"
                cmd = New SqlCommand(query, con)
                cmd.ExecuteNonQuery()
                Select Case pesan
                    Case "simpan"
                    Case "ubah"
                End Select
            Case "simpan2"
                cmd = New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@" + field, File.ReadAllBytes(imgPath))
                cmd.ExecuteNonQuery()
                Select Case pesan
                    Case "simpan"
                    Case "ubah"
                End Select
            Case "hapus"
                Dim hapus As Integer = MsgBox("Apakah Anda yakin ingin menghapus data tersebut?", MsgBoxStyle.YesNo)
                If hapus = vbYes Then
                    cmd = New SqlCommand(query, con)
                    dr = cmd.ExecuteReader()
                    Select Case pesan
                        Case "hapus"
                    End Select
                End If
            Case "cari"
                cmd = New SqlCommand(query, con)
                dr = cmd.ExecuteReader()
                dr.Read()
        End Select
    End Sub

    Public Sub Msg(ByVal pesan As String, ByVal mode As String)
        Select Case mode
            Case "info"
                MessageBox.Show(pesan, "Informasi", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Case "error"
                MessageBox.Show(pesan, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select
    End Sub

    Public Function CheckPassword(ByVal password As String) As Boolean
        If Regex.IsMatch(password, "^.*(?=.{6,})(?=.*[a-z])(?=.*[A-Z])(?=.*\d).*$") Then
            Return True
        Else
            Msg("Password harus mengandung huruf besar, kecil dan karakter khusus", "error")
            Return False
        End If
    End Function

    Public Function GenerateID(ByVal query As String, ByVal prefix As String, ByVal field As String, ByVal pad As Integer) As String
        isi(query, "cari", "")

        If Not dr.HasRows Then
            Return prefix + "1".PadLeft(pad, "0")
        Else
            Dim OldID As String = dr.Item(field).ToString().Substring(2)
            Dim number As Integer = CInt(OldID) + 1
            Dim NewID As String = number.ToString().PadLeft(pad, "0")
            Return prefix + NewID
        End If

    End Function

    Public Sub DesignDatagrid(ByVal dg As DataGridView)
        dg.BackgroundColor = Color.White

        dg.ColumnHeadersDefaultCellStyle.BackColor = Color.LightSkyBlue
        dg.RowHeadersDefaultCellStyle.BackColor = Color.White
        dg.DefaultCellStyle.BackColor = Color.White
        dg.GridColor = Color.White

        dg.BorderStyle = BorderStyle.None
        dg.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        dg.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnMode.Fill
        dg.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        dg.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        dg.EnableHeadersVisualStyles = False
        dg.MultiSelect = False
    End Sub

End Module
