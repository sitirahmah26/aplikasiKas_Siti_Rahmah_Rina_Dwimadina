Imports System.Data.Odbc
Public Class Form1
    Dim saldosekarang As Integer

    Sub TampilGrid()
        bukakoneksi()

        DA = New OdbcDataAdapter("select * From tbl_tugas", CONN)
        DS = New DataSet
        DA.Fill(DS, "tbl_tugas")
        DataGridView1.DataSource = DS.Tables("tbl_tugas")

        tutupkoneksi()
    End Sub

    Sub getSaldoSekarang()
        bukakoneksi()

        CMD = New OdbcCommand("select * from tbl_tugas order by ID_Transaksi desc limit 1", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            Label6.Text = "0"
        Else
            Label6.Text = RD.Item("Saldo_Sekarang")
            saldosekarang = RD.Item("Saldo_Sekarang")
        End If

        tutupkoneksi()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TampilGrid()
        Form_Load()
        MunculCombo()
        getSaldoSekarang()
    End Sub

    Private Sub Form_Load()
        For i = 1 To 31
            ComboBox1.Items.Add(Str(i))
        Next i

        For j = 1 To 12
            ComboBox2.Items.Add(Str(j))
        Next j

        For k = 2005 To 2020
            ComboBox3.Items.Add(Str(k))
        Next k
    End Sub

    Sub MunculCombo()
        ComboBox4.Items.Add("Masuk")
        ComboBox4.Items.Add("Keluar")
    End Sub

    Sub KosongkanData()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("Silahkan Isi Semua Form")
        Else
            bukakoneksi()
            Dim tanggal As String = ComboBox1.Text + " -" + ComboBox2.Text + " -" + ComboBox3.Text
            If ComboBox4.Text = "Masuk" Then
                saldosekarang = saldosekarang + TextBox4.Text
                Dim simpan As String = "insert into tbl_tugas values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & tanggal & "', 
                '" & ComboBox4.Text & "','" & TextBox4.Text & "','" & Label6.Text & "','" & TextBox5.Text & "')"

                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Input data berhasil")
                Label6.Text = saldosekarang

            ElseIf ComboBox4.Text = "Keluar" Then
                If saldosekarang < TextBox4.Text Then
                    MsgBox("Saldo Tidak Cukup!")
                Else
                    saldosekarang = saldosekarang - TextBox4.Text
                    Label6.Text = saldosekarang
                    Dim simpan As String = "insert into tbl_tugas values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & tanggal & "', 
                    '" & ComboBox4.Text & "','" & TextBox4.Text & "','" & Label6.Text & "','" & TextBox5.Text & "')"

                    CMD = New OdbcCommand(simpan, CONN)
                    CMD.ExecuteNonQuery()
                    MsgBox("Input data berhasil")

                End If
            End If

            TampilGrid()
            KosongkanData()
            tutupkoneksi()
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        TextBox1.MaxLength = 6
        If e.KeyChar = Chr(13) Then
            bukakoneksi()
            CMD = New OdbcCommand("Select * From tbl_tugas where ID_Transaksi='" & TextBox1.Text & "'", CONN)
            RD = CMD.ExecuteReader
            RD.Read()
            If Not RD.HasRows Then
                MsgBox("ID Transaksi Tidak Ada, Silahkan coba lagi!")
                TextBox1.Focus()
            Else
                TextBox2.Text = RD.Item("Nama")
                TextBox3.Text = RD.Item("Telepon")
                ComboBox1.Text = RD.Item("Tanggal")
                ComboBox2.Text = RD.Item("Tanggal")
                ComboBox3.Text = RD.Item("Tanggal")
                ComboBox4.Text = RD.Item("Jenis")
                TextBox4.Text = RD.Item("Jumlah")
                TextBox5.Text = RD.Item("Keterangan")
                Label6.Text = RD.Item("Saldo_Sekarang")
                TextBox2.Focus()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        bukakoneksi()
        Dim tanggal As String = ComboBox1.Text + "-" + ComboBox2.Text + "-" + ComboBox3.Text
        Dim edit As String = "update tbl_tugas set 
        Nama='" & TextBox2.Text & "',
        Telepon='" & TextBox3.Text & "',
        Tanggal='" & tanggal & "' ,
        Jenis='" & ComboBox4.Text & "',
        Jumlah='" & TextBox4.Text & "',
        Keterangan='" & TextBox5.Text & "' WHERE ID_Transaksi='" & TextBox1.Text & "'"

        CMD = New OdbcCommand(edit, CONN)
        CMD.ExecuteNonQuery()
        MsgBox("Data Berhasil diUpdate")
        TampilGrid()
        KosongkanData()
        tutupkoneksi()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Silahkan Pilih Data yang di hapus dengan Masukan ID Transaksi dan ENTER")
        Else
            If MessageBox.Show("Yakin akan dihapus..?", "", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                bukakoneksi()
                Dim hapus As String = "delete From tbl_tugas where ID_Transaksi='" & TextBox1.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                TampilGrid()
                KosongkanData()
                tutupkoneksi()
            End If
        End If
    End Sub
End Class