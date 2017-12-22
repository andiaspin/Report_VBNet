Imports System.Data.Odbc
Imports CrystalDecisions.CrystalReports.Engine
Public Class Laporan

    Private Sub Laporan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call koneksinya()
        cmd = New OdbcCommand("select distinct tanggal as tgl from laporan", conn)
        dr = cmd.ExecuteReader
        ComboBox1.Items.Add("Pilih")
        ComboBox1.SelectedIndex = 0
        While dr.Read()

            ComboBox1.Items.Add(dr.Item("tgl"))
        End While
        cmd = New OdbcCommand("select distinct month(tanggal) as bulan from laporan", conn)
        dr = cmd.ExecuteReader
        ComboBox2.Items.Add("Pilih")
        ComboBox2.SelectedIndex = 0
        While dr.Read()

            ComboBox2.Items.Add(dr.Item("bulan"))
        End While

        cmd = New OdbcCommand("select distinct year(tanggal) as tahun from laporan", conn)
        dr = cmd.ExecuteReader
        ComboBox3.Items.Add("Pilih")
        ComboBox3.SelectedIndex = 0
        While dr.Read()
            ComboBox3.Items.Add(dr.Item("tahun"))
        End While
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            CrystalReportViewer1.ReportSource = Nothing
            Exit Sub
        Else
            ComboBox2.SelectedIndex = 0
            ComboBox3.SelectedIndex = 0
            Dim ReportKu As New ReportDocument
            'menentukan lokasi report yang akan ditampilkan
            ReportKu.Load(Application.StartupPath & "\Laporan\Lap_Harian.rpt")
            'nilai parameter TanggalMulai di ambil dari inputan dtpTanggalMulai
            ReportKu.SetParameterValue("Harinya", ComboBox1.Text)
            'nilai parameter TanggalSelesai di ambil dari inputan dtpTanggalSelesai
            'tampilkan ke CrystalReportViewer1
            CrystalReportViewer1.ReportSource = ReportKu
            CrystalReportViewer1.Refresh()


        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 Then
            CrystalReportViewer1.ReportSource = Nothing
            Exit Sub
        Else
            ComboBox1.SelectedIndex = 0
            ComboBox3.SelectedIndex = 0
            Dim ReportKu As New ReportDocument
            'menentukan lokasi report yang akan ditampilkan
            ReportKu.Load(Application.StartupPath & "\Laporan\Lap_Bulanan.rpt")
            'nilai parameter TanggalMulai di ambil dari inputan dtpTanggalMulai
            ReportKu.SetParameterValue("bulannya", ComboBox2.Text)
            'nilai parameter TanggalSelesai di ambil dari inputan dtpTanggalSelesai
            'tampilkan ke CrystalReportViewer1
            CrystalReportViewer1.ReportSource = ReportKu
            CrystalReportViewer1.Refresh()


        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.SelectedIndex = 0 Then
            CrystalReportViewer1.ReportSource = Nothing
            Exit Sub
        Else
            ComboBox2.SelectedIndex = 0
            ComboBox1.SelectedIndex = 0
            Dim ReportKu As New ReportDocument
            'menentukan lokasi report yang akan ditampilkan
            ReportKu.Load(Application.StartupPath & "\Laporan\Lap_Tahunan.rpt")
            'nilai parameter TanggalMulai di ambil dari inputan dtpTanggalMulai
            ReportKu.SetParameterValue("tahunnya", ComboBox3.Text)
            'nilai parameter TanggalSelesai di ambil dari inputan dtpTanggalSelesai
            'tampilkan ke CrystalReportViewer1
            CrystalReportViewer1.ReportSource = ReportKu
            CrystalReportViewer1.Refresh()


        End If
    End Sub
End Class