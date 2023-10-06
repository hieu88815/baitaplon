using BUS;
using DTO;
using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;


namespace baitaplon
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public string Ngay(string n)
        {
            string[] ans = n.Split(' ');
            return ans[0];
        }
        PhongDangThue_BUS phongdangthue = new PhongDangThue_BUS();
        KhachHang_BUS khachhang = new KhachHang_BUS();
        Phong_BUS phong = new Phong_BUS();
        GiaPhong_BUS giaphong = new GiaPhong_BUS();
        NhanVien_BUS nhanvien = new NhanVien_BUS();
        ChucVu_BUS chucvu = new ChucVu_BUS();
        SqlConnection conn = new SqlConnection("Data Source=HIEU;Initial Catalog=QuanLyKhachSan;Integrated Security=True");
        public DataTable Lay_DL(string sql)
        {
            SqlDataAdapter ad = new SqlDataAdapter(sql, conn);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            return dt;
        }
        private void tabPage3_Click(object sender, EventArgs e)
        {
            data_phongdangthue.DataSource = phongdangthue.Load_PhongDangThue();
        }

        private void insert_Click_1(object sender, EventArgs e)
        {
            PhongDangThue_DTO phongDTO = new PhongDangThue_DTO(ID_phongdangthue.Text, maphong_phongdangthue.Text, ngaynhanphong_phongdangthue.Text, ngaytraphong_phongdangthue.Text);
            phongdangthue.Insert_phongdangthue(phongDTO);
            tabPage3_Click(sender, e);
            ID_phongdangthue.Clear();
            maphong_phongdangthue.Clear();
            ngaynhanphong_phongdangthue.Clear();
            ngaytraphong_phongdangthue.Clear();
        }

        private void update_Click_1(object sender, EventArgs e)
        {
            PhongDangThue_DTO phongDTO = new PhongDangThue_DTO(ID_phongdangthue.Text, maphong_phongdangthue.Text, ngaynhanphong_phongdangthue.Text, ngaytraphong_phongdangthue.Text);
            phongdangthue.Update_phongdangthue(phongDTO);
            tabPage3_Click(sender, e);
            ID_phongdangthue.Clear();
            maphong_phongdangthue.Clear();
            ngaynhanphong_phongdangthue.Clear();
            ngaytraphong_phongdangthue.Clear();
        }

        private void delete_Click_1(object sender, EventArgs e)
        {
            phongdangthue.Delete_phongdangthue(ID_phongdangthue.Text, maphong_phongdangthue.Text);
            tabPage3_Click(sender, e);
            ID_phongdangthue.Clear();
            maphong_phongdangthue.Clear();
            ngaynhanphong_phongdangthue.Clear();
            ngaytraphong_phongdangthue.Clear();
        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            ID_phongdangthue.Text = data_phongdangthue.Rows[i].Cells[0].Value.ToString();
            maphong_phongdangthue.Text = data_phongdangthue.Rows[i].Cells[1].Value.ToString();
            ngaynhanphong_phongdangthue.Text = Ngay(data_phongdangthue.Rows[i].Cells[2].Value.ToString());
            ngaytraphong_phongdangthue.Text = Ngay(data_phongdangthue.Rows[i].Cells[3].Value.ToString());
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {
            data_khachhang.DataSource = khachhang.Load_KhachHang();
        }

        private void insert_khachhang_Click(object sender, EventArgs e)
        {
            KhachHang_DTO ob = new KhachHang_DTO(ID_khachhang.Text, name_khachhang.Text, date_khachhang.Text, number_khachhang.Text);
            khachhang.Insert_KhachHang(ob);
            tabPage4_Click(sender, e);
            ID_khachhang.Clear();
            name_khachhang.Clear();
            date_khachhang.Clear();
            number_khachhang.Clear();
        }

        private void update_khachhang_Click(object sender, EventArgs e)
        {
            KhachHang_DTO ob = new KhachHang_DTO(ID_khachhang.Text, name_khachhang.Text, date_khachhang.Text, number_khachhang.Text);
            khachhang.Update_KhachHang(ob);
            tabPage4_Click(sender, e);
            ID_khachhang.Clear();
            name_khachhang.Clear();
            date_khachhang.Clear();
            number_khachhang.Clear();
        }

        private void delete_khachhang_Click(object sender, EventArgs e)
        {
            khachhang.Delete_khachhang(ID_khachhang.Text);
            tabPage4_Click(sender, e);
            ID_khachhang.Clear();
            name_khachhang.Clear();
            date_khachhang.Clear();
            number_khachhang.Clear();
        }

        private void data_khachhang_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            ID_khachhang.Text = data_khachhang.Rows[i].Cells[0].Value.ToString();
            name_khachhang.Text = data_khachhang.Rows[i].Cells[1].Value.ToString();
            date_khachhang.Text = Ngay(data_khachhang.Rows[i].Cells[2].Value.ToString());
            number_khachhang.Text = data_khachhang.Rows[i].Cells[3].Value.ToString();
        }

        private void insert_phong_Click(object sender, EventArgs e)
        {
            Phong_DTO phongDTO = new Phong_DTO(maphong_phong.Text, tenphong_phong.Text, loaiphong_phong.Text);
            phong.Insert_Phong(phongDTO);
            tabPage5_Click(sender, e);
            maphong_phong.Clear();
            tenphong_phong.Clear();
            loaiphong_phong.Clear();
        }

        private void tabPage5_Click(object sender, EventArgs e)
        {
            data_phong.DataSource = phong.Load_Phong();
        }

        private void update_phong_Click(object sender, EventArgs e)
        {
            Phong_DTO phongDTO = new Phong_DTO(maphong_phong.Text, tenphong_phong.Text, loaiphong_phong.Text);
            phong.Update_Phong(phongDTO);
            tabPage5_Click(sender, e);
            maphong_phong.Clear();
            tenphong_phong.Clear();
            loaiphong_phong.Clear();
        }

        private void delete_phong_Click(object sender, EventArgs e)
        {
            phong.Delete_Phong(maphong_phong.Text);
            tabPage5_Click(sender, e);
            maphong_phong.Clear();
            tenphong_phong.Clear();
            loaiphong_phong.Clear();
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void data_phong_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            maphong_phong.Text = data_phong.Rows[i].Cells[0].Value.ToString();
            tenphong_phong.Text = data_phong.Rows[i].Cells[1].Value.ToString();
            loaiphong_phong.Text = data_phong.Rows[i].Cells[2].Value.ToString();
        }

        private void update_giaphong_Click(object sender, EventArgs e)
        {
            GiaPhong_DTO ob = new GiaPhong_DTO(loaiphong_giaphong.Text, Double.Parse(gia_giaphong.Text));
            giaphong.Update_Giaphong(ob);
            tabPage6_Click(sender, e);
            loaiphong_giaphong.Clear();
            gia_giaphong.Clear();
        }

        private void delete_giaphong_Click(object sender, EventArgs e)
        {
            giaphong.Delete_GiaPhong(loaiphong_giaphong.Text);
            tabPage6_Click(sender, e);
            loaiphong_giaphong.Clear();
            gia_giaphong.Clear();
        }

        private void insert_giaphong_Click(object sender, EventArgs e)
        {
            GiaPhong_DTO ob = new GiaPhong_DTO(loaiphong_giaphong.Text, Double.Parse(gia_giaphong.Text));
            giaphong.Insert_Giaphong(ob);
            tabPage6_Click(sender, e);
            loaiphong_giaphong.Clear();
            gia_giaphong.Clear();
        }

        private void data_giaphong_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            loaiphong_giaphong.Text = data_giaphong.Rows[i].Cells[0].Value.ToString();
            gia_giaphong.Text = data_giaphong.Rows[i].Cells[1].Value.ToString();
        }

        private void tabPage6_Click(object sender, EventArgs e)
        {
            data_giaphong.DataSource = giaphong.Load_GiaPhong();
        }

        private void phòng_Click(object sender, EventArgs e)
        {

        }

        private void machucvu_nhanvien_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage7_Click(object sender, EventArgs e)
        {
            data_nhanvien.DataSource = nhanvien.Load_Nhanvien();
        }

        private void sogiocong_nhanvien_TextChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void data_nhanvien_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            ID_nhanvien.Text = data_nhanvien.Rows[i].Cells[0].Value.ToString();
            name_nhanvien.Text = data_nhanvien.Rows[i].Cells[1].Value.ToString();
            date_nhanvien.Text = Ngay(data_nhanvien.Rows[i].Cells[2].Value.ToString());
            sdt_nhanvien.Text = data_nhanvien.Rows[i].Cells[3].Value.ToString();
            machucvu_nhanvien.Text = data_nhanvien.Rows[i].Cells[4].Value.ToString();
            sogiocong_nhanvien.Text = data_nhanvien.Rows[i].Cells[5].Value.ToString();
        }

        private void insert_nhanvien_Click(object sender, EventArgs e)
        {
            NhanVien_DTO ob = new NhanVien_DTO(ID_nhanvien.Text, name_nhanvien.Text, date_nhanvien.Text, sdt_nhanvien.Text, machucvu_nhanvien.Text, int.Parse(sogiocong_nhanvien.Text));
            nhanvien.Insert_KhachHang(ob);
            tabPage7_Click(sender, e);
            ID_nhanvien.Clear();
            name_nhanvien.Clear();
            date_nhanvien.Clear();
            machucvu_nhanvien.Clear();
            sogiocong_nhanvien.Clear();
        }

        private void delete_nhanvien_Click(object sender, EventArgs e)
        {
            nhanvien.Delete_khachhang(ID_nhanvien.Text);
            tabPage7_Click(sender, e);
            ID_nhanvien.Clear();
            name_nhanvien.Clear();
            date_nhanvien.Clear();
            machucvu_nhanvien.Clear();
            sogiocong_nhanvien.Clear();
        }

        private void update_nhanvien_Click(object sender, EventArgs e)
        {
            NhanVien_DTO ob = new NhanVien_DTO(ID_nhanvien.Text, name_nhanvien.Text, date_nhanvien.Text, sdt_nhanvien.Text, machucvu_nhanvien.Text, int.Parse(sogiocong_nhanvien.Text));
            nhanvien.Update_KhachHang(ob);
            tabPage7_Click(sender, e);
            ID_nhanvien.Clear();
            name_nhanvien.Clear();
            date_nhanvien.Clear();
            machucvu_nhanvien.Clear();
            sogiocong_nhanvien.Clear();
        }

        private void tabPage8_Click(object sender, EventArgs e)
        {
            data_chucvu.DataSource = chucvu.Load_Chucvu();
        }

        private void s_Click(object sender, EventArgs e)
        {

        }

        private void data_chucvu_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = e.RowIndex;
            machucvu_chucvu.Text = data_chucvu.Rows[i].Cells[0].Value.ToString();
            tenchucvu_chucvu.Text = data_chucvu.Rows[i].Cells[1].Value.ToString();
            luongcung_chucvu.Text = data_chucvu.Rows[i].Cells[2].Value.ToString();
        }

        private void update_chucvu_Click(object sender, EventArgs e)
        {
            ChucVu_DTO ob = new ChucVu_DTO(machucvu_chucvu.Text, tenchucvu_chucvu.Text, Double.Parse(luongcung_chucvu.Text));
            chucvu.Update_Chucvu(ob);
            tabPage8_Click(sender, e);
            machucvu_chucvu.Clear();
            tenchucvu_chucvu.Clear();
            luongcung_chucvu.Clear();
        }

        private void delete_chucvu_Click(object sender, EventArgs e)
        {
            chucvu.Delete_Chucvu(machucvu_chucvu.Text);
            tabPage8_Click(sender, e);
            machucvu_chucvu.Clear();
            tenchucvu_chucvu.Clear();
            luongcung_chucvu.Clear();
        }

        private void insert_chucvu_Click(object sender, EventArgs e)
        {
            ChucVu_DTO ob = new ChucVu_DTO(machucvu_chucvu.Text, tenchucvu_chucvu.Text, Double.Parse(luongcung_chucvu.Text));
            chucvu.Insert_Chucvu(ob);
            tabPage8_Click(sender, e);
            machucvu_chucvu.Clear();
            tenchucvu_chucvu.Clear();
            luongcung_chucvu.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String ans = "";
            if (chonbang_baocao.SelectedItem.ToString() == "chức vụ") ans = "chucvu";
            else if (chonbang_baocao.SelectedItem.ToString() == "giá phòng") ans = "giaphong";
            else if (chonbang_baocao.SelectedItem.ToString() == "khách hàng") ans = "khachhang";
            else if (chonbang_baocao.SelectedItem.ToString() == "nhân viên") ans = "nhanvien";
            else if (chonbang_baocao.SelectedItem.ToString() == "phòng") ans = "phong";
            else if (chonbang_baocao.SelectedItem.ToString() == "phòng đang thuê") ans = "phongdangthue";
            if (ans == "chucvu" || ans == "giaphong" || ans == "khachhang" || ans == "phong" || ans == "nhanvien" || ans == "phongdangthue")
            {
                string sql = "select * from " + ans;
                if (lenh.Text != "")
                {
                    sql = sql + " where " + lenh.Text;
                }
                //string sql = "select * from phongdangthue";
                DataTable dt = new DataTable();
                dt = Lay_DL(sql);
                reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
                reportViewer1.LocalReport.ReportPath = @"C:\Users\hieu8\source\repos\baitaplon\baitaplon\" + ans + ".rdlc";
                if (dt.Rows.Count > 0)
                {
                    ReportDataSource rds = new ReportDataSource();
                    rds.Name = "QLKS";
                    rds.Value = dt;
                    reportViewer1.LocalReport.DataSources.Clear();
                    reportViewer1.LocalReport.DataSources.Add(rds);
                    reportViewer1.RefreshReport();
                }
                else MessageBox.Show("khong co du lieu");
            }
            else MessageBox.Show("khong co du lieu");
        }

        private void tabPage9_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'quanLyKhachSanDataSet1.chucvu' table. You can move, or remove it, as needed.
            this.chucvuTableAdapter.Fill(this.quanLyKhachSanDataSet1.chucvu);

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void chucvuBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        

        private void tabPage10_Click(object sender, EventArgs e)
        {

        }

        private void chonmuc_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            string filePath = "";
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Excel | *.xlsx | Excel | *.xls";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath = dialog.FileName;
            }
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Đường dẫn báo cáo không hợp lệ");
                return;
            }
            using (ExcelPackage p = new ExcelPackage())
            {
                byte[] bin = new byte[] { };
                int n = 1;
                foreach (var chon in chonmuc.CheckedItems)
                {
                    if (chon.ToString() == "chức vụ")
                    {
                        data_xuatExcel.DataSource = chucvu.Load_Chucvu();
                        p.Workbook.Worksheets.Add("chức vụ");
                        ExcelWorksheet ws = p.Workbook.Worksheets[n];
                        ws.Cells.Style.Font.Size = 11;
                        ws.Cells.Style.Font.Name = "Calibri";
                        string[] arrColumnHeader = { "mã chức vụ", "tên chức vụ", "lương cứng" };
                        var countColHeader = arrColumnHeader.Count();
                        ws.Cells[1, 1].Value = "Thống kê chức vụ";
                        ws.Cells[1, 1, 1, countColHeader].Merge = true;
                        ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                        int colIndex = 1;
                        int rowIndex = 2;
                        foreach (var item in arrColumnHeader)
                        {
                            var cell = ws.Cells[rowIndex, colIndex];
                            cell.Value = item;
                            colIndex++;
                        }
                        List<ChucVu> userList = new List<ChucVu>();
                        for (int i = 0; i < data_xuatExcel.Rows.Count - 1; i++)
                        {
                            ChucVu ob = new ChucVu();

                            ob.machucvu = data_xuatExcel.Rows[i].Cells[0].Value.ToString();
                            ob.tenchucvu = data_xuatExcel.Rows[i].Cells[1].Value.ToString();
                            ob.luongcung = data_xuatExcel.Rows[i].Cells[2].Value.ToString();
                            userList.Add(ob);
                        }
                        foreach (var item in userList)
                        {
                            colIndex = 1;
                            rowIndex++;
                            ws.Cells[rowIndex, colIndex++].Value = item.machucvu;
                            ws.Cells[rowIndex, colIndex++].Value = item.tenchucvu;
                            ws.Cells[rowIndex, colIndex++].Value = item.luongcung;
                        }
                        n++;
                    }
                    else if (chon.ToString() == "giá phòng")
                    {
                        data_xuatExcel.DataSource = phong.Load_Phong();
                        data_xuatExcel.DataSource = chucvu.Load_Chucvu();
                        p.Workbook.Worksheets.Add("giá phòng");
                        ExcelWorksheet ws = p.Workbook.Worksheets[n];
                        ws.Cells.Style.Font.Size = 11;
                        ws.Cells.Style.Font.Name = "Calibri";
                        string[] arrColumnHeader = { "loaiphong", "gia" };
                        var countColHeader = arrColumnHeader.Count();
                        ws.Cells[1, 1].Value = "Thống kê giá phòng";
                        ws.Cells[1, 1, 1, countColHeader].Merge = true;
                        ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                        int colIndex = 1;
                        int rowIndex = 2;
                        foreach (var item in arrColumnHeader)
                        {
                            var cell = ws.Cells[rowIndex, colIndex];
                            cell.Value = item;
                            colIndex++;
                        }
                        List<GiaPhong> userList = new List<GiaPhong>();
                        for (int i = 0; i < data_xuatExcel.Rows.Count - 1; i++)
                        {
                            GiaPhong ob = new GiaPhong();

                            ob.loaiphong = data_xuatExcel.Rows[i].Cells[0].Value.ToString();
                            ob.gia = data_xuatExcel.Rows[i].Cells[1].Value.ToString();
                            userList.Add(ob);
                        }
                        foreach (var item in userList)
                        {
                            colIndex = 1;
                            rowIndex++;
                            ws.Cells[rowIndex, colIndex++].Value = item.loaiphong;
                            ws.Cells[rowIndex, colIndex++].Value = item.gia;
                        }
                        n++;
                    }
                    else if (chon.ToString() == "khách hàng")
                    {
                        data_xuatExcel.DataSource = khachhang.Load_KhachHang();
                        p.Workbook.Worksheets.Add("khách hàng");
                        ExcelWorksheet ws = p.Workbook.Worksheets[n];
                        ws.Cells.Style.Font.Size = 11;
                        ws.Cells.Style.Font.Name = "Calibri";
                        string[] arrColumnHeader = { "ID", "họ và tên", "ngày sinh", "số điện thoại" };
                        var countColHeader = arrColumnHeader.Count();
                        ws.Cells[1, 1].Value = "Thống kê khách hàng";
                        ws.Cells[1, 1, 1, countColHeader].Merge = true;
                        ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                        int colIndex = 1;
                        int rowIndex = 2;
                        foreach (var item in arrColumnHeader)
                        {
                            var cell = ws.Cells[rowIndex, colIndex];
                            cell.Value = item;
                            colIndex++;
                        }
                        List<KhachHang> userList = new List<KhachHang>();
                        for (int i = 0; i < data_xuatExcel.Rows.Count - 1; i++)
                        {
                            KhachHang ob = new KhachHang();
                            ob.ID = data_xuatExcel.Rows[i].Cells[0].Value.ToString();
                            ob.hovaten = data_xuatExcel.Rows[i].Cells[1].Value.ToString();
                            ob.ngaysinh = data_xuatExcel.Rows[i].Cells[2].Value.ToString();
                            ob.sdt = data_xuatExcel.Rows[i].Cells[3].Value.ToString();
                            userList.Add(ob);
                        }
                        foreach (var item in userList)
                        {
                            colIndex = 1;
                            rowIndex++;
                            ws.Cells[rowIndex, colIndex++].Value = item.ID;
                            ws.Cells[rowIndex, colIndex++].Value = item.hovaten;
                            ws.Cells[rowIndex, colIndex++].Value = item.ngaysinh;
                            ws.Cells[rowIndex, colIndex++].Value = item.sdt;
                        }
                        n++;
                    }
                    else if (chon.ToString() == "nhân viên")
                    {
                        data_xuatExcel.DataSource = nhanvien.Load_Nhanvien();
                        p.Workbook.Worksheets.Add("nhân viên");
                        ExcelWorksheet ws = p.Workbook.Worksheets[n];
                        ws.Cells.Style.Font.Size = 11;
                        ws.Cells.Style.Font.Name = "Calibri";
                        string[] arrColumnHeader = { "ID", "họ và tên", "ngày sinh", "số điện thoại", "mã chức vụ" };
                        var countColHeader = arrColumnHeader.Count();
                        ws.Cells[1, 1].Value = "Thống kê nhân viên";
                        ws.Cells[1, 1, 1, countColHeader].Merge = true;
                        ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                        int colIndex = 1;
                        int rowIndex = 2;
                        foreach (var item in arrColumnHeader)
                        {
                            var cell = ws.Cells[rowIndex, colIndex];
                            cell.Value = item;
                            colIndex++;
                        }
                        List<NhanVien> userList = new List<NhanVien>();
                        for (int i = 0; i < data_xuatExcel.Rows.Count - 1; i++)
                        {
                            NhanVien ob = new NhanVien();
                            ob.id = data_xuatExcel.Rows[i].Cells[0].Value.ToString();
                            ob.hovaten = data_xuatExcel.Rows[i].Cells[1].Value.ToString();
                            ob.ngaysinh = data_xuatExcel.Rows[i].Cells[2].Value.ToString();
                            ob.sdt = data_xuatExcel.Rows[i].Cells[3].Value.ToString();
                            ob.machucvu = data_xuatExcel.Rows[i].Cells[4].Value.ToString();
                            ob.sogiocong = data_xuatExcel.Rows[i].Cells[5].Value.ToString();
                            userList.Add(ob);
                        }
                        foreach (var item in userList)
                        {
                            colIndex = 1;
                            rowIndex++;
                            ws.Cells[rowIndex, colIndex++].Value = item.id;
                            ws.Cells[rowIndex, colIndex++].Value = item.hovaten;
                            ws.Cells[rowIndex, colIndex++].Value = item.ngaysinh;
                            ws.Cells[rowIndex, colIndex++].Value = item.sdt;
                            ws.Cells[rowIndex, colIndex++].Value = item.machucvu;
                            ws.Cells[rowIndex, colIndex++].Value = item.sogiocong;
                        }
                        n++;
                    }
                    else if (chon.ToString() == "phòng")
                    {
                        data_xuatExcel.DataSource = phong.Load_Phong();
                        p.Workbook.Worksheets.Add("phòng");
                        ExcelWorksheet ws = p.Workbook.Worksheets[n];
                        ws.Cells.Style.Font.Size = 11;
                        ws.Cells.Style.Font.Name = "Calibri";
                        string[] arrColumnHeader = { "mã phòng", "tên phòng", "loại phòng" };
                        var countColHeader = arrColumnHeader.Count();
                        ws.Cells[1, 1].Value = "Thống kê phòng";
                        ws.Cells[1, 1, 1, countColHeader].Merge = true;
                        ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                        int colIndex = 1;
                        int rowIndex = 2;
                        foreach (var item in arrColumnHeader)
                        {
                            var cell = ws.Cells[rowIndex, colIndex];
                            cell.Value = item;
                            colIndex++;
                        }
                        List<Phong> userList = new List<Phong>();
                        for (int i = 0; i < data_xuatExcel.Rows.Count - 1; i++)
                        {
                            Phong ob = new Phong();

                            ob.maphong = data_xuatExcel.Rows[i].Cells[0].Value.ToString();
                            ob.tenphong = data_xuatExcel.Rows[i].Cells[1].Value.ToString();
                            ob.loaiphong = data_xuatExcel.Rows[i].Cells[2].Value.ToString();
                            userList.Add(ob);
                        }
                        foreach (var item in userList)
                        {
                            colIndex = 1;
                            rowIndex++;
                            ws.Cells[rowIndex, colIndex++].Value = item.maphong;
                            ws.Cells[rowIndex, colIndex++].Value = item.tenphong;
                            ws.Cells[rowIndex, colIndex++].Value = item.loaiphong;
                        }
                        n++;
                    }
                    else if (chon.ToString() == "phòng đang thuê")
                    {
                        data_xuatExcel.DataSource = phongdangthue.Load_PhongDangThue();
                        p.Workbook.Worksheets.Add("phòng đang thuê");
                        ExcelWorksheet ws = p.Workbook.Worksheets[n];
                        ws.Cells.Style.Font.Size = 11;
                        ws.Cells.Style.Font.Name = "Calibri";
                        string[] arrColumnHeader = { "ID", "mã phòng", "ngày nhận phong", "ngày trả phòng" };
                        var countColHeader = arrColumnHeader.Count();
                        ws.Cells[1, 1].Value = "Thống kê phòng đang thuê";
                        ws.Cells[1, 1, 1, countColHeader].Merge = true;
                        ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                        int colIndex = 1;
                        int rowIndex = 2;
                        foreach (var item in arrColumnHeader)
                        {
                            var cell = ws.Cells[rowIndex, colIndex];
                            cell.Value = item;
                            colIndex++;
                        }
                        List<PhongDangThue> userList = new List<PhongDangThue>();
                        for (int i = 0; i < data_xuatExcel.Rows.Count - 1; i++)
                        {
                            PhongDangThue ob = new PhongDangThue();

                            ob.id = data_xuatExcel.Rows[i].Cells[0].Value.ToString();
                            ob.maphong = data_xuatExcel.Rows[i].Cells[1].Value.ToString();
                            ob.ngaynhanphong = data_xuatExcel.Rows[i].Cells[2].Value.ToString();
                            ob.ngaytraphong = data_xuatExcel.Rows[i].Cells[3].Value.ToString();
                            userList.Add(ob);
                        }
                        foreach (var item in userList)
                        {
                            colIndex = 1;
                            rowIndex++;
                            ws.Cells[rowIndex, colIndex++].Value = item.id;
                            ws.Cells[rowIndex, colIndex++].Value = item.maphong;
                            ws.Cells[rowIndex, colIndex++].Value = item.ngaynhanphong;
                            ws.Cells[rowIndex, colIndex++].Value = item.ngaytraphong;
                        }
                        n++;
                    }

                }
                bin = p.GetAsByteArray();
                File.WriteAllBytes(filePath, bin);
                MessageBox.Show("Xuất excel thành công!");
            }
        }
    }
}