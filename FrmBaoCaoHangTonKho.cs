using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EDCZONE.Class;
using System.Data.SqlClient;
using COMExcel = Microsoft.Office.Interop.Excel;

namespace EDCZONE
{
    public partial class FrmBaoCaoHangTonKho : Form
    {
        DataTable tblBaoCao;
        public FrmBaoCaoHangTonKho()
        {
            InitializeComponent();
        }

        private void dgv_BaoCao_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaSP.Text = dgv_BaoCao.CurrentRow.Cells["MaSP"].Value.ToString();
            txtTenSP.Text = dgv_BaoCao.CurrentRow.Cells["TenSP"].Value.ToString();
            txtDonGia.Text = dgv_BaoCao.CurrentRow.Cells["DonGia"].Value.ToString();
            txtMaNCC.Text = dgv_BaoCao.CurrentRow.Cells["SoLuong"].Value.ToString();
            txtSoLuong.Text = dgv_BaoCao.CurrentRow.Cells["SoLuong"].Value.ToString();
            txtDVT.Text = dgv_BaoCao.CurrentRow.Cells["DVT"].Value.ToString();
            

        }

        private void btnBaoCao_Click(object sender, EventArgs e)
        {
            string sql;
            sql = "SELECT * FROM tblsanpham";
            tblBaoCao = Functions.GetDataToTable(sql);

            //Đọc dữ liệu từ bảng
            dgv_BaoCao.DataSource = tblBaoCao; //Nguồn dữ liệu 
        }

        private void dgv_BaoCao_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void FrmBaoCaoHangTonKho_Load(object sender, EventArgs e)
        {
            LoadDataGridView();
        }
        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT * FROM tblsanpham";
            tblBaoCao = Functions.GetDataToTable(sql);

            //Đọc dữ liệu từ bảng
            dgv_BaoCao.DataSource = tblBaoCao; //Nguồn dữ liệu            



           
        }
        private void ResetValue()
        {
            txtMaSP.Text = "";
            txtTenSP.Text = "";
            txtDonGia.Text = "";
            txtMaNCC.Text = "";
            txtSoLuong.Text = "";
            txtDVT.Text = "";




        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            ResetValue();
            dgv_BaoCao.DataSource = null;
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook;
            COMExcel.Worksheet exSheet;
            COMExcel.Range exRange;
            string sql;
            int hang = 0, cot = 0;
            DataTable tblThongTinSP_TonKho;
            exBook = exApp.Workbooks.Add(COMExcel.XlWBATemplate.xlWBATWorksheet);
            exSheet = exBook.Worksheets[1];
            // Định dạng chung
            exRange = exSheet.Cells[1, 1];
            exRange.Range["A1:Z300"].Font.Name = "Times new roman"; //Font chữ
            exRange.Range["A1:B3"].Font.Size = 10;
            exRange.Range["A1:B3"].Font.Bold = true;
            exRange.Range["A1:B3"].Font.ColorIndex = 5; //Màu xanh da trời
            exRange.Range["A1:A1"].ColumnWidth = 7;
            exRange.Range["B1:B1"].ColumnWidth = 15;
            exRange.Range["A1:B1"].MergeCells = true;
            exRange.Range["A1:B1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A1:B1"].Value = "Công Ty Cổ Phần EDCZONE";
            exRange.Range["A2:B2"].MergeCells = true;
            exRange.Range["A2:B2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:B2"].Value = "96 Trần Phú- Ba Đình - Hà Nội";
            exRange.Range["A3:B3"].MergeCells = true;
            exRange.Range["A3:B3"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A3:B3"].Value = "Điện thoại: (04)38526419";
            exRange.Range["C2:E2"].Font.Size = 16;
            exRange.Range["C2:E2"].Font.Bold = true;
            exRange.Range["C2:E2"].Font.ColorIndex = 3; //Màu đỏ
            exRange.Range["C2:E2"].MergeCells = true;
            exRange.Range["C2:E2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:E2"].Value = "BÁO CÁO SẢN PHẨM TỒN KHO";

            sql = "SELECT * FROM tblsanpham";
            tblThongTinSP_TonKho = Functions.GetDataToTable(sql);
            exRange.Range["A6:M6"].Font.Bold = true;
            exRange.Range["A6:M6"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C6:O6"].ColumnWidth = 12;
            exRange.Range["A6:A6"].Value = "STT";
            exRange.Range["B6:B6"].Value = "Mã sản phẩm";
            exRange.Range["C6:C6"].Value = "Tên Sản Phẩm";
            exRange.Range["D6:D6"].Value = "Có thể bán";
            exRange.Range["E6:E6"].Value = "Mô tả SP";
            exRange.Range["F6:F6"].Value = "Đơn giá";
            exRange.Range["G6:G6"].Value = "Đơn vị tính";
            exRange.Range["E6:E6"].Value = "Mã NCC";

            for (hang = 0; hang < tblThongTinSP_TonKho.Rows.Count; hang++)
            {
                //Điền số thứ tự vào cột 1 từ dòng 7
                exSheet.Cells[1][hang + 7] = hang + 1;
                for (cot = 0; cot < tblThongTinSP_TonKho.Columns.Count; cot++)
                //Điền thông tin hàng từ cột thứ 2, dòng 7
                {
                    exSheet.Cells[cot + 2][hang + 7] = tblThongTinSP_TonKho.Rows[hang][cot].ToString();
                }
            }

            exRange.Range["D20:F20"].MergeCells = true;
            exRange.Range["D20:F20"].Font.Italic = true;
            exRange.Range["D20:F20"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            int day = DateTime.Now.Day;
            int month = DateTime.Now.Month;
            int year = DateTime.Now.Year;
            exRange.Range["D20:F20"].Value = "Hà Nội, ngày " + day + " tháng " + month + " năm " + year;
            exRange.Range["D21:F21"].MergeCells = true;
            exRange.Range["D21:F21"].Font.Italic = true;
            exRange.Range["D21:F21"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["D21:F21"].Value = "Nhân viên lập báo cáo";
            exRange.Range["D22:F22"].MergeCells = true;
            exRange.Range["D22:F22"].Font.Italic = true;
            exRange.Range["D22:F22"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["D22:F22"].Value = "(Kí và ghi rõ họ tên)";
            exSheet.Name = "Báo cáo";
            exApp.Visible = true;

        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("bạn có chắc chắn muốn thoát chương trình không", "Hỏi Thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                this.Close();
        }
    }
}
