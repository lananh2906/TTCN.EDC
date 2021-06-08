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
    public partial class FrmPhieuNhapKho : Form
    {
        DataTable tableBC;
        string sql;
        
        public FrmPhieuNhapKho()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void FrmPhieuNhapKho_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            sql = "SELECT * FROM tblphieunhapkho WHERE NgayLapPNK = N'" + dtpNgayBan.Text + "'";
            tableBC = Functions.GetDataToTable(sql);

            dataGridView1.DataSource = tableBC;
            if (tableBC.Rows.Count == 0)
            {
                MessageBox.Show("Không có hóa đơn thỏa mãn điều kiện!!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                

            }
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            // Khởi động chương trình Excel
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook; //Trong 1 chương trình Excel có nhiều Workbook
            COMExcel.Worksheet exSheet; //Trong 1 Workbook có nhiều Worksheet
            COMExcel.Range exRange;
            string sql;
            int hang = 0, cot = 0;
            DataTable tableThongTinNhap;
            exBook = exApp.Workbooks.Add(COMExcel.XlWBATemplate.xlWBATWorksheet);
            exSheet = exBook.Worksheets[1];
            // Định dạng chung
            exRange = exSheet.Cells[1, 1];
            exRange.Range["A1:B3"].Font.Size = 10;
            exRange.Range["A1:B3"].Font.Name = "Times new roman";
            exRange.Range["A1:B3"].Font.Bold = true;
            exRange.Range["A1:B3"].Font.ColorIndex = 5; //Màu xanh da trời
            exRange.Range["A1:A1"].ColumnWidth = 15;
            exRange.Range["B1:B1"].ColumnWidth = 20;
            exRange.Range["A1:B1"].MergeCells = true;
            exRange.Range["A1:B1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A1:B1"].Value = "Công tu cổ phần EDCZOne";

            exRange.Range["A2:B2"].MergeCells = true;
            exRange.Range["A2:B2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:B2"].Value = "96 Trần Phú-Ba Đình_HN";

            exRange.Range["A3:B3"].MergeCells = true;
            exRange.Range["A3:B3"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A3:B3"].Value = "Điện thoại: 0123456789";


            exRange.Range["C2:F2"].Font.Size = 16;
            exRange.Range["C2:F2"].Font.Name = "Times new roman";
            exRange.Range["C2:F2"].Font.Bold = true;
            exRange.Range["C2:F2"].Font.ColorIndex = 3; //Màu đỏ
            exRange.Range["C2:F2"].MergeCells = true;
            exRange.Range["C2:F2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:F2"].Value = "PHIẾU NHẬP KHO";

            exRange.Range["B4:B4"].MergeCells = true;
            exRange.Range["B4:B4"].Font.Bold = true;
            exRange.Range["B4:B4"].Font.Italic = true;
            exRange.Range["B4:B4"].Value = "Kì báo cáo: ";

            exRange.Range["C4:C4"].MergeCells = true;




            sql = "SELECT * FROM tblphieunhapkho WHERE NgayLapPNK = N'" + dtpNgayBan.Text + "'";
            tableThongTinNhap = Functions.GetDataToTable(sql);
            //Tạo dòng tiêu đề bảng
            exRange.Range["A5:F5"].Font.Bold = true;
            exRange.Range["A5:F5"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C5:F5"].ColumnWidth = 15;
            exRange.Range["A5:A5"].Value = "STT";
            exRange.Range["B5:B5"].Value = "MaPNK";
            exRange.Range["C5:C5"].Value = "Mã Nhân Viên";
            exRange.Range["D5:D5"].Value = "Ngày Nhập";
            
            for (hang = 0; hang <= tableThongTinNhap.Rows.Count - 1; hang++)
            {
                //Điền số thứ tự vào cột 1 từ dòng 12
                exSheet.Cells[1][hang + 6] = hang + 1;
                for (cot = 0; cot <= tableThongTinNhap.Columns.Count - 1; cot++)
                    //Điền thông tin hàng từ cột thứ 2, dòng 12
                    exSheet.Cells[cot + 2][hang + 6] = tableThongTinNhap.Rows[hang][cot].ToString();
            }
            

            exRange = exSheet.Cells[4][hang + 10];
            exRange.Range["A1:C1"].MergeCells = true;
            exRange.Range["A1:C1"].Font.Italic = true;
            exRange.Range["A1:C1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            int day = DateTime.Now.Day;
            int month = DateTime.Now.Month;
            int year = DateTime.Now.Year;
            exRange.Range["A1:C1"].Value = "Hà Nội, ngày " + day + " tháng " + month + " năm " + year;

            exRange.Range["A2:C2"].MergeCells = true;
            exRange.Range["A2:C2"].Font.Italic = true;
            exRange.Range["A2:C2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:C2"].Value = "Nhân viên lập báo cáo";

            exRange.Range["A3:C3"].MergeCells = true;
            exRange.Range["A3:C3"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A3:C3"].Value = "(Kí, Ghi rõ họ tên)";

            exSheet.Name = "Phiếu NHập Kho";
            exApp.Visible = true;
        }
    }
}
