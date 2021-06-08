using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EDCZONE.Class;

namespace EDCZONE
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            Functions.Connect(); //Mở kết nối
        }

        private void mnukhachhang_Click(object sender, EventArgs e)
        {
            FrmKhachHang frm = new FrmKhachHang();
            frm.Show();
        }

        private void mnunhacc_Click(object sender, EventArgs e)
        {
            FrmNhaCungCap frm = new FrmNhaCungCap();
            frm.Show();
        }

        private void mnusanpham_Click(object sender, EventArgs e)
        {
            FrmSanPham frm = new FrmSanPham();
            frm.Show();

        }

        private void mnuhoadonban_Click(object sender, EventArgs e)
        {
            FrmHoaDonBan frm = new FrmHoaDonBan();
            frm.Show();
        }

        private void mnuhoadonnhap_Click(object sender, EventArgs e)
        {
            FrmHoaDonNhapHang frm = new FrmHoaDonNhapHang();
            frm.Show();
        }

        private void mnutimkiemsp_Click(object sender, EventArgs e)
        {
            FrmTimKiemSanPham frm = new FrmTimKiemSanPham();
            frm.Show();
        }

        private void mnuhangtonkho_Click(object sender, EventArgs e)
        {
            FrmBaoCaoHangTonKho frm = new FrmBaoCaoHangTonKho();
            frm.Show();
        }

        private void mnuphieunhapkho_Click(object sender, EventArgs e)
        {
            FrmPhieuNhapKho frm = new FrmPhieuNhapKho();
            frm.Show();
        }

        private void mnuphieuxuatkho_Click(object sender, EventArgs e)
        {
            FrmPhieuXuatKho frm = new FrmPhieuXuatKho();
            frm.Show();
        }

        private void mnunhanvien_Click(object sender, EventArgs e)
        {
            FrmNhanVien frm = new FrmNhanVien();
            frm.Show();
        }
    }
}
