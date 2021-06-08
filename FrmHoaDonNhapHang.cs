using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using EDCZONE.Class;
using COMExcel = Microsoft.Office.Interop.Excel;

namespace EDCZONE
{
    public partial class FrmHoaDonNhapHang : Form
    {
        DataTable tblCTHDN;
        public FrmHoaDonNhapHang()
        {
            InitializeComponent();
        }

        private void FrmHoaDonNhapHang_Load(object sender, EventArgs e)
        {
            btnThemHĐ.Enabled = true;
            btnLuuHD.Enabled = false;
            btnXoaHĐ.Enabled = false;

            txtSoHDN.ReadOnly = true;
            txtTenNV.ReadOnly = true;
            txtTenNCC.ReadOnly = true;
            txtThanhTien.Enabled = true;
            textTongTien.ReadOnly = true;
            txtTenSP.ReadOnly = true;
            txtSDT.ReadOnly = true;
            txtDiaChi.ReadOnly = true;
            txtGiamGia.Text = "0";
            textTongTien.Text = "0";
            txtThanhTien.Text = "";

            Functions.FillCombo1("select MaNV from tblnhanhvien", cboMaNV, "MaNV");
            cboMaNV.SelectedIndex = -1;
            Functions.FillCombo1("select MaNCC from tblncc", cboMaNCC, "MaNCC");
            cboMaNCC.SelectedIndex = -1;
            Functions.FillCombo1("select MaSP from tblsanpham", cboMaSP, "MaSP");
            cboMaSP.SelectedIndex = -1;

            if (txtSoHDN.Text != "")
            {
                LoadInfoHoaDon();
                btnXoaHĐ.Enabled = true;

            }
            LoadDataGridView();
        }
        private void LoadDataGridView()
        {
            string sql;
            sql = "Select * from tblchitietphieunhaphang";
            tblCTHDN = Functions.GetDataToTable(sql);
            dataGridView_HDN.DataSource = tblCTHDN;
            dataGridView_HDN.AllowUserToAddRows = false;
            dataGridView_HDN.EditMode = DataGridViewEditMode.EditProgrammatically;
        }
        private void LoadInfoHoaDon()
        {
            string str;
            str = "Select NgayLapPNK from tblphieunhaphang where MaPNH = N'" + txtSoHDN.Text + "'";
            dtpNgayNhap.Value = DateTime.Parse(Functions.GetFieldValues(str));
            str = "Select MaNV from tblphieunhaphang where MaPNH = N'" + txtSoHDN.Text + "'";
            cboMaNV.Text = Functions.GetFieldValues(str);
            str = "Select MaNCC from tblphieunhaphang where MaPNH = N'" + txtSoHDN.Text + "'";
            cboMaNCC.Text = Functions.GetFieldValues(str);
            str = "Select TongTien from tblphieunhaphang where MaPNH = N'" + txtSoHDN.Text + "'";
            textTongTien.Text = Functions.GetFieldValues(str);
            tblBangChu.Text = "Bằng chữ: " + Functions.ChuyenSoSangChu(textTongTien.Text);
        }

        private void btnThemHĐ_Click(object sender, EventArgs e)
        {
            btnXoaHĐ.Enabled = false;
            btnLuuHD.Enabled = true;

            btnThemHĐ.Enabled = false;
            ResetValues();
            txtSoHDN.Text = Functions.CreateKey("HDN");
            LoadDataGridView();
        }
        private void ResetValues()
        {
            
            cboMaSP.Text = "";
            txtSoLuong.Text = "";
            txtGiamGia.Text = "0";
            txtDonGia.Text = "";
            txtThanhTien.Text = "0";
        }

        private void btnLuuHD_Click(object sender, EventArgs e)
        {
            string sql;
            Double sl, SLcon, tong, Tongmoi;
            sql = "SELECT MaPNH FROM tblphieunhaphang WHERE MaPNH='" + txtSoHDN.Text + "'";
            if (!Functions.CheckKey(sql))
            {
                // Mã hóa đơn chưa có, tiến hành lưu các thông tin chung
                // Mã HDBan được sinh tự động do đó không có trường hợp trùng khóa

                sql = "INSERT INTO tblphieunhaphang(MaPNH, MaNCC, MaNV, NgayLapPNH, TongTien) VALUES (N'" + txtSoHDN.Text.Trim() + "',N'" + cboMaNCC.SelectedValue +
                      "', N'" + cboMaNV.SelectedValue + "','" + dtpNgayNhap.Value + "','" + textTongTien.Text + "')";
                Functions.RunSQL(sql);
            }
            // Lưu thông tin của các mặt hàng

            sl = Convert.ToDouble(Functions.GetFieldValues("SELECT SoLuong FROM tblsanpham WHERE MaSP = N'" + cboMaSP.SelectedValue + "'"));

            sql = "INSERT INTO tblchitietphieunhaphang(MaPNH,MaSP,SoLuong,DonGiaNhap,GiamGia,ThanhTien) VALUES(N'" + txtSoHDN.Text.Trim() + "','" + cboMaSP.SelectedValue + "','" + txtSoLuong.Text + "','" + txtDonGia.Text + "','" + txtGiamGia.Text + "','" + txtThanhTien.Text + "')";
            Functions.RunSQL(sql);
            LoadDataGridView();
            // Cập nhật lại số lượng của mặt hàng vào bảng SanPham
            SLcon = sl + Convert.ToDouble(txtSoLuong.Text);
            sql = "UPDATE tblsanpham SET SoLuong =" + SLcon + " WHERE MaSP= N'" + cboMaSP.SelectedValue + "'";
            Functions.RunSQL(sql);
            // Cập nhật lại tổng tiền cho hóa đơn nhập
            tong = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTien FROM tblphieunhaphang WHERE MaPNH = N'" + txtSoHDN.Text + "'"));
            Tongmoi = tong + Convert.ToDouble(txtThanhTien.Text);
            sql = "UPDATE tblphieunhaphang SET TongTien =" + Tongmoi + " WHERE MaPNH = N'" + txtSoHDN.Text + "'";
            Functions.RunSQL(sql);
            textTongTien.Text = Tongmoi.ToString();
            tblBangChu.Text = "Bằng chữ: " + Functions.ChuyenSoSangChuoi(Double.Parse(Tongmoi.ToString()));
            Functions.RunSQL(sql);
            ResetValues();
            btnXoaHĐ.Enabled = true;
            btnThemHĐ.Enabled = true;


        }
        private void cboMaSP_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaSP.Text == "")
            {
                txtTenSP.Text = "";
                txtDonGia.Text = "";
                txtGiamGia.Text = "0";
                txtSoLuong.Text = "";
                txtThanhTien.Text = "0";
            }
            //Khi chọn Mã giày dép thì các thông tin của giày dép sẽ hiện ra
            str = "Select TenSP from tblsanpham where MaSP = N'" + cboMaSP.SelectedValue + "'";
            txtTenSP.Text = Functions.GetFieldValues(str);
            str = "Select DonGia from tblsanpham where MaSP='" + cboMaSP.SelectedValue + "'";
            txtDonGia.Text = Functions.GetFieldValues(str);
        }

        private void cboMaNV_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaNV.Text == "")
                txtTenNV.Text = "";
            // Khi chọn Mã nhân viên thì tên nhân viên tự động hiện ra
            str = "Select HoTen from tblnhanhvien where MaNV =N'" + cboMaNV.SelectedValue + "'";
            txtTenNV.Text = Functions.GetFieldValues(str);
        }

        private void cboMaNCC_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaNCC.Text == "")
            {
                txtTenNCC.Text = "";
                txtDiaChi.Text = "";
                txtSDT.Text = "";
            }
            //Khi chọn Mã nhà cung cấp thì các thông tin của nhà cung cấp sẽ hiện ra
            str = "Select TenNCC from tblncc where MaNCC = N'" + cboMaNCC.SelectedValue + "'";
            txtTenNCC.Text = Functions.GetFieldValues(str);
            str = "Select DiaChi from tblncc where MaNCC = N'" + cboMaNCC.SelectedValue + "'";
            txtDiaChi.Text = Functions.GetFieldValues(str);
            str = "Select SDT from tblncc where MaNCC= N'" + cboMaNCC.SelectedValue + "'";
            txtSDT.Text = Functions.GetFieldValues(str);
        }

        private void txtSoLuong_TextChanged(object sender, EventArgs e)
        {
            double tt, sl, dg, gg;
            if (txtSoLuong.Text == "")
                sl = 0;
            else
                sl = Convert.ToDouble(txtSoLuong.Text);
            if (txtGiamGia.Text == "")
                gg = 0;
            else
                gg = Convert.ToDouble(txtGiamGia.Text);
            if (txtDonGia.Text == "")
                dg = 0;
            else
                dg = Convert.ToDouble(txtDonGia.Text);
            tt = sl * dg - sl * dg * gg / 100;
            txtThanhTien.Text = tt.ToString();
        }

        private void txtGiamGia_TextChanged(object sender, EventArgs e)
        {
            double tt, sl, dg, gg;
            if (txtSoLuong.Text == "")
                sl = 0;
            else
                sl = Convert.ToDouble(txtSoLuong.Text);
            if (txtGiamGia.Text == "")
                gg = 0;
            else
                gg = Convert.ToDouble(txtGiamGia.Text);
            if (txtDonGia.Text == "")
                dg = 0;
            else
                dg = Convert.ToDouble(txtDonGia.Text);
            tt = sl * dg - sl * dg * gg / 100;
            txtThanhTien.Text = tt.ToString();
        }


        private void dataGridView_HDN_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            string MaHangxoa, sql;
            Double ThanhTienxoa, SoLuongxoa, sl, slcon;
            if (tblCTHDN.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if ((MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                //Xóa hàng và cập nhật lại số lượng hàng 
                MaHangxoa = dataGridView_HDN.CurrentRow.Cells["MaSP"].Value.ToString();
                SoLuongxoa = Convert.ToDouble(dataGridView_HDN.CurrentRow.Cells["SoLuong"].Value.ToString());
                ThanhTienxoa = Convert.ToDouble(dataGridView_HDN.CurrentRow.Cells["ThanhTien"].Value.ToString());
                sql = "DELETE tblchitietphieunhaphang WHERE MaPNH=N'" + txtSoHDN.Text + "' AND MaSP = N'" + MaHangxoa + "'";
                Functions.RunSQL(sql);
                // Cập nhật lại số lượng cho các mặt hàng
                sl = Convert.ToDouble(Functions.GetFieldValues("SELECT SoLuong FROM tblsanpham WHERE MaSP = N'" + MaHangxoa + "'"));
                slcon = sl + SoLuongxoa;
                sql = "UPDATE tblsanpham SET SoLuong =" + slcon + " WHERE MaSP= N'" + MaHangxoa + "'";
                Functions.RunSQL(sql);
                ResetValues();
                LoadDataGridView();
            }
        }
    }
}
