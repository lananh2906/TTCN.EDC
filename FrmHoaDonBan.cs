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
    public partial class FrmHoaDonBan : Form
    {
        DataTable tblcthdb;
        public FrmHoaDonBan()
        {
            InitializeComponent();
        }

        private void FrmHoaDonBan_Load(object sender, EventArgs e)
        {
            btnThemHĐ.Enabled = true;
            btnLuu.Enabled = false;
            btnXoaHĐ.Enabled = false;
            btnInHĐ.Enabled = false;

            txtSoHDB.ReadOnly = true;
            txtTenNV.ReadOnly = true;
            txtTenKH.ReadOnly = true;
            txtDiaChi.ReadOnly = true;
            txtSDT.ReadOnly = true;
            txtTenSP.ReadOnly = true;
            txtDonGia.ReadOnly = true;
            txtThanhTien.ReadOnly = true;
            txtTongTien.ReadOnly = true;
            txtGiamGia.Text = "0";
            txtTongTien.Text = "0";

            Functions.FillCombo1("select MaNV from tblnhanhvien", cboMaNV, "MaNV");
            cboMaNV.SelectedIndex = -1;
            Functions.FillCombo1("select MaKH from tblkhachhang", cboMaKhach, "MaKH");
            cboMaKhach.SelectedIndex = -1;
            Functions.FillCombo1("select MaSP from tblsanpham", cboMaSP, "MaSP");
            cboMaSP.SelectedIndex = -1;
            if (txtSoHDB.Text != "")
            {
                LoadInfoHoaDon();
                btnXoaHĐ.Enabled = true;
                btnHuy.Enabled = true;
                btnInHĐ.Enabled = true;
            }
            LoadDataGridView();

        }
        private void LoadDataGridView()
        {
            string sql;

            sql = "SELECT* FROM tblchitiethoadonban";
            tblcthdb = Functions.GetDataToTable(sql);
            dataGridView_HDB.DataSource = tblcthdb;
            dataGridView_HDB.AllowUserToAddRows = false;
            dataGridView_HDB.EditMode = DataGridViewEditMode.EditProgrammatically;
        }
        private void LoadInfoHoaDon()
        {
            string str;
            
         
            str = "SELECT NgayBan FROM tblhoadonban WHERE MaHDB = N'" + txtSoHDB.Text + "'";
            dtpNgayBan.Value = DateTime.Parse(Functions.GetFieldValues(str));
            str = "SELECT MaNV FROM tblhoadonban WHERE MaHDB = N'" + txtSoHDB.Text + "'";
            cboMaNV.SelectedValue = Functions.GetFieldValues(str);
            str = "SELECT MaKH FROM tblhoadonban WHERE MaHDB = N'" + txtSoHDB.Text + "'";
            cboMaKhach.SelectedValue = Functions.GetFieldValues(str);
            str = "SELECT TongTien FROM tblhoadonban WHERE SoHDB = N'" + txtSoHDB.Text + "'";
            txtTien.Text = Functions.GetFieldValues(str);
            lblChu.Text = "Bằng chữ: " + Functions.ChuyenSoSangChu(txtTien.Text);
        }

        private void btnThemHĐ_Click(object sender, EventArgs e)
        {
            btnXoaHĐ.Enabled = false;
            btnLuu.Enabled = true;
            btnHuy.Enabled = true;
            btnThemHĐ.Enabled = false;
            ResetValues();
            txtSoHDB.Text = Functions.CreateKey("HDB");
            LoadDataGridView();
        }
        private void ResetValues()
        {
            txtSoHDB.Text = "";
            dtpNgayBan.Text = DateTime.Now.ToShortDateString();
            cboMaNV.Text = "";
            cboMaKhach.Text = "";
            txtTien.Text = "0";
            cboMaSP.Text = "";
            txtSoLuong.Text = "";
            txtGiamGia.Text = "0";
            txtThanhTien.Text = "0";
            lblChu.Text = "Bằng chữ: ";
        }

        private void btnLuuHD_Click(object sender, EventArgs e)
        {
            string sql;
            Double sl,SLcon,tong,Tongmoi;
            sql = "SELECT MaHDB FROM tblhoadonban WHERE MaHDB='" + txtSoHDB.Text + "'";
            if (!Functions.CheckKey(sql))
            {
                // Mã hóa đơn chưa có, tiến hành lưu các thông tin chung
                // Mã HDBan được sinh tự động do đó không có trường hợp trùng khóa
              
                sql = "INSERT INTO tblhoadonban(MaHDB, NgayBan, MaNV, MaKH, TongTien) VALUES (N'" + txtSoHDB.Text.Trim() + "','" +
                        dtpNgayBan.Value + "',N'" + cboMaNV.SelectedValue + "',N'" +
                        cboMaKhach.SelectedValue + "'," + txtTien.Text + ")";
                Functions.RunSQL(sql);
            }
            // Lưu thông tin của các mặt hàng

            
            


            // Kiểm tra xem số lượng hàng trong kho còn đủ để cung cấp không?
            sl = Convert.ToDouble(Functions.GetFieldValues("SELECT SoLuong FROM tblsanpham WHERE MaSP = N'" + cboMaSP.SelectedValue + "'"));

            if (Convert.ToDouble(txtSoLuong.Text) > sl)
            {
                MessageBox.Show("Số lượng mặt hàng này chỉ còn " + sl, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtSoLuong.Text = "";
                txtSoLuong.Focus();
                return;
            }

            sql = "INSERT INTO tblchitiethoadonban(MaHDB,MaSP,DonGiaBan,SoLuong,GiamGia,ThanhTien) VALUES(N'" + txtSoHDB.Text.Trim() + "','" + cboMaSP.SelectedValue + "','" + txtSoLuong.Text + "','" + txtDonGia.Text + "','" + txtGiamGia.Text + "','" + txtThanhTien.Text + "')";
            Functions.RunSQL(sql);
            LoadDataGridView();
            // Cập nhật lại số lượng của mặt hàng vào bảng SanPham
            SLcon = sl - Convert.ToDouble(txtSoLuong.Text);
            sql = "UPDATE tblsanpham SET SoLuong =" + SLcon + " WHERE MaSP= N'" + cboMaSP.SelectedValue + "'";
            Functions.RunSQL(sql);
            // Cập nhật lại tổng tiền cho hóa đơn bán
            tong = Convert.ToDouble(Functions.GetFieldValues("SELECT TongTien FROM tblhoadonban WHERE MaHDB = N'" + txtSoHDB.Text + "'"));
            Tongmoi = tong + Convert.ToDouble(txtThanhTien.Text);
            sql = "UPDATE tblhoadonban SET TongTien =" + Tongmoi + " WHERE MaHDB = N'" + txtSoHDB.Text + "'";
            Functions.RunSQL(sql);
            txtTien.Text = Tongmoi.ToString();
            lblChu.Text = "Bằng chữ: " + Functions.ChuyenSoSangChuoi(Double.Parse(Tongmoi.ToString()));
            Functions.RunSQL(sql);
            ResetValuesHang();
            btnXoaHĐ.Enabled = true;
            btnThemHĐ.Enabled = true;
        }
        private void ResetValuesHang()
        {
            cboMaSP.Text = "";
            txtSoLuong.Text = "";
            txtGiamGia.Text = "0";
            txtThanhTien.Text = "0";
        }

        private void cboMaSP_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaSP.Text == "")
            {
                txtTenSP.Text = "";
                txtDonGia.Text = "";
                txtGiamGia.Text = "";
                txtSoLuong.Text = "";
                txtThanhTien.Text = "";
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

        private void cboMaKhach_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str;
            if (cboMaKhach.Text == "")
            {
                txtTenKH.Text = "";
                txtDiaChi.Text = "";
                txtSDT.Text = "";
            }
            //Khi chọn Mã khách hàng thì các thông tin của khách hàng sẽ hiện ra
            str = "Select TenKH from tblkhachhang where MaKH = N'" + cboMaKhach.SelectedValue + "'";
            txtTenKH.Text = Functions.GetFieldValues(str);
            str = "Select DiaChi from tblkhachhang where MaKH = N'" + cboMaKhach.SelectedValue + "'";
            txtDiaChi.Text = Functions.GetFieldValues(str);
            str = "Select SDT from tblkhachhang where MaKH= N'" + cboMaKhach.SelectedValue + "'";
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

        private void dataGridView_HDB_DoubleClick(object sender, EventArgs e)
        {
            string MaHangxoa, sql;
            Double ThanhTienxoa, SoLuongxoa, sl, slcon;
            if (tblcthdb.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if ((MessageBox.Show("Bạn có chắc chắn muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes))
            {
                //Xóa hàng và cập nhật lại số lượng hàng 
                MaHangxoa = dataGridView_HDB.CurrentRow.Cells["MaSP"].Value.ToString();
                SoLuongxoa = Convert.ToDouble(dataGridView_HDB.CurrentRow.Cells["SoLuong"].Value.ToString());
                ThanhTienxoa = Convert.ToDouble(dataGridView_HDB.CurrentRow.Cells["ThanhTien"].Value.ToString());
                sql = "DELETE tblchitiethoadonban WHERE MaHDB=N'" + txtSoHDB.Text + "' AND MaSP = N'" + MaHangxoa + "'";
                Functions.RunSQL(sql);
                // Cập nhật lại số lượng cho các mặt hàng
                sl = Convert.ToDouble(Functions.GetFieldValues("SELECT SoLuong FROM tblsanpham WHERE MaSP = N'" + MaHangxoa + "'"));
                slcon = sl + SoLuongxoa;
                sql = "UPDATE tblsanpham SET SoLuong =" + slcon + " WHERE MaSP= N'" + MaHangxoa + "'";
                Functions.RunSQL(sql);
                ResetValuesHang();
                LoadDataGridView();
            }
        }

        
    }
}
