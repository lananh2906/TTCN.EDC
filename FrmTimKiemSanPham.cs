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

namespace EDCZONE
{
    public partial class FrmTimKiemSanPham : Form
    {
        DataTable tblsanpham;
        public FrmTimKiemSanPham()
        {
            InitializeComponent();
        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void FrmTimKiemSanPham_Load(object sender, EventArgs e)
        {
            dgvTimKiemSanPham.DataSource = null;
        }
        private void LoadDataGridView()
        {

            dgvTimKiemSanPham.AllowUserToAddRows = false;
            dgvTimKiemSanPham.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void dgvTimKiemSanPham_DoubleClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn có muốn hiển thị thông tin chi tiết?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {

                txtMaSP.Text = dgvTimKiemSanPham.CurrentRow.Cells["MaSP"].Value.ToString();
                txtTenSP.Text = dgvTimKiemSanPham.CurrentRow.Cells["DonGiaNhap"].Value.ToString();
                txtMaNCC.Text = dgvTimKiemSanPham.CurrentRow.Cells["MaNCC"].Value.ToString();
                txtTenNCC.Text = dgvTimKiemSanPham.CurrentRow.Cells["TenNCC"].Value.ToString();
                txtTongTon.Text = dgvTimKiemSanPham.CurrentRow.Cells["TongTon"].Value.ToString();
                txtLoi.Text = dgvTimKiemSanPham.CurrentRow.Cells["Loi"].Value.ToString();
                txtCoTheBan.Text = dgvTimKiemSanPham.CurrentRow.Cells["CoTheBan"].Value.ToString();
                txtDonGiaBan.Text = dgvTimKiemSanPham.CurrentRow.Cells["DGB"].Value.ToString();
                txtKhuyenMai.Text = dgvTimKiemSanPham.CurrentRow.Cells["KhuyenMai"].Value.ToString();
                txtDanhMuc.Text = dgvTimKiemSanPham.CurrentRow.Cells["DanhMuc"].Value.ToString();

            }
        }
        private void ResetValues()
        {
            txtMaSP.Text = "";
            txtTenSP.Text = "";
            txtMaNCC.Text = "";
            txtTenNCC.Text = "";
            txtTongTon.Text = "";
            txtLoi.Text = "";
            txtCoTheBan.Text = "";
            txtDonGiaBan.Text = "";
            txtKhuyenMai.Text = "";
            txtDanhMuc.Text = "";
            txtMaSP.Focus();
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string sql;
            if ((txtMaSP.Text == "") && (txtTenSP.Text == "") && (txtMaNCC.Text == "") && (txtTenNCC.Text == ""))

            {
                MessageBox.Show("Hãy nhập một điều kiện tìm kiếm!!!", "Yêu cầu...", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            sql = "select *from tblsanpham where 1=1";
            if (txtMaSP.Text != "")
            {
                sql = sql + " AND MaSP Like '%" + txtMaSP.Text + "%'";
            }
            if (txtTenSP.Text != "")
            {
                sql = sql + " AND TenSP Like '%" + txtTenSP.Text + "%'";
            }
            if (txtMaNCC.Text != "")
            {
                sql = sql + " AND MaNCC Like '%" + txtMaNCC.Text + "%'";
            }
            if (txtTenNCC.Text != "")
            {
                sql = sql + " AND TenNCC Like '%" + txtTenNCC.Text + "%'";
            }

            tblsanpham = Functions.GetDataToTable(sql);
            if (tblsanpham.Rows.Count == 0)
            {
                MessageBox.Show("Không có bản ghi nào thỏa mãn điều kiện!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Có " + tblsanpham.Rows.Count + " bản ghi thỏa mãn điều kiện!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dgvTimKiemSanPham.DataSource = tblsanpham;
            LoadDataGridView();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("bạn có chắc chắn muốn thoát chương trình không", "Hỏi Thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                this.Close();
        }
    }
}
