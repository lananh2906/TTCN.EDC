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
    public partial class FrmKhachHang : Form
    {
        DataTable tblkhachhang;
        public FrmKhachHang()
        {
            InitializeComponent();
        }

        private void FrmKhachHang_Load(object sender, EventArgs e)
        {
            LoadDataGridView();
        }
        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT * FROM tblkhachhang";
            tblkhachhang = Functions.GetDataToTable(sql);

            //Đọc dữ liệu từ bảng
            dgvkhachhang.DataSource = tblkhachhang; //Nguồn dữ liệu            
            dgvkhachhang.AllowUserToAddRows = false; //Không cho người dùng thêm dữ liệu trực tiếp
            dgvkhachhang.EditMode = DataGridViewEditMode.EditProgrammatically; //Không cho sửa dữ liệu trực tiếp
        }


        private void dgvkhachhang_Click(object sender, EventArgs e)
        {
            txtMaKH.Text = dgvkhachhang.CurrentRow.Cells["MaKH"].Value.ToString();
            txtTenKH.Text = dgvkhachhang.CurrentRow.Cells["TenKH"].Value.ToString();
            txtDiaChi.Text = dgvkhachhang.CurrentRow.Cells["DiaChi"].Value.ToString();
            txtSdt.Text = dgvkhachhang.CurrentRow.Cells["SDT"].Value.ToString();
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnHuy.Enabled = true;
            btnLuu.Enabled = true;
            btnThem.Enabled = false;

            txtMaKH.Enabled = true;
            txtMaKH.Focus();
            ResetValues();
        }
        private void ResetValues()
        {
            txtMaKH.Text = "";
            txtTenKH.Text = "";
            txtDiaChi.Text = "";
            txtSdt.Text = "";

        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblkhachhang.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaKH.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtTenKH.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên khách hàng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenKH.Focus();
                return;
            }
            if (txtDiaChi.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập địa chỉ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtDiaChi.Focus();
                return;
            }
            if (txtSdt.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập số điện thoại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSdt.Focus();
                return;
            }

            sql = "UPDATE tblkhachhang SET MaKH='" + txtMaKH.Text +
                "',TenKH='" + txtTenKH.Text +
                 "',DiaChi='" + txtDiaChi.Text +
                 "',SDT='" + txtSdt.Text +
                 "' WHERE MaKH='" + txtMaKH.Text + "'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnHuy.Enabled = false;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblkhachhang.Rows.Count == 0)
            {
                MessageBox.Show("không có dữ liệu");
            }
            if (txtMaKH.Text == "")
            {
                MessageBox.Show("bạn chưa chọn mã khách hàng");
            }
            if (MessageBox.Show("bạn muốn xóa không", "thông báo",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                sql = "delete from tblkhachhang where MaKH ='" + txtMaKH.Text + "'";
                Functions.RunSQL(sql);
                LoadDataGridView();
                ResetValues();
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblkhachhang.Rows.Count == 0)
            {
                MessageBox.Show("không có dữ liệu");
                return;
            }
            if (txtMaKH.Text == "")
            {
                MessageBox.Show("nhập mã khách hàng");
                txtMaKH.Focus();

            }
            if (txtTenKH.Text == "")
            {
                MessageBox.Show("nhập tên khách hàng");
                txtTenKH.Focus();
            }
            if (txtDiaChi.Text == "")
            {
                MessageBox.Show("nhập địa chỉ");
                txtDiaChi.Focus();

            }
            if (txtSdt.Text == "")
            {
                MessageBox.Show("nhập số điện thoại");
                txtSdt.Focus();

            }


            sql = "select MaKH from tblkhachhang where MaKH ='" + txtMaKH.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("mã này đã có, bạn phải nhập mã khác");
                txtMaKH.Focus();
                return;
            }
            sql = "INSERT INTO tblkhachhang(MaKH,TenKH,DiaChi,SDT) " +
                "VALUES " + "('" + txtMaKH.Text.Trim() +
                        "','" + txtTenKH.Text.Trim() +
                        "','" + txtDiaChi.Text +
                        "','" + txtSdt.Text.Trim() + "')";

            btnXoa.Enabled = true;
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnHuy.Enabled = false;
            btnLuu.Enabled = false;
            txtMaKH.Enabled = false;
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            ResetValues();
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
            btnThem.Enabled = true;
            btnLuu.Enabled = false;
            btnHuy.Enabled = false;
            txtMaKH.Enabled = false;
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("bạn có chắc chắn muốn thoát chương trình không", "Hỏi Thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                this.Close();
        }
    }
}
