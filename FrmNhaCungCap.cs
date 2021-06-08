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
    public partial class FrmNhaCungCap : Form
    {
        DataTable tblncc;
        public FrmNhaCungCap()
        {
            InitializeComponent();
        }

        private void FrmNhaCungCap_Load(object sender, EventArgs e)
        {
            LoadDataGridView();
        }
        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT MaNCC, TenNCC, DiaChi, SDT FROM tblncc";
            tblncc = Functions.GetDataToTable(sql);


            //Đọc dữ liệu từ bảng
            dgvnhacc.DataSource = tblncc; //Nguồn dữ liệu            
            dgvnhacc.Columns[0].HeaderText = "Mã NCC ";
            dgvnhacc.Columns[1].HeaderText = "Tên NCC";
            dgvnhacc.Columns[2].HeaderText = "Địa Chỉ";
            dgvnhacc.Columns[3].HeaderText = "SDT";


            dgvnhacc.AllowUserToAddRows = false; //Không cho người dùng thêm dữ liệu trực tiếp
            dgvnhacc.EditMode = DataGridViewEditMode.EditProgrammatically; //Không cho sửa dữ liệu trực tiếp
        }

        private void dgvnhacc_Click(object sender, EventArgs e)
        {
            txtMaNCC.Text = dgvnhacc.CurrentRow.Cells["MaNCC"].Value.ToString();
            txtTenNCC.Text = dgvnhacc.CurrentRow.Cells["TenNCC"].Value.ToString();
            txtDiaChi.Text = dgvnhacc.CurrentRow.Cells["DiaChi"].Value.ToString();
            txtSdt.Text = dgvnhacc.CurrentRow.Cells["SDT"].Value.ToString();
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
       
            btnLuu.Enabled = true;
            btnThem.Enabled = false;
            ResetValue(); //Xoá trắng các textbox
            txtMaNCC.Enabled = true;//cho phép nhập mới
            txtMaNCC.Focus();
        }
        private void ResetValue()
        {
            txtMaNCC.Text = "";
            txtTenNCC.Text = "";
            txtDiaChi.Text = "";
            txtSdt.Text = "";
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblncc.Rows.Count == 0)
            {
                MessageBox.Show("không có dữ liệu");
                return;
            }
            if (txtMaNCC.Text == "")
            {
                MessageBox.Show("nhập mã NCC");
                txtMaNCC.Focus();

            }
            if (txtTenNCC.Text == "")
            {
                MessageBox.Show("nhập tên NCC");
                txtTenNCC.Focus();
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


            sql = "select MaNCC from tblncc where MaNCC ='" + txtMaNCC.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("mã này đã có, bạn phải nhập mã khác");
                txtMaNCC.Focus();
                return;
            }
            sql = "INSERT INTO tblncc(MaNCC,TenNCC,DiaChi,SDT) " +
                 "VALUES " + "('" + txtMaNCC.Text.Trim() +
                         "','" + txtTenNCC.Text.Trim() +
                         "','" + txtDiaChi.Text +
                         "','" + txtSdt.Text.Trim() + "')";
            btnXoa.Enabled = true;
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnHuy.Enabled = false;
            btnLuu.Enabled = false;
            txtMaNCC.Enabled = false;
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValue();
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblncc.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaNCC.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtTenNCC.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên NCC", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenNCC.Focus();
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
            sql = "UPDATE tblncc SET MaNCC='" + txtMaNCC.Text +
                 "',TenNCC='" + txtTenNCC.Text +
                  "',DiaChi='" + txtDiaChi.Text +
                  "',SDT='" + txtSdt.Text +
                  "' WHERE MaNCC='" + txtMaNCC.Text + "'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValue();
            btnHuy.Enabled = false;


        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblncc.Rows.Count == 0)
            {
                MessageBox.Show("không có dữ liệu");
            }
            if (txtMaNCC.Text == "")
            {
                MessageBox.Show("bạn chưa chọn mã NCC");
            }
            if (MessageBox.Show("Bạn có muốn xoá không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                sql = "DELETE tblncc WHERE MaNCC=N'" + txtMaNCC.Text + "'";
                Functions.RunSQL(sql);
                LoadDataGridView();
                ResetValue();
            }
        }

        private void btnHuy_Click(object sender, EventArgs e)
        {
            ResetValue();
         
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
            btnThem.Enabled = true;
            btnLuu.Enabled = false;
            btnHuy.Enabled = false;
            txtMaNCC.Enabled = false;
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("bạn có chắc chắn muốn thoát chương trình không", "Hỏi Thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                this.Close();
        }
    }
}
