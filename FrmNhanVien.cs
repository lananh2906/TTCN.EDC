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
    public partial class FrmNhanVien : Form
    {
        DataTable tblnhanvien;
        public FrmNhanVien()
        {
            InitializeComponent();
        }

        private void FrmNhanVien_Load(object sender, EventArgs e)
        {
            LoadDataGridView();
        }
        private void LoadDataGridView()
        {
            string sql;
            sql = "SELECT * FROM tblnhanhvien";
            tblnhanvien = Functions.GetDataToTable(sql);  ; //Đọc dữ liệu từ bảng
            dgvNhanVien.DataSource = tblnhanvien; //Nguồn dữ liệu            
            


            dgvNhanVien.AllowUserToAddRows = false; //Không cho người dùng thêm dữ liệu trực tiếp
            dgvNhanVien.EditMode = DataGridViewEditMode.EditProgrammatically; //Không cho sửa dữ liệu trực tiếp
        }

        private void dgvNhanVien_Click(object sender, EventArgs e)
        {
            txtMaNV.Text = dgvNhanVien.CurrentRow.Cells["MaNV"].Value.ToString();
            txtTenNV.Text = dgvNhanVien.CurrentRow.Cells["HoTen"].Value.ToString();
            txtNamSinh.Text = dgvNhanVien.CurrentRow.Cells["NamSinh"].Value.ToString();
            txtQueQuan.Text = dgvNhanVien.CurrentRow.Cells["QueQuan"].Value.ToString();
            txtGioiTinh.Text = dgvNhanVien.CurrentRow.Cells["GioiTinh"].Value.ToString();
            txtEmail.Text = dgvNhanVien.CurrentRow.Cells["Email"].Value.ToString();
            txtSdt.Text = dgvNhanVien.CurrentRow.Cells["SDT"].Value.ToString();
            txtMaCV.Text = dgvNhanVien.CurrentRow.Cells["MaCV"].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            btnSua.Enabled = false;
            btnXoa.Enabled = false;
            btnHuy.Enabled = true;
            btnLuu.Enabled = true;
            btnThem.Enabled = false;

            txtMaNV.Enabled = true;
            txtMaNV.Focus();
            ResetValues();
        }
        private void ResetValues()
        {
            txtMaNV.Text = "";
            txtTenNV.Text = "";
            txtNamSinh.Text = "";
            txtQueQuan.Text = "";
            txtGioiTinh.Text = "";
            txtSdt.Text = "";
            txtEmail.Text = "";
            txtMaCV.Text = "";

            //txtDonGiaNhap.Enabled = false;
            // txtDonGiaBan.Enabled = false;


        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            string sql;

            if (tblnhanvien.Rows.Count == 0)
            {
                MessageBox.Show("Không còn dữ liệu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtMaNV.Text == "")
            {
                MessageBox.Show("Bạn chưa chọn bản ghi nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (txtTenNV.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập tên nhân viên", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenNV.Focus();
                return;
            }
            if (txtNamSinh.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập năm sinh", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtNamSinh.Focus();
                return;
            }
            if (txtQueQuan.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập quê quán", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtQueQuan.Focus();
                return;
            }
            if (txtGioiTinh.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập giới tính", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtGioiTinh.Focus();
                return;
            }
            if (txtSdt.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập số điện thoại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtSdt.Focus();
                return;
            }
            if (txtEmail.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập email", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtEmail.Focus();
                return;
            }
            if (txtMaCV.Text.Trim().Length == 0)
            {
                MessageBox.Show("Bạn phải nhập mã chức vụ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMaCV.Focus();
                return;
            }
            sql = "UPDATE tblnhanhvien SET MaNV='" + txtMaNV.Text +
                "',HoTen='" + txtTenNV.Text +
                "',NamSinh='" + txtNamSinh.Text +
                "',QueQuan='" + txtQueQuan.Text +
                "',GioiTinh='" + txtGioiTinh.Text +
                "',SDT='" + txtSdt.Text +
                "',Email='" + txtEmail.Text +
                "',MaCV='" + txtMaCV.Text + "' WHERE MaNV='" + txtMaNV.Text + "'";
            Functions.RunSQL(sql);
            LoadDataGridView();
            ResetValues();
            btnHuy.Enabled = false;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblnhanvien.Rows.Count == 0)
            {
                MessageBox.Show("không có dữ liệu");
            }
            if (txtMaNV.Text == "")
            {
                MessageBox.Show("bạn chưa chọn mã nhân viên");
            }
            if (MessageBox.Show("Bạn có muốn xoá bản ghi này không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {

                sql = "DELETE tblnhanhvien WHERE MaNV='" + txtMaNV.Text + "'";
                Functions.RunSQL(sql);
                LoadDataGridView();
                ResetValues();
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            string sql;
            if (tblnhanvien.Rows.Count == 0)
            {
                MessageBox.Show("không có dữ liệu");
                return;
            }
            if (txtMaNV.Text == "")
            {
                MessageBox.Show("nhập mã nhân viên");
                txtMaNV.Focus();

            }
            if (txtTenNV.Text == "")
            {
                MessageBox.Show("nhập tên nhân viên");
                txtTenNV.Focus();
            }
            if (txtNamSinh.Text == "")
            {
                MessageBox.Show("nhập năm sinh");
                txtNamSinh.Focus();

            }
            if (txtQueQuan.Text == "")
            {
                MessageBox.Show("nhập quê quán");
                txtQueQuan.Focus();

            }
            if (txtGioiTinh.Text == "")
            {
                MessageBox.Show("nhập giới tính");
                txtGioiTinh.Focus();

            }
            if (txtSdt.Text == "")
            {
                MessageBox.Show("nhập số điện thoại");
                txtSdt.Focus();

            }
            if (txtEmail.Text == "")
            {
                MessageBox.Show("nhập email");
                txtEmail.Focus();

            }
            if (txtMaCV.Text == "")
            {
                MessageBox.Show("nhập mã chức vụ");
                txtMaCV.Focus();

            }
            sql = "select MaNV from tblnhanhvien where MaNV ='" + txtMaNV.Text.Trim() + "'";
            if (Functions.CheckKey(sql))
            {
                MessageBox.Show("mã này đã có, bạn phải nhập mã khác");
                txtMaNV.Focus();
                return;
            }
            sql = "INSERT INTO tblnhanhvien(MaNV,HoTen,NamSinh,QueQuan,GioiTinh,Email,SDT,MaCV) VALUES ('" + txtMaNV.Text.Trim() + "','" + txtTenNV.Text.Trim() + "','" + txtNamSinh.Text + "','" + txtQueQuan.Text + "','" + txtGioiTinh.Text + "','" + txtEmail.Text + "','" + txtSdt.Text + "','" + txtMaCV.Text.Trim() + "')";

            btnXoa.Enabled = true;
            btnThem.Enabled = true;
            btnSua.Enabled = true;
            btnHuy.Enabled = false;
            btnLuu.Enabled = false;
            txtMaNV.Enabled = false;
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
            txtMaNV.Enabled = false;
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("bạn có chắc chắn muốn thoát chương trình không", "Hỏi Thoát", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                this.Close();
        }
    }
}
