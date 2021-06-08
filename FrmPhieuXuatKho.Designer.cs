
namespace EDCZONE
{
    partial class FrmPhieuXuatKho
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnthoat = new System.Windows.Forms.Button();
            this.btnIn = new System.Windows.Forms.Button();
            this.btnxemhoadon = new System.Windows.Forms.Button();
            this.btnxacnhan = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.NhanVien = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btntimkiem = new System.Windows.Forms.Button();
            this.dtpNgayBan = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.dgv_phieuxuatkho = new System.Windows.Forms.DataGridView();
            this.MaDH = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MaKH = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MaVC = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_phieuxuatkho)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel1.Controls.Add(this.btnthoat);
            this.panel1.Controls.Add(this.btnIn);
            this.panel1.Controls.Add(this.btnxemhoadon);
            this.panel1.Controls.Add(this.btnxacnhan);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 493);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(905, 100);
            this.panel1.TabIndex = 0;
            // 
            // btnthoat
            // 
            this.btnthoat.Font = new System.Drawing.Font("Times New Roman", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnthoat.ForeColor = System.Drawing.Color.DarkBlue;
            this.btnthoat.Image = global::EDCZONE.Properties.Resources.s_cancel1;
            this.btnthoat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnthoat.Location = new System.Drawing.Point(676, 32);
            this.btnthoat.Name = "btnthoat";
            this.btnthoat.Size = new System.Drawing.Size(123, 37);
            this.btnthoat.TabIndex = 7;
            this.btnthoat.Text = "&Thoát";
            this.btnthoat.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnthoat.UseVisualStyleBackColor = true;
            this.btnthoat.Click += new System.EventHandler(this.btnthoat_Click);
            // 
            // btnIn
            // 
            this.btnIn.Font = new System.Drawing.Font("Times New Roman", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnIn.ForeColor = System.Drawing.Color.DarkBlue;
            this.btnIn.Image = global::EDCZONE.Properties.Resources.save;
            this.btnIn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnIn.Location = new System.Drawing.Point(500, 32);
            this.btnIn.Name = "btnIn";
            this.btnIn.Size = new System.Drawing.Size(157, 37);
            this.btnIn.TabIndex = 6;
            this.btnIn.Text = "&In Hóa Đơn";
            this.btnIn.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnIn.UseVisualStyleBackColor = true;
            this.btnIn.Click += new System.EventHandler(this.btnIn_Click);
            // 
            // btnxemhoadon
            // 
            this.btnxemhoadon.Font = new System.Drawing.Font("Times New Roman", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnxemhoadon.ForeColor = System.Drawing.Color.DarkBlue;
            this.btnxemhoadon.Image = global::EDCZONE.Properties.Resources.view_report;
            this.btnxemhoadon.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnxemhoadon.Location = new System.Drawing.Point(97, 32);
            this.btnxemhoadon.Name = "btnxemhoadon";
            this.btnxemhoadon.Size = new System.Drawing.Size(180, 37);
            this.btnxemhoadon.TabIndex = 5;
            this.btnxemhoadon.Text = "&Xem Hóa Đơn";
            this.btnxemhoadon.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnxemhoadon.UseVisualStyleBackColor = true;
            // 
            // btnxacnhan
            // 
            this.btnxacnhan.Font = new System.Drawing.Font("Times New Roman", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnxacnhan.ForeColor = System.Drawing.Color.DarkBlue;
            this.btnxacnhan.Image = global::EDCZONE.Properties.Resources.okind_active;
            this.btnxacnhan.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnxacnhan.Location = new System.Drawing.Point(300, 32);
            this.btnxacnhan.Name = "btnxacnhan";
            this.btnxacnhan.Size = new System.Drawing.Size(176, 37);
            this.btnxacnhan.TabIndex = 4;
            this.btnxacnhan.Text = "&Xác Nhận HĐ";
            this.btnxacnhan.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnxacnhan.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.panel2.Controls.Add(this.NhanVien);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(905, 85);
            this.panel2.TabIndex = 1;
            // 
            // NhanVien
            // 
            this.NhanVien.AutoSize = true;
            this.NhanVien.Font = new System.Drawing.Font("Times New Roman", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NhanVien.ForeColor = System.Drawing.Color.DarkBlue;
            this.NhanVien.Location = new System.Drawing.Point(294, 32);
            this.NhanVien.Name = "NhanVien";
            this.NhanVien.Size = new System.Drawing.Size(238, 36);
            this.NhanVien.TabIndex = 27;
            this.NhanVien.Text = "Phiếu Xuất Kho";
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.groupBox1.Controls.Add(this.btntimkiem);
            this.groupBox1.Controls.Add(this.dtpNgayBan);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.ForeColor = System.Drawing.Color.DarkBlue;
            this.groupBox1.Location = new System.Drawing.Point(56, 107);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(756, 108);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Thực hiện";
            // 
            // btntimkiem
            // 
            this.btntimkiem.Font = new System.Drawing.Font("Times New Roman", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btntimkiem.ForeColor = System.Drawing.Color.DarkBlue;
            this.btntimkiem.Image = global::EDCZONE.Properties.Resources.b_search;
            this.btntimkiem.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btntimkiem.Location = new System.Drawing.Point(522, 45);
            this.btntimkiem.Name = "btntimkiem";
            this.btntimkiem.Size = new System.Drawing.Size(121, 33);
            this.btntimkiem.TabIndex = 19;
            this.btntimkiem.Text = "&Tìm Kiếm";
            this.btntimkiem.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btntimkiem.UseVisualStyleBackColor = true;
            this.btntimkiem.Click += new System.EventHandler(this.btntimkiem_Click);
            // 
            // dtpNgayBan
            // 
            this.dtpNgayBan.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpNgayBan.Location = new System.Drawing.Point(244, 45);
            this.dtpNgayBan.Name = "dtpNgayBan";
            this.dtpNgayBan.Size = new System.Drawing.Size(236, 28);
            this.dtpNgayBan.TabIndex = 18;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.DarkBlue;
            this.label5.Location = new System.Drawing.Point(101, 45);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(120, 25);
            this.label5.TabIndex = 17;
            this.label5.Text = "Chọn ngày:";
            // 
            // dgv_phieuxuatkho
            // 
            this.dgv_phieuxuatkho.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgv_phieuxuatkho.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_phieuxuatkho.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.MaDH,
            this.MaKH,
            this.MaVC});
            this.dgv_phieuxuatkho.Location = new System.Drawing.Point(56, 252);
            this.dgv_phieuxuatkho.Name = "dgv_phieuxuatkho";
            this.dgv_phieuxuatkho.RowHeadersWidth = 62;
            this.dgv_phieuxuatkho.RowTemplate.Height = 28;
            this.dgv_phieuxuatkho.Size = new System.Drawing.Size(756, 197);
            this.dgv_phieuxuatkho.TabIndex = 3;
            // 
            // MaDH
            // 
            this.MaDH.HeaderText = "Mã ĐH";
            this.MaDH.MinimumWidth = 8;
            this.MaDH.Name = "MaDH";
            // 
            // MaKH
            // 
            this.MaKH.HeaderText = "Mã KH";
            this.MaKH.MinimumWidth = 8;
            this.MaKH.Name = "MaKH";
            // 
            // MaVC
            // 
            this.MaVC.HeaderText = "MaVC";
            this.MaVC.MinimumWidth = 8;
            this.MaVC.Name = "MaVC";
            // 
            // FrmPhieuXuatKho
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(905, 593);
            this.Controls.Add(this.dgv_phieuxuatkho);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "FrmPhieuXuatKho";
            this.Text = "FrmPhieuXuatKho";
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_phieuxuatkho)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView dgv_phieuxuatkho;
        private System.Windows.Forms.Label NhanVien;
        private System.Windows.Forms.DateTimePicker dtpNgayBan;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnthoat;
        private System.Windows.Forms.Button btnIn;
        private System.Windows.Forms.Button btnxemhoadon;
        private System.Windows.Forms.Button btnxacnhan;
        private System.Windows.Forms.Button btntimkiem;
        private System.Windows.Forms.DataGridViewTextBoxColumn MaDH;
        private System.Windows.Forms.DataGridViewTextBoxColumn MaKH;
        private System.Windows.Forms.DataGridViewTextBoxColumn MaVC;
    }
}