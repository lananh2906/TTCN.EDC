namespace EDCZONE
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.mnudanhmuc = new System.Windows.Forms.ToolStripMenuItem();
            this.mnunhanvien = new System.Windows.Forms.ToolStripMenuItem();
            this.mnukhachhang = new System.Windows.Forms.ToolStripMenuItem();
            this.mnunhacc = new System.Windows.Forms.ToolStripMenuItem();
            this.mnusanpham = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuhoadon = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuhoadonban = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuhoadonnhap = new System.Windows.Forms.ToolStripMenuItem();
            this.mnutimkiem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnutimkiemsp = new System.Windows.Forms.ToolStripMenuItem();
            this.mnukho = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuhangtonkho = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuphieunhapkho = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuphieuxuatkho = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuthoat = new System.Windows.Forms.ToolStripMenuItem();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.GripMargin = new System.Windows.Forms.Padding(2, 2, 0, 2);
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnudanhmuc,
            this.mnuhoadon,
            this.mnutimkiem,
            this.mnukho,
            this.mnuthoat});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(12, 3, 0, 3);
            this.menuStrip1.Size = new System.Drawing.Size(761, 35);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // mnudanhmuc
            // 
            this.mnudanhmuc.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnunhanvien,
            this.mnukhachhang,
            this.mnunhacc,
            this.mnusanpham});
            this.mnudanhmuc.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mnudanhmuc.Image = global::EDCZONE.Properties.Resources.edit_icon;
            this.mnudanhmuc.Name = "mnudanhmuc";
            this.mnudanhmuc.Size = new System.Drawing.Size(151, 29);
            this.mnudanhmuc.Text = "&Danh Mục";
            // 
            // mnunhanvien
            // 
            this.mnunhanvien.Image = ((System.Drawing.Image)(resources.GetObject("mnunhanvien.Image")));
            this.mnunhanvien.Name = "mnunhanvien";
            this.mnunhanvien.Size = new System.Drawing.Size(270, 34);
            this.mnunhanvien.Text = "&Nhân Viên";
            this.mnunhanvien.Click += new System.EventHandler(this.mnunhanvien_Click);
            // 
            // mnukhachhang
            // 
            this.mnukhachhang.Image = ((System.Drawing.Image)(resources.GetObject("mnukhachhang.Image")));
            this.mnukhachhang.Name = "mnukhachhang";
            this.mnukhachhang.Size = new System.Drawing.Size(270, 34);
            this.mnukhachhang.Text = "&Khách Hàng";
            this.mnukhachhang.Click += new System.EventHandler(this.mnukhachhang_Click);
            // 
            // mnunhacc
            // 
            this.mnunhacc.Image = ((System.Drawing.Image)(resources.GetObject("mnunhacc.Image")));
            this.mnunhacc.Name = "mnunhacc";
            this.mnunhacc.Size = new System.Drawing.Size(270, 34);
            this.mnunhacc.Text = "&Nhà CC";
            this.mnunhacc.Click += new System.EventHandler(this.mnunhacc_Click);
            // 
            // mnusanpham
            // 
            this.mnusanpham.Image = ((System.Drawing.Image)(resources.GetObject("mnusanpham.Image")));
            this.mnusanpham.Name = "mnusanpham";
            this.mnusanpham.Size = new System.Drawing.Size(270, 34);
            this.mnusanpham.Text = "&Sản Phẩm";
            this.mnusanpham.Click += new System.EventHandler(this.mnusanpham_Click);
            // 
            // mnuhoadon
            // 
            this.mnuhoadon.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuhoadonban,
            this.mnuhoadonnhap});
            this.mnuhoadon.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mnuhoadon.Image = global::EDCZONE.Properties.Resources.reports_icon1;
            this.mnuhoadon.Name = "mnuhoadon";
            this.mnuhoadon.Size = new System.Drawing.Size(137, 29);
            this.mnuhoadon.Text = "Hóa Đơn";
            // 
            // mnuhoadonban
            // 
            this.mnuhoadonban.Name = "mnuhoadonban";
            this.mnuhoadonban.Size = new System.Drawing.Size(256, 34);
            this.mnuhoadonban.Text = "Hóa Đơn Bán";
            this.mnuhoadonban.Click += new System.EventHandler(this.mnuhoadonban_Click);
            // 
            // mnuhoadonnhap
            // 
            this.mnuhoadonnhap.Name = "mnuhoadonnhap";
            this.mnuhoadonnhap.Size = new System.Drawing.Size(256, 34);
            this.mnuhoadonnhap.Text = "Hóa Đơn Nhập";
            this.mnuhoadonnhap.Click += new System.EventHandler(this.mnuhoadonnhap_Click);
            // 
            // mnutimkiem
            // 
            this.mnutimkiem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnutimkiemsp});
            this.mnutimkiem.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mnutimkiem.Image = global::EDCZONE.Properties.Resources.b_search;
            this.mnutimkiem.Name = "mnutimkiem";
            this.mnutimkiem.Size = new System.Drawing.Size(149, 29);
            this.mnutimkiem.Text = "Tìm Kiếm";
            // 
            // mnutimkiemsp
            // 
            this.mnutimkiemsp.Name = "mnutimkiemsp";
            this.mnutimkiemsp.Size = new System.Drawing.Size(313, 34);
            this.mnutimkiemsp.Text = "Tìm Kiếm Sản Phẩm";
            this.mnutimkiemsp.Click += new System.EventHandler(this.mnutimkiemsp_Click);
            // 
            // mnukho
            // 
            this.mnukho.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuhangtonkho,
            this.mnuphieunhapkho,
            this.mnuphieuxuatkho});
            this.mnukho.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mnukho.Image = global::EDCZONE.Properties.Resources.user_green_locked_exp;
            this.mnukho.Name = "mnukho";
            this.mnukho.Size = new System.Drawing.Size(93, 29);
            this.mnukho.Text = "Kho";
            // 
            // mnuhangtonkho
            // 
            this.mnuhangtonkho.Name = "mnuhangtonkho";
            this.mnuhangtonkho.Size = new System.Drawing.Size(272, 34);
            this.mnuhangtonkho.Text = "Hàng Tồn Kho";
            this.mnuhangtonkho.Click += new System.EventHandler(this.mnuhangtonkho_Click);
            // 
            // mnuphieunhapkho
            // 
            this.mnuphieunhapkho.Name = "mnuphieunhapkho";
            this.mnuphieunhapkho.Size = new System.Drawing.Size(272, 34);
            this.mnuphieunhapkho.Text = "Phiếu Nhập Kho";
            this.mnuphieunhapkho.Click += new System.EventHandler(this.mnuphieunhapkho_Click);
            // 
            // mnuphieuxuatkho
            // 
            this.mnuphieuxuatkho.Name = "mnuphieuxuatkho";
            this.mnuphieuxuatkho.Size = new System.Drawing.Size(272, 34);
            this.mnuphieuxuatkho.Text = "Phiếu Xuất Kho";
            this.mnuphieuxuatkho.Click += new System.EventHandler(this.mnuphieuxuatkho_Click);
            // 
            // mnuthoat
            // 
            this.mnuthoat.Font = new System.Drawing.Font("Times New Roman", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mnuthoat.Image = global::EDCZONE.Properties.Resources.s_cancel1;
            this.mnuthoat.Name = "mnuthoat";
            this.mnuthoat.Size = new System.Drawing.Size(108, 29);
            this.mnuthoat.Text = "Thoát";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = global::EDCZONE.Properties.Resources.ctyedczone;
            this.pictureBox1.Location = new System.Drawing.Point(0, 41);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(1024, 707);
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 23F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::EDCZONE.Properties.Resources.z2460532493241_002ebe5564bdc0a9fdfb7450f17be33a;
            this.ClientSize = new System.Drawing.Size(761, 661);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Times New Roman", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.DarkBlue;
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "frmMain";
            this.Text = "Công Ty Cổ Phần EDCzone";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem mnudanhmuc;
        private System.Windows.Forms.ToolStripMenuItem mnunhanvien;
        private System.Windows.Forms.ToolStripMenuItem mnukhachhang;
        private System.Windows.Forms.ToolStripMenuItem mnunhacc;
        private System.Windows.Forms.ToolStripMenuItem mnusanpham;
        private System.Windows.Forms.ToolStripMenuItem mnuhoadon;
        private System.Windows.Forms.ToolStripMenuItem mnuhoadonban;
        private System.Windows.Forms.ToolStripMenuItem mnuhoadonnhap;
        private System.Windows.Forms.ToolStripMenuItem mnutimkiem;
        private System.Windows.Forms.ToolStripMenuItem mnukho;
        private System.Windows.Forms.ToolStripMenuItem mnuhangtonkho;
        private System.Windows.Forms.ToolStripMenuItem mnuphieunhapkho;
        private System.Windows.Forms.ToolStripMenuItem mnuphieuxuatkho;
        private System.Windows.Forms.ToolStripMenuItem mnuthoat;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ToolStripMenuItem mnutimkiemsp;
    }
}

