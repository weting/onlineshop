
namespace onlineshop
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtbx_inputfile = new System.Windows.Forms.TextBox();
            this.btn_fileview = new System.Windows.Forms.Button();
            this.btn_website = new System.Windows.Forms.Button();
            this.txtbx_web = new System.Windows.Forms.TextBox();
            this.txtbx_data = new System.Windows.Forms.TextBox();
            this.btn_analyze = new System.Windows.Forms.Button();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.btn_updateexcel = new System.Windows.Forms.Button();
            this.btn_filecheck = new System.Windows.Forms.Button();
            this.btn_updateproduct = new System.Windows.Forms.Button();
            this.btn_publishproduct = new System.Windows.Forms.Button();
            this.btn_fileviewstock = new System.Windows.Forms.Button();
            this.btn_filecheckstock = new System.Windows.Forms.Button();
            this.txtbx_inputfilestock = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_getpicture = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "追蹤商品";
            // 
            // txtbx_inputfile
            // 
            this.txtbx_inputfile.Location = new System.Drawing.Point(80, 12);
            this.txtbx_inputfile.Name = "txtbx_inputfile";
            this.txtbx_inputfile.Size = new System.Drawing.Size(237, 22);
            this.txtbx_inputfile.TabIndex = 1;
            // 
            // btn_fileview
            // 
            this.btn_fileview.Location = new System.Drawing.Point(323, 10);
            this.btn_fileview.Name = "btn_fileview";
            this.btn_fileview.Size = new System.Drawing.Size(75, 23);
            this.btn_fileview.TabIndex = 2;
            this.btn_fileview.Text = "瀏覽";
            this.btn_fileview.UseVisualStyleBackColor = true;
            this.btn_fileview.Click += new System.EventHandler(this.btn_fileview_Click);
            // 
            // btn_website
            // 
            this.btn_website.Location = new System.Drawing.Point(485, 9);
            this.btn_website.Name = "btn_website";
            this.btn_website.Size = new System.Drawing.Size(87, 24);
            this.btn_website.TabIndex = 3;
            this.btn_website.Text = "網站";
            this.btn_website.UseVisualStyleBackColor = true;
            this.btn_website.Click += new System.EventHandler(this.btn_webview);
            // 
            // txtbx_web
            // 
            this.txtbx_web.Location = new System.Drawing.Point(731, 417);
            this.txtbx_web.Multiline = true;
            this.txtbx_web.Name = "txtbx_web";
            this.txtbx_web.Size = new System.Drawing.Size(435, 176);
            this.txtbx_web.TabIndex = 4;
            // 
            // txtbx_data
            // 
            this.txtbx_data.Location = new System.Drawing.Point(731, 248);
            this.txtbx_data.Multiline = true;
            this.txtbx_data.Name = "txtbx_data";
            this.txtbx_data.Size = new System.Drawing.Size(435, 163);
            this.txtbx_data.TabIndex = 5;
            // 
            // btn_analyze
            // 
            this.btn_analyze.Location = new System.Drawing.Point(578, 9);
            this.btn_analyze.Name = "btn_analyze";
            this.btn_analyze.Size = new System.Drawing.Size(69, 24);
            this.btn_analyze.TabIndex = 6;
            this.btn_analyze.Text = "分析";
            this.btn_analyze.UseVisualStyleBackColor = true;
            this.btn_analyze.Click += new System.EventHandler(this.btn_analyze_Click);
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(25, 73);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(690, 520);
            this.webBrowser1.TabIndex = 7;
            this.webBrowser1.Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.webBrowser1_Navigated);
            // 
            // btn_updateexcel
            // 
            this.btn_updateexcel.Location = new System.Drawing.Point(732, 9);
            this.btn_updateexcel.Name = "btn_updateexcel";
            this.btn_updateexcel.Size = new System.Drawing.Size(161, 53);
            this.btn_updateexcel.TabIndex = 8;
            this.btn_updateexcel.Text = "更新追蹤商品excel";
            this.btn_updateexcel.UseVisualStyleBackColor = true;
            this.btn_updateexcel.Click += new System.EventHandler(this.btn_updateexcel_Click);
            // 
            // btn_filecheck
            // 
            this.btn_filecheck.Location = new System.Drawing.Point(404, 10);
            this.btn_filecheck.Name = "btn_filecheck";
            this.btn_filecheck.Size = new System.Drawing.Size(75, 23);
            this.btn_filecheck.TabIndex = 9;
            this.btn_filecheck.Text = "確認";
            this.btn_filecheck.UseVisualStyleBackColor = true;
            this.btn_filecheck.Click += new System.EventHandler(this.btn_filecheck_Click);
            // 
            // btn_updateproduct
            // 
            this.btn_updateproduct.Location = new System.Drawing.Point(732, 73);
            this.btn_updateproduct.Name = "btn_updateproduct";
            this.btn_updateproduct.Size = new System.Drawing.Size(161, 53);
            this.btn_updateproduct.TabIndex = 10;
            this.btn_updateproduct.Text = "更新商品到露天";
            this.btn_updateproduct.UseVisualStyleBackColor = true;
            this.btn_updateproduct.Click += new System.EventHandler(this.btn_updateproduct_Click);
            // 
            // btn_publishproduct
            // 
            this.btn_publishproduct.Location = new System.Drawing.Point(731, 142);
            this.btn_publishproduct.Name = "btn_publishproduct";
            this.btn_publishproduct.Size = new System.Drawing.Size(162, 56);
            this.btn_publishproduct.TabIndex = 11;
            this.btn_publishproduct.Text = "上架商品到露天";
            this.btn_publishproduct.UseVisualStyleBackColor = true;
            // 
            // btn_fileviewstock
            // 
            this.btn_fileviewstock.Location = new System.Drawing.Point(323, 39);
            this.btn_fileviewstock.Name = "btn_fileviewstock";
            this.btn_fileviewstock.Size = new System.Drawing.Size(75, 23);
            this.btn_fileviewstock.TabIndex = 12;
            this.btn_fileviewstock.Text = "瀏覽";
            this.btn_fileviewstock.UseVisualStyleBackColor = true;
            this.btn_fileviewstock.Click += new System.EventHandler(this.btn_fileviewstock_Click);
            // 
            // btn_filecheckstock
            // 
            this.btn_filecheckstock.Location = new System.Drawing.Point(404, 39);
            this.btn_filecheckstock.Name = "btn_filecheckstock";
            this.btn_filecheckstock.Size = new System.Drawing.Size(75, 23);
            this.btn_filecheckstock.TabIndex = 13;
            this.btn_filecheckstock.Text = "確認";
            this.btn_filecheckstock.UseVisualStyleBackColor = true;
            this.btn_filecheckstock.Click += new System.EventHandler(this.btn_filecheckstock_Click);
            // 
            // txtbx_inputfilestock
            // 
            this.txtbx_inputfilestock.Location = new System.Drawing.Point(80, 39);
            this.txtbx_inputfilestock.Name = "txtbx_inputfilestock";
            this.txtbx_inputfilestock.Size = new System.Drawing.Size(237, 22);
            this.txtbx_inputfilestock.TabIndex = 14;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 15;
            this.label2.Text = "庫存商品";
            // 
            // btn_getpicture
            // 
            this.btn_getpicture.Location = new System.Drawing.Point(485, 38);
            this.btn_getpicture.Name = "btn_getpicture";
            this.btn_getpicture.Size = new System.Drawing.Size(87, 25);
            this.btn_getpicture.TabIndex = 16;
            this.btn_getpicture.Text = "獲取圖片";
            this.btn_getpicture.UseVisualStyleBackColor = true;
            this.btn_getpicture.Click += new System.EventHandler(this.btn_getpicture_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1177, 671);
            this.Controls.Add(this.btn_getpicture);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtbx_inputfilestock);
            this.Controls.Add(this.btn_filecheckstock);
            this.Controls.Add(this.btn_fileviewstock);
            this.Controls.Add(this.btn_publishproduct);
            this.Controls.Add(this.btn_updateproduct);
            this.Controls.Add(this.btn_filecheck);
            this.Controls.Add(this.btn_updateexcel);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.btn_analyze);
            this.Controls.Add(this.txtbx_data);
            this.Controls.Add(this.txtbx_web);
            this.Controls.Add(this.btn_website);
            this.Controls.Add(this.btn_fileview);
            this.Controls.Add(this.txtbx_inputfile);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtbx_inputfile;
        private System.Windows.Forms.Button btn_fileview;
        private System.Windows.Forms.Button btn_website;
        private System.Windows.Forms.TextBox txtbx_web;
        private System.Windows.Forms.TextBox txtbx_data;
        private System.Windows.Forms.Button btn_analyze;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Button btn_updateexcel;
        private System.Windows.Forms.Button btn_filecheck;
        private System.Windows.Forms.Button btn_updateproduct;
        private System.Windows.Forms.Button btn_publishproduct;
        private System.Windows.Forms.Button btn_fileviewstock;
        private System.Windows.Forms.Button btn_filecheckstock;
        private System.Windows.Forms.TextBox txtbx_inputfilestock;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_getpicture;
    }
}

