using CsvHelper;
using FlexImport.Classes;
using FlexImport.Services;
using OfficeOpenXml;
using System.IO;

namespace FlexImport
{
    partial class Form1 : Form
    {
        private System.ComponentModel.IContainer Components = null;

        private System.Windows.Forms.Label LblSheet;
        private System.Windows.Forms.TextBox TxtSheet;
        private System.Windows.Forms.Button BtnSelectSheet;
        private System.Windows.Forms.Label LblPath;
        private System.Windows.Forms.TextBox TxtPath;
        private System.Windows.Forms.Button BtnSelectPath;
        private System.Windows.Forms.Button BtnConvert;
        private System.Windows.Forms.Label LblWarning;
        private System.Windows.Forms.Label LblCategory;
        private System.Windows.Forms.ComboBox CmbCategory;
        private System.Windows.Forms.Label LblSubCategory;
        private System.Windows.Forms.ComboBox CmbSubCategory;
        private System.Windows.Forms.Label LblGenre;
        private System.Windows.Forms.ComboBox CmbGenre;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (Components != null))
            {
                Components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            LblSheet = new Label();
            TxtSheet = new TextBox();
            BtnSelectSheet = new Button();
            LblPath = new Label();
            TxtPath = new TextBox();
            BtnSelectPath = new Button();
            BtnConvert = new Button();
            LblWarning = new Label();
            LblCategory = new Label();
            CmbCategory = new ComboBox();
            LblSubCategory = new Label();
            CmbSubCategory = new ComboBox();
            LblGenre = new Label();
            CmbGenre = new ComboBox();
            SuspendLayout();
            // 
            // LblSheet
            // 
            LblSheet.Location = new Point(20, 42);
            LblSheet.Name = "LblSheet";
            LblSheet.Size = new Size(250, 20);
            LblSheet.TabIndex = 1;
            LblSheet.Text = "Selecione a planilha exportada do Bling:";
            // 
            // TxtSheet
            // 
            TxtSheet.Location = new Point(20, 67);
            TxtSheet.Name = "TxtSheet";
            TxtSheet.Size = new Size(400, 23);
            TxtSheet.TabIndex = 2;
            // 
            // BtnSelectSheet
            // 
            BtnSelectSheet.Location = new Point(430, 67);
            BtnSelectSheet.Name = "BtnSelectSheet";
            BtnSelectSheet.Size = new Size(75, 23);
            BtnSelectSheet.TabIndex = 3;
            BtnSelectSheet.Text = "Selecionar";
            BtnSelectSheet.Click += BtnSelectSheet_Click;
            // 
            // LblPath
            // 
            LblPath.Location = new Point(20, 104);
            LblPath.Name = "LblPath";
            LblPath.Size = new Size(250, 20);
            LblPath.TabIndex = 9;
            LblPath.Text = "Selecione onde será salva a nova planilha:";
            // 
            // TxtPath
            // 
            TxtPath.Location = new Point(20, 127);
            TxtPath.Name = "TxtPath";
            TxtPath.Size = new Size(400, 23);
            TxtPath.TabIndex = 5;
            TxtPath.Text = "C:\\Users\\andre\\Downloads";
            // 
            // BtnSelectPath
            // 
            BtnSelectPath.Location = new Point(430, 127);
            BtnSelectPath.Name = "BtnSelectPath";
            BtnSelectPath.Size = new Size(75, 23);
            BtnSelectPath.TabIndex = 6;
            BtnSelectPath.Text = "Selecionar";
            BtnSelectPath.Click += BtnSelectPath_Click;
            // 
            // BtnConvert
            // 
            BtnConvert.Location = new Point(405, 315);
            BtnConvert.Name = "BtnConvert";
            BtnConvert.Size = new Size(100, 30);
            BtnConvert.TabIndex = 11;
            BtnConvert.Text = "Converter";
            BtnConvert.Click += BtnConvert_Click;
            // 
            // LblWarning
            // 
            LblWarning.ForeColor = Color.Red;
            LblWarning.Location = new Point(20, 10);
            LblWarning.Name = "LblWarning";
            LblWarning.Size = new Size(400, 20);
            LblWarning.TabIndex = 0;
            LblWarning.Text = "A exportação do bling deve ser feita em CSV.";
            // 
            // LblCategory
            // 
            LblCategory.Location = new Point(20, 168);
            LblCategory.Name = "LblCategory";
            LblCategory.Size = new Size(200, 23);
            LblCategory.TabIndex = 0;
            LblCategory.Text = "Selecione a categoria:";
            // 
            // CmbCategory
            // 
            CmbCategory.DropDownStyle = ComboBoxStyle.DropDownList;
            CmbCategory.Location = new Point(20, 194);
            CmbCategory.Name = "CmbCategory";
            CmbCategory.Size = new Size(200, 23);
            CmbCategory.TabIndex = 8;
            // 
            // LblSubCategory
            // 
            LblSubCategory.Location = new Point(20, 231);
            LblSubCategory.Name = "LblSubCategory";
            LblSubCategory.Size = new Size(200, 23);
            LblSubCategory.TabIndex = 0;
            LblSubCategory.Text = "Selecione a subcategoria (Opcional):";
            // 
            // CmbSubCategory
            // 
            CmbSubCategory.DropDownStyle = ComboBoxStyle.DropDownList;
            CmbSubCategory.Location = new Point(20, 257);
            CmbSubCategory.Name = "CmbSubCategory";
            CmbSubCategory.Size = new Size(200, 23);
            CmbSubCategory.TabIndex = 10;
            // 
            // LblGenre
            // 
            LblGenre.Location = new Point(20, 296);
            LblGenre.Name = "LblGenre";
            LblGenre.Size = new Size(200, 23);
            LblGenre.TabIndex = 0;
            LblGenre.Text = "Selecione o gênero:";
            // 
            // CmbGenre
            // 
            CmbGenre.DropDownStyle = ComboBoxStyle.DropDownList;
            CmbGenre.Location = new Point(20, 322);
            CmbGenre.Name = "CmbGenre";
            CmbGenre.Size = new Size(200, 23);
            CmbGenre.TabIndex = 10;
            // 
            // Form1
            // 
            ClientSize = new Size(537, 368);
            Controls.Add(LblWarning);
            Controls.Add(LblSheet);
            Controls.Add(TxtSheet);
            Controls.Add(BtnSelectSheet);
            Controls.Add(LblPath);
            Controls.Add(TxtPath);
            Controls.Add(BtnSelectPath);
            Controls.Add(LblCategory);
            Controls.Add(CmbCategory);
            Controls.Add(LblSubCategory);
            Controls.Add(CmbSubCategory);
            Controls.Add(LblGenre);
            Controls.Add(CmbGenre);
            Controls.Add(BtnConvert);
            Name = "Form1";
            Text = "FlexImport - Conversor de Planilhas";
            ResumeLayout(false);
            PerformLayout();
        }
    }
}