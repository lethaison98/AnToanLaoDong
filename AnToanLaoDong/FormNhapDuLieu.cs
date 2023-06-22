using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace AnToanLaoDong
{
    public partial class FormNhapDuLieu : Form
    {
        private XDocument xmldoc;
        private string File_Hoc_Vien = "HocVien.xml";
        private string File_Lop_Dao_Tao = "LopDaoTao.xml";
        public FormNhapDuLieu()
        {
            InitializeComponent();
        }
        DataSet ds;
        private void btn_import_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                        {
                            IExcelDataReader reader;
                            if (ofd.FilterIndex == 2)
                            {
                                reader = ExcelReaderFactory.CreateBinaryReader(stream);
                            }
                            else
                            {
                                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                            }

                            ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });

                            cb_sheet.Items.Clear();
                            foreach (DataTable dt in ds.Tables)
                            {
                                cb_sheet.Items.Add(dt.TableName);
                            }
                            reader.Close();
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Vui lòng đóng file excel và thử lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }

        }

        private void cb_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.DataSource = ds.Tables[cb_sheet.SelectedIndex];
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            try
            {
                DataTable dt = new DataTable();
                dt = ds.Tables[cb_sheet.SelectedIndex];

                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if(dt.Rows[i][0].ToString() != "" && dt.Rows[i][1].ToString() != ""  && dt.Rows[i][2].ToString() != "")
                        {
                            XElement hocVien = new XElement("HocVien",
                            new XElement("ID", dt.Rows[i][0].ToString()),
                            new XElement("MaLop", dt.Rows[i][1].ToString()),
                            new XElement("HoTen", dt.Rows[i][2].ToString()),
                            new XElement("GioiTinh", dt.Rows[i][3].ToString()),
                            new XElement("NgaySinh", dt.Rows[i][4].ToString().Split(' ')[0]),
                            new XElement("CCCD", dt.Rows[i][5].ToString()),
                            new XElement("QuocTich", dt.Rows[i][6].ToString()),
                            new XElement("DoiTuong", dt.Rows[i][7].ToString()),
                            new XElement("DonVi", dt.Rows[i][8].ToString()),
                            new XElement("TuNgay", dt.Rows[i][9].ToString().Split(' ')[0]),
                            new XElement("DenNgay", dt.Rows[i][10].ToString().Split(' ')[0]),
                            new XElement("SoCNATLD", dt.Rows[i][11].ToString()),
                            new XElement("ChucVu", dt.Rows[i][12].ToString()),
                            new XElement("XepLoai", dt.Rows[i][13].ToString()),
                            new XElement("NgayCapCN", dt.Rows[i][14].ToString().Split(' ')[0]),
                            new XElement("HieuLucCN", dt.Rows[i][15].ToString().Split(' ')[0]),
                            new XElement("SoTheATLD", dt.Rows[i][16].ToString()),
                            new XElement("CongViec", dt.Rows[i][17].ToString()),
                            new XElement("KhoaHuanLuyen", dt.Rows[i][18].ToString()),
                            new XElement("NgayCapThe", dt.Rows[i][19].ToString().Split(' ')[0]),
                            new XElement("HieuLucThe", dt.Rows[i][20].ToString().Split(' ')[0])
                            );
                            xmldoc.Root.Add(hocVien);
                        }
                        
                    }
                    xmldoc.Save(File_Hoc_Vien);
                    MessageBox.Show("Thêm mới học viên thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ds.Clear();
                }
                else
                {
                    MessageBox.Show("File không đúng định dạng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch{
                MessageBox.Show("Vui lòng chọn file và sheet hợp lệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
   
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            FormLopDaoTao f = new FormLopDaoTao();
            this.Hide();
            f.ShowDialog();
        }

    }
}
