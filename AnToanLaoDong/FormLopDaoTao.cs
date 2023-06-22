using AnToanLaoDong.ReportTemplate;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace AnToanLaoDong
{
    public partial class FormLopDaoTao : Form
    {
        private XDocument xmldoc;
        private string File_Lop_Dao_Tao = "LopDaoTao.xml";
        private string File_Hoc_Vien = "HocVien.xml";
        public FormLopDaoTao()
        {
            InitializeComponent();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            DialogResult dg = MessageBox.Show("Bạn có chắc muốn thoát?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dg == DialogResult.OK)
            {
                Application.Exit();
            }
        }

        private void FormLopDaoTao_Load(object sender, EventArgs e)
        {
            LayDSLopDaoTao();
        }

        private void LayDSLopDaoTao()
        {
            try
            {
                xmldoc = XDocument.Load(File_Lop_Dao_Tao);
                var lopDaoTao = xmldoc.Descendants("LopDaoTao");
                if (txtSearchMaLop.Text == null || txtSearchMaLop.Text == "")
                {

                }
                else
                {
                    lopDaoTao = lopDaoTao.Where(x => x.Element("MaLop").Value.ToLower().Contains(txtSearchMaLop.Text.ToLower()));
                }
                if (txtSearchDonVi.Text == null || txtSearchDonVi.Text == "")
                {

                }
                else
                {
                    lopDaoTao = lopDaoTao.Where(x => x.Element("DonVi").Value.ToLower().Contains(txtSearchDonVi.Text.ToLower()));
                }
                lopDaoTao = lopDaoTao.OrderByDescending(p => Int32.Parse(p.Element("ID").Value));
                xmldoc = XDocument.Load(File_Lop_Dao_Tao);
                var data = lopDaoTao.Select(p => new
                {
                    ID = (string)p.Element("ID"),
                    MaLop = (string)p.Element("MaLop"),
                    Nam = (string)p.Element("Nam"),
                    SoLuongHV = (string)p.Element("SoLuongHV"),
                    TuNgay = (string)p.Element("TuNgay"),
                    DenNgay = (string)p.Element("DenNgay"),
                    DonVi = (string)p.Element("DonVi"),
                    DiaChi = (string)p.Element("DiaChi"),
                    NoiDungDT1 = (string)p.Element("NoiDungDT1"),
                    NoiDungDT2 = (string)p.Element("NoiDungDT2"),
                    NoiDungDT3 = (string)p.Element("NoiDungDT3"),
                    NoiDungDT4 = (string)p.Element("NoiDungDT4"),
                    NoiDungDT5 = (string)p.Element("NoiDungDT5"),
                }).ToList();

                if (txtSearchMaLop.Text == "" && txtSearchDonVi.Text == "")
                {
                    data = data.Take(10).ToList();
                }
                dtgDSLop.DataSource = data;
                dtgDSLop.Columns[0].Width = 35;
                dtgDSLop.Columns[0].HeaderText = "ID";
                dtgDSLop.Columns[1].Width = 130;
                dtgDSLop.Columns[1].HeaderText = "Mã lớp";
                dtgDSLop.Columns[2].Width = 100;
                dtgDSLop.Columns[2].HeaderText = "Năm";
                dtgDSLop.Columns[3].Width = 120;
                dtgDSLop.Columns[3].HeaderText = "Số lượng HV";
                dtgDSLop.Columns[4].Width = 80;
                dtgDSLop.Columns[4].HeaderText = "Từ ngày";
                dtgDSLop.Columns[5].Width = 80;
                dtgDSLop.Columns[5].HeaderText = "Đến ngày";
                dtgDSLop.Columns[6].Width = 250;
                dtgDSLop.Columns[6].HeaderText = "Đơn vị";
                dtgDSLop.Columns[7].Width = 250;
                dtgDSLop.Columns[7].HeaderText = "Địa chỉ";
                dtgDSLop.Columns[8].Width = 0;
                dtgDSLop.Columns[9].Width = 0;
                dtgDSLop.Columns[10].Width = 0;
                dtgDSLop.Columns[11].Width = 0;
                dtgDSLop.Columns[12].Width = 0;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void Reset()
        {
            txtMaLop.Text = "";
            txtDonVi.Text = "";
            txtDiaChi.Text = "";
            txtNam.Text = "";
            txtID.Text = "";
            txtSoLuongHV.Text = "";
            txtTuNgay.Text = "";
            txtDenNgay.Text = "";
            txtNoiDungDT1.Text = "";
            txtNoiDungDT2.Text = "";
            txtNoiDungDT3.Text = "";
            txtNoiDungDT4.Text = "";
            txtNoiDungDT5.Text = "";
        }

        private bool KiemTraThongTin()
        {
            if (txtMaLop.Text == "")
            {
                MessageBox.Show("Vui lòng điền mã lớp.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaLop.Focus();
                return false;
            }
            //if (txtDonVi.Text == "")
            //{
            //    MessageBox.Show("Vui lòng điền tên đơn vị tham gia đào tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    txtDonVi.Focus();
            //    return false;
            //}
            //if (txtDiaChi.Text == "")
            //{
            //    MessageBox.Show("Vui lòng điền địa chỉ đơn vị tham gia đào tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    txtDiaChi.Focus();
            //    return false;
            //}

            if (txtNam.Text == "")
            {
                MessageBox.Show("Vui lòng điền năm đào tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtNam.Focus();
                return false;
            }
            if (txtSoLuongHV.Text == "")
            {
                MessageBox.Show("Vui lòng điền số lượng học viên tham gia đào tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtSoLuongHV.Focus();
                return false;
            }
            if (txtTuNgay.Text == "")
            {
                MessageBox.Show("Vui lòng điền thời gian bắt đầu đào tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTuNgay.Focus();
                return false;
            }
            if (txtDenNgay.Text == "")
            {
                MessageBox.Show("Vui lòng điền thời gian kết thúc đào tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtTuNgay.Focus();
                return false;
            }
            return true;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if (KiemTraThongTin())
            {
                XElement Lop = xmldoc.Descendants("LopDaoTao").FirstOrDefault(p => (p.Element("ID").Value == txtID.Text || p.Element("MaLop").Value.ToLower() == txtMaLop.Text.ToLower()));
                if (Lop != null)
                {
                    MessageBox.Show("Trong danh sách lớp không được có 2 ID hoặc Mã lớp trùng nhau.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtID.Focus();

                }
                else
                {
                    XElement lopDaoTao = new XElement("LopDaoTao",
                    new XElement("ID", txtID.Text),
                    new XElement("MaLop", txtMaLop.Text),
                    new XElement("DonVi", txtDonVi.Text),
                    new XElement("DiaChi", txtDiaChi.Text),
                    new XElement("Nam", txtNam.Text),
                    new XElement("SoLuongHV", txtSoLuongHV.Text),
                    new XElement("TuNgay", txtTuNgay.Text),
                    new XElement("DenNgay", txtDenNgay.Text),
                    new XElement("NoiDungDT1", txtNoiDungDT1.Text),
                    new XElement("NoiDungDT2", txtNoiDungDT2.Text),
                    new XElement("NoiDungDT3", txtNoiDungDT3.Text),
                    new XElement("NoiDungDT4", txtNoiDungDT4.Text),
                    new XElement("NoiDungDT5", txtNoiDungDT5.Text)
                    );
                    xmldoc.Root.Add(lopDaoTao);
                    xmldoc.Save(File_Lop_Dao_Tao);
                    LayDSLopDaoTao();
                    MessageBox.Show("Thêm mới lớp đào tạo thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Reset();
                }
            }
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (KiemTraThongTin())
            {
                XElement emp = xmldoc.Descendants("LopDaoTao").FirstOrDefault(p => (p.Element("ID").Value == txtID.Text));
                if (emp != null)
                {
                    emp.Element("MaLop").Value = txtMaLop.Text;
                    emp.Element("DonVi").Value = txtDonVi.Text;
                    emp.Element("DiaChi").Value = txtDiaChi.Text;
                    emp.Element("Nam").Value = txtNam.Text;
                    emp.Element("SoLuongHV").Value = txtSoLuongHV.Text;
                    emp.Element("TuNgay").Value = txtTuNgay.Text;
                    emp.Element("DenNgay").Value = txtDenNgay.Text;
                    emp.Element("NoiDungDT1").Value = txtNoiDungDT1.Text;
                    emp.Element("NoiDungDT2").Value = txtNoiDungDT2.Text;
                    emp.Element("NoiDungDT3").Value = txtNoiDungDT3.Text;
                    emp.Element("NoiDungDT4").Value = txtNoiDungDT4.Text;
                    emp.Element("NoiDungDT5").Value = txtNoiDungDT5.Text;
                    xmldoc.Save(File_Lop_Dao_Tao);
                    LayDSLopDaoTao();
                    MessageBox.Show("Sửa lớp đào tạo thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Reset();
                }
            }
        }

        private void dtgDSLop_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            row = dtgDSLop.Rows[e.RowIndex];
            txtID.Text = Convert.ToString(row.Cells["ID"].Value);
            txtMaLop.Text = Convert.ToString(row.Cells["MaLop"].Value);
            txtDonVi.Text = Convert.ToString(row.Cells["DonVi"].Value);
            txtDiaChi.Text = Convert.ToString(row.Cells["DiaChi"].Value);
            txtNam.Text = Convert.ToString(row.Cells["Nam"].Value);
            txtSoLuongHV.Text = Convert.ToString(row.Cells["SoLuongHV"].Value);
            txtTuNgay.Text = Convert.ToString(row.Cells["TuNgay"].Value);
            txtDenNgay.Text = Convert.ToString(row.Cells["DenNgay"].Value);
            txtNoiDungDT1.Text = Convert.ToString(row.Cells["NoiDungDT1"].Value);
            txtNoiDungDT2.Text = Convert.ToString(row.Cells["NoiDungDT2"].Value);
            txtNoiDungDT3.Text = Convert.ToString(row.Cells["NoiDungDT3"].Value);
            txtNoiDungDT4.Text = Convert.ToString(row.Cells["NoiDungDT4"].Value);
            txtNoiDungDT5.Text = Convert.ToString(row.Cells["NoiDungDT5"].Value);
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show("Bạn chắc chắn xóa lớp này ??",
                                          "Xác nhận xóa!!",
                                          MessageBoxButtons.YesNo);
            if (confirmResult == DialogResult.Yes)
            {
                if (txtID.Text == null || txtID.Text == "")
                {
                    MessageBox.Show("Vui lòng điền ID lớp cần xóa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtID.Focus();
                }
                else
                {
                    var xmlDaoTao = XDocument.Load(File_Lop_Dao_Tao);
                    var xmlHocVien = XDocument.Load(File_Hoc_Vien);
                    XElement emp = xmlDaoTao.Descendants("LopDaoTao").FirstOrDefault(p => (p.Element("ID").Value == txtID.Text));
                    var dsHocVien = xmlHocVien.Descendants("HocVien");
                    dsHocVien = dsHocVien.Where(x => x.Element("MaLop").Value.ToLower().Contains(txtMaLop.Text.ToLower()));

                    if (emp != null)
                    {
                        emp.Remove();
                        xmlDaoTao.Save(File_Lop_Dao_Tao);
                        dsHocVien.Remove();
                        xmlHocVien.Save(File_Hoc_Vien);
                        LayDSLopDaoTao();
                    }
                }
            }            
        }

        private void btnNhapHV_Click(object sender, EventArgs e)
        {
            FormHocVien f = new FormHocVien(txtMaLop.Text, txtDonVi.Text, txtTuNgay.Text, txtDenNgay.Text);
            this.Hide();
            f.ShowDialog();
        }



        private void btnMau6_Click(object sender, EventArgs e)
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            var xmllop = XDocument.Load(File_Lop_Dao_Tao);
            var info = xmllop.Descendants("LopDaoTao").Where(p => (p.Element("ID").Value == txtID.Text));

            if (info.Count() == 0)
            {
                MessageBox.Show("Chưa chọn lớp đào tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                if (info.Count() == 1)
                {
                    var data = xmldoc.Descendants("HocVien").Where(p => (p.Element("MaLop").Value == txtMaLop.Text && p.Element("SoTheATLD").Value != "")).Select(p => new
                    {
                        HoTen = (string)p.Element("HoTen").Value,
                        NgaySinh = (string)p.Element("NgaySinh").Value,
                        CCCD = (string)p.Element("CCCD").Value,
                        CongViec = (string)p.Element("CongViec").Value,
                        DonVi = (string)p.Element("DonVi").Value,
                        SoTheATLD = (string)p.Element("SoTheATLD").Value,
                        KhoaHL = (string)p.Element("KhoaHuanLuyen").Value,
                        TuNgayDD = (string)p.Element("TuNgay").Value,
                        TuNgayMM = (string)p.Element("TuNgay").Value.ToString().Split('/')[1],
                        TuNgayYY = (string)(Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[2]) - 2000).ToString(),
                        DenNgayDD = (string)p.Element("DenNgay").Value,
                        DenNgayMM = (string)p.Element("DenNgay").Value.ToString().Split('/')[1],
                        DenNgayYY = (string)(Int32.Parse(p.Element("DenNgay").Value.ToString().Split('/')[2]) - 2000).ToString(),
                        NgayCapTheDD = (string)p.Element("NgayCapThe").Value.ToString().Split('/')[0],
                        NgayCapTheMM = (string)p.Element("NgayCapThe").Value.ToString().Split('/')[1],
                        NgayCapTheY = (string)p.Element("NgayCapThe").Value.ToString().Split('/')[2],
                        HieuLucThe = (string)p.Element("HieuLucThe").Value,
                    }).ToList();

                    MauTheATLD thongTin = new MauTheATLD();
                    thongTin.SetDataSource(data);
                    FormInBaoCao f = new FormInBaoCao();
                    f.crystalReportViewer1.ReportSource = thongTin;
                    f.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Có học viên khác trong lớp trùng ID.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btnMau8_Click(object sender, EventArgs e)
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            var xmllop = XDocument.Load(File_Lop_Dao_Tao);
            var info = xmllop.Descendants("LopDaoTao").Where(p => (p.Element("ID").Value == txtID.Text));
            if (info.Count() == 0)
            {
                MessageBox.Show("Chưa chọn lớp đào tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                if (info.Count() == 1)
                {
                    var xmlDaoTao = XDocument.Load(File_Lop_Dao_Tao);
                    var lopDaoTao = xmlDaoTao.Descendants("LopDaoTao").First(p => (p.Element("MaLop").Value == txtMaLop.Text));
                    var data = xmldoc.Descendants("HocVien").Where(p => (p.Element("MaLop").Value == txtMaLop.Text && p.Element("SoCNATLD").Value != "")).Select(p => new
                    {
                        SoCNATLD = p.Element("SoCNATLD").Value,
                        HoTen = p.Element("HoTen").Value,
                        GioiTinh = p.Element("GioiTinh").Value,
                        NgaySinh = p.Element("NgaySinh").Value,
                        QuocTich = p.Element("QuocTich").Value,
                        CCCD = p.Element("CCCD").Value,
                        ChucVu = p.Element("ChucVu").Value,
                        DoiTuong = p.Element("DoiTuong").Value,
                        DonVi = p.Element("DonVi").Value,
                        TuNgayDD = p.Element("TuNgay").Value.ToString().Split('/')[0],
                        TuNgayMM = p.Element("TuNgay").Value.ToString().Split('/')[1],
                        TuNgayYYYY = p.Element("TuNgay").Value.ToString().Split('/')[2],
                        DenNgayDD = p.Element("DenNgay").Value.ToString().Split('/')[0],
                        DenNgayMM = p.Element("DenNgay").Value.ToString().Split('/')[1],
                        DenNgayYYYY = p.Element("DenNgay").Value.ToString().Split('/')[2],
                        XepLoai = p.Element("XepLoai").Value,
                        NgayCapCNDD = p.Element("NgayCapCN").Value.ToString().Split('/')[0],
                        NgayCapCNMM = p.Element("NgayCapCN").Value.ToString().Split('/')[1],
                        NgayCapCNYYYY = p.Element("NgayCapCN").Value.ToString().Split('/')[2],
                        HieuLucCNDD = p.Element("HieuLucCN").Value.ToString().Split('/')[0],
                        HieuLucCNMM = p.Element("HieuLucCN").Value.ToString().Split('/')[1],
                        HieuLucCNYYYY = p.Element("HieuLucCN").Value.ToString().Split('/')[2],
                        NoiDungDT1 = lopDaoTao.Element("NoiDungDT1").Value,
                        NoiDungDT2 = lopDaoTao.Element("NoiDungDT2").Value,
                        NoiDungDT3 = lopDaoTao.Element("NoiDungDT3").Value,
                        NoiDungDT4 = lopDaoTao.Element("NoiDungDT4").Value,
                        NoiDungDT5 = lopDaoTao.Element("NoiDungDT5").Value,
                    }).ToList();

                    MauChungNhanATLD thongTin = new MauChungNhanATLD();
                    thongTin.SetDataSource(data);
                    FormInBaoCao f = new FormInBaoCao();
                    f.crystalReportViewer1.ReportSource = thongTin;
                    f.ShowDialog();
                }
                else
                {
                    MessageBox.Show("Có lớp đào tạo khác trong lớp trùng ID, vui lòng xóa dữ liệu thừa hoặc chỉnh sang ID khác", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void btnMau9_Click(object sender, EventArgs e)
        {
            FormBaoCao f = new FormBaoCao();
            this.Hide();
            f.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormNhapDuLieu f = new FormNhapDuLieu();
            this.Hide();
            f.ShowDialog();
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            LayDSLopDaoTao();
        }


        private void txtSearchDonVi_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
