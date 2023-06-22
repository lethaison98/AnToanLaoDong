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
    public partial class FormHocVien : Form
    {
        private string MaLop, TuNgay, DenNgay;
        private XDocument xmldoc;
        private string File_Hoc_Vien = "HocVien.xml";
        private string File_Lop_Dao_Tao = "LopDaoTao.xml";

        public FormHocVien()
        {
            InitializeComponent();
        }
        public FormHocVien(string maLop, string donVi, string tuNgay, string denNgay)
        {
            InitializeComponent();
            this.MaLop = maLop;
            this.TuNgay = tuNgay;
            this.DenNgay = denNgay;
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            FormLopDaoTao f = new FormLopDaoTao();
            this.Hide();
            f.ShowDialog();
        }

        private void FormHocVien_Load(object sender, EventArgs e)
        {
            txtMaLop.Text = MaLop;
            txtSearchMaLop.Text = MaLop;
            txtTuNgay.Text = TuNgay;
            txtDenNgay.Text = DenNgay;
            LayDSNS();
        }

        public void LoadDataFromXmlFile()
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            var data = xmldoc.Descendants("HocVien").Select(p => new
            {
                ID = (string)p.Element("ID"),
                MaLop = (string)p.Element("MaLop"),
                HoTen = (string)p.Element("HoTen"),
                NgaySinh = (string)p.Element("NgaySinh"),
                CCCD = (string)p.Element("CCCD"),
                QuocTich = (string)p.Element("QuocTich"),
                CongViec = (string)p.Element("CongViec"),
                DonVi = (string)p.Element("DonVi"),
                TuNgay = (string)p.Element("TuNgay"),
                DenNgay = (string)p.Element("DenNgay"),
            }).OrderBy(p => p.ID).ToList();
        }
        private void LayDSNS()
        {
            try
            {
                xmldoc = XDocument.Load(File_Hoc_Vien);
                var hocVien = xmldoc.Descendants("HocVien");
                if (txtSearchMaLop.Text == null || txtSearchMaLop.Text == "")
                {

                }
                else
                {
                    hocVien = hocVien.Where(x => x.Element("MaLop").Value.ToLower().Contains(txtSearchMaLop.Text.ToLower()));
                }
                if (txtSearchDonVi.Text == null || txtSearchDonVi.Text == "")
                {

                }
                else
                {
                    hocVien = hocVien.Where(x => x.Element("DonVi").Value.ToLower().Contains(txtSearchDonVi.Text.ToLower()));
                }
                if (txtSearchHoTen.Text == null || txtSearchHoTen.Text == "")
                {

                }
                else
                {
                    hocVien = hocVien.Where(x => x.Element("HoTen").Value.ToLower().Contains(txtSearchHoTen.Text.ToLower()));
                }
                var list = hocVien.OrderBy(x => x.Element("MaLop").Value).ThenBy(x => Int32.Parse(x.Element("ID").Value)).ToList();
                hocVien = hocVien.OrderBy(x => x.Element("MaLop").Value).ThenBy(x => Int32.Parse(x.Element("ID").Value)).ToList();

                var data = hocVien.Select(p => new
                {
                    ID = (string)p.Element("ID"),
                    MaLop = (string)p.Element("MaLop"),
                    HoTen = (string)p.Element("HoTen"),
                    GioiTinh = (string)p.Element("GioiTinh"),
                    NgaySinh = (string)p.Element("NgaySinh"),
                    CCCD = (string)p.Element("CCCD"),
                    QuocTich = (string)p.Element("QuocTich"),
                    DoiTuong = (string)p.Element("DoiTuong"),
                    DonVi = (string)p.Element("DonVi"),
                    TuNgay = (string)p.Element("TuNgay"),
                    DenNgay = (string)p.Element("DenNgay"),
                    SoCNATLD = (string)p.Element("SoCNATLD"),
                    ChucVu = (string)p.Element("ChucVu"),
                    CongViec = (string)p.Element("CongViec"),
                    XepLoai = (string)p.Element("XepLoai"),
                    NgayCapCN = (string)p.Element("NgayCapCN"),
                    HieuLucCN = (string)p.Element("HieuLucCN"),
                    SoTheATLD = (string)p.Element("SoTheATLD"),
                    KhoaHuanLuyen = (string)p.Element("KhoaHuanLuyen"),
                    NgayCapThe = (string)p.Element("NgayCapThe"),
                    HieuLucThe = (string)p.Element("HieuLucThe"),
                }).ToList();
                dtgDSNS.DataSource = data;
                dtgDSNS.Columns[0].Width = 35;
                dtgDSNS.Columns[0].HeaderText = "ID";
                dtgDSNS.Columns[1].Width = 130;
                dtgDSNS.Columns[1].HeaderText = "Mã lớp";
                dtgDSNS.Columns[2].Width = 130;
                dtgDSNS.Columns[2].HeaderText = "Họ Tên";
                dtgDSNS.Columns[3].Width = 70;
                dtgDSNS.Columns[3].HeaderText = "Giới tính";
                dtgDSNS.Columns[4].Width = 80;
                dtgDSNS.Columns[4].HeaderText = "Ngày sinh";
                dtgDSNS.Columns[5].Width = 100;
                dtgDSNS.Columns[5].HeaderText = "Quốc tịch";
                dtgDSNS.Columns[6].Width = 80;
                dtgDSNS.Columns[6].HeaderText = "CMND/CCCD";
                dtgDSNS.Columns[7].Width = 80;
                dtgDSNS.Columns[7].HeaderText = "Đối tượng";
                dtgDSNS.Columns[8].Width = 170;
                dtgDSNS.Columns[8].HeaderText = "Đơn vị";
                dtgDSNS.Columns[9].Width = 80;
                dtgDSNS.Columns[9].HeaderText = "Từ ngày";
                dtgDSNS.Columns[10].Width = 80;
                dtgDSNS.Columns[10].HeaderText = "Đến ngày";
                dtgDSNS.Columns[11].Width = 0;
                dtgDSNS.Columns[12].Width = 0;
                dtgDSNS.Columns[13].Width = 0;
                dtgDSNS.Columns[14].Width = 0;
                dtgDSNS.Columns[15].Width = 0;
                dtgDSNS.Columns[16].Width = 0;
                dtgDSNS.Columns[17].Width = 0;
                dtgDSNS.Columns[18].Width = 0;
                dtgDSNS.Columns[19].Width = 0;
                dtgDSNS.Columns[20].Width = 0;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

       

        private void Reset()
        {
            txtDonVi.Text = "";
            txtCongViec.Text = "";
            txtQuocTich.Text = "";
            txtHoTen.Text = "";
            txtMaLop.Text = "";
            txtID.Text = "";
            txtCCCD.Text = "";
            rdoNam.Checked = false;
            rdoNu.Checked = false;
            txtNgaySinh.Text = "";
            txtTuNgay.Text ="";
            txtDenNgay.Text = "";
            txtSoCNATLD.Text = "";
            txtChucVu.Text = "";
            lstDoiTuong.SelectedItem = null;
            txtXepLoai.Text = "";
            txtNgayCapCN.Text = "";
            txtHieuLucCN.Text = "";
            txtSoTheATLD.Text = "";
            txtKhoaHuanLuyen.Text = "";
            txtNgayCapThe.Text = "";
            txtHieuLucThe.Text = "";
        }

        private bool KiemTraThongTin()
        {
            if (txtHoTen.Text == "")
            {
                MessageBox.Show("Vui lòng điền họ và tên học viên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtHoTen.Focus();
                return false;
            }
            if (txtMaLop.Text == "")
            {
                MessageBox.Show("Vui lòng điền mã lớp của học viên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtMaLop.Focus();
                return false;
            }
            if (txtDonVi.Text == "")
            {
                MessageBox.Show("Vui lòng điền nơi làm việc của học viên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtDonVi.Focus();
                return false;
            }
            if (lstDoiTuong.SelectedItem == null)
            {
                MessageBox.Show("Vui lòng chọn đối tượng huấn luyện cho học viên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtCongViec.Focus();
                return false;
            }
            if (rdoNam.Checked == false && rdoNu.Checked == false)
            {
                MessageBox.Show("Vui lòng chọn giới tính cho học viên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (txtCCCD.Text == "")
            {
                MessageBox.Show("Vui lòng điền số CMND/CCCD của học viên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtCCCD.Focus();
                return false;
            }
            if (txtQuocTich.Text == "")
            {
                MessageBox.Show("Vui lòng điền quốc tịch.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtCCCD.Focus();
                return false;
            }
            return true;
        }
        private void btnThem_Click(object sender, EventArgs e)
        {
            if (KiemTraThongTin())
            {
                XElement HV = xmldoc.Descendants("HocVien").FirstOrDefault(p => (p.Element("ID").Value == txtID.Text && p.Element("MaLop").Value == txtMaLop.Text));
                if (HV != null)
                {
                    MessageBox.Show("Trong một lớp không được có 2 ID trùng nhau.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtID.Focus();

                }
                else
                {
                    var gioiTinh = "";
                    if (rdoNam.Checked == true)
                    {
                        gioiTinh = rdoNam.Text;
                    }
                    else
                    {
                        gioiTinh = rdoNu.Text;
                    }
                    
                    XElement hocVien = new XElement("HocVien",
                        new XElement("ID", txtID.Text),
                        new XElement("MaLop", txtMaLop.Text),
                        new XElement("HoTen", txtHoTen.Text),
                        new XElement("GioiTinh", gioiTinh),
                        new XElement("NgaySinh", txtNgaySinh.Text),
                        new XElement("QuocTich", txtQuocTich.Text),
                        new XElement("CCCD", txtCCCD.Text),
                        new XElement("CongViec", txtCongViec.Text),
                        new XElement("DonVi", txtDonVi.Text),
                        new XElement("TuNgay", txtTuNgay.Text),
                        new XElement("DenNgay", txtDenNgay.Text),
                        new XElement("SoCNATLD", txtSoCNATLD.Text),
                        new XElement("ChucVu", txtChucVu.Text),
                        new XElement("DoiTuong", lstDoiTuong.SelectedItem.ToString()),
                        new XElement("XepLoai", txtXepLoai.Text),
                        new XElement("NgayCapCN", txtNgayCapCN.Text),
                        new XElement("HieuLucCN", txtHieuLucCN.Text),
                        new XElement("SoTheATLD", txtSoTheATLD.Text),
                        new XElement("KhoaHuanLuyen", txtKhoaHuanLuyen.Text),
                        new XElement("NgayCapThe", txtNgayCapThe.Text),
                        new XElement("HieuLucThe", txtHieuLucThe.Text)
                        );
                    xmldoc.Root.Add(hocVien);
                    xmldoc.Save(File_Hoc_Vien);
                    LayDSNS();
                    MessageBox.Show("Thêm mới học viên thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Reset();
                }
            }
        }
        private void btnSua_Click(object sender, EventArgs e)
        {
            if (KiemTraThongTin())
            {
                var gioiTinh = "";
                if (rdoNam.Checked == true)
                {
                    gioiTinh = rdoNam.Text;
                }
                else
                {
                    gioiTinh = rdoNu.Text;
                }

                XElement emp = xmldoc.Descendants("HocVien").FirstOrDefault(p => (p.Element("ID").Value == txtID.Text && p.Element("MaLop").Value == txtMaLop.Text));
                if (emp != null)
                {
                    emp.Element("HoTen").Value = txtHoTen.Text;
                    emp.Element("NgaySinh").Value = txtNgaySinh.Text;
                    emp.Element("GioiTinh").Value = gioiTinh;
                    emp.Element("QuocTich").Value = txtQuocTich.Text;
                    emp.Element("CCCD").Value = txtCCCD.Text;
                    emp.Element("CongViec").Value = txtCongViec.Text;
                    emp.Element("DonVi").Value = txtDonVi.Text;
                    emp.Element("TuNgay").Value = txtTuNgay.Text;
                    emp.Element("DenNgay").Value = txtDenNgay.Text;
                    emp.Element("SoCNATLD").Value = txtSoCNATLD.Text;
                    emp.Element("XepLoai").Value = txtXepLoai.Text;
                    emp.Element("ChucVu").Value = txtChucVu.Text;
                    emp.Element("DoiTuong").Value = lstDoiTuong.SelectedItem.ToString();
                    emp.Element("NgayCapCN").Value = txtNgayCapCN.Text;
                    emp.Element("HieuLucCN").Value = txtHieuLucCN.Text;
                    emp.Element("SoTheATLD").Value = txtSoTheATLD.Text;
                    emp.Element("KhoaHuanLuyen").Value = txtKhoaHuanLuyen.Text;
                    emp.Element("NgayCapThe").Value = txtNgayCapThe.Text;
                    emp.Element("HieuLucThe").Value = txtHieuLucThe.Text;
                    
                    xmldoc.Save(File_Hoc_Vien);
                    LayDSNS();
                    MessageBox.Show("Sửa học viên thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Reset();
                }
            }
        }


        private void dtgDSNS_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            row = dtgDSNS.Rows[e.RowIndex];
            txtID.Text = Convert.ToString(row.Cells["ID"].Value);
            txtMaLop.Text = Convert.ToString(row.Cells["MaLop"].Value);
            txtHoTen.Text = Convert.ToString(row.Cells["HoTen"].Value);
            txtNgaySinh.Text = Convert.ToString(row.Cells["NgaySinh"].Value);
            txtDonVi.Text = Convert.ToString(row.Cells["DonVi"].Value);
            txtCongViec.Text = Convert.ToString(row.Cells["CongViec"].Value);
            txtQuocTich.Text = Convert.ToString(row.Cells["QuocTich"].Value);
            txtCCCD.Text = Convert.ToString(row.Cells["CCCD"].Value);
            string GioiTinh = Convert.ToString(row.Cells["GioiTinh"].Value);
            txtTuNgay.Text = Convert.ToString(row.Cells["TuNgay"].Value);
            txtDenNgay.Text = Convert.ToString(row.Cells["DenNgay"].Value);
            txtSoCNATLD.Text = Convert.ToString(row.Cells["SoCNATLD"].Value);
            txtXepLoai.Text = Convert.ToString(row.Cells["XepLoai"].Value);
            txtChucVu.Text = Convert.ToString(row.Cells["ChucVu"].Value);
            lstDoiTuong.SelectedItem = Convert.ToString(row.Cells["DoiTuong"].Value);
            txtNgayCapCN.Text = Convert.ToString(row.Cells["NgayCapCN"].Value);
            txtHieuLucCN.Text = Convert.ToString(row.Cells["HieuLucCN"].Value);
            txtSoTheATLD.Text = Convert.ToString(row.Cells["SoTheATLD"].Value);
            txtKhoaHuanLuyen.Text = Convert.ToString(row.Cells["KhoaHuanLuyen"].Value);
            txtNgayCapThe.Text = Convert.ToString(row.Cells["NgayCapThe"].Value);
            txtHieuLucThe.Text = Convert.ToString(row.Cells["HieuLucThe"].Value);
            if (GioiTinh.Trim() == "Nam")
            {
                rdoNam.Checked = true;
            }
            else
            {
                rdoNu.Checked = true;
            }

        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            var confirmResult = MessageBox.Show("Bạn chắc chắn xóa học viên này ??",
                                     "Xác nhận xóa!!",
                                     MessageBoxButtons.YesNo);
            if (confirmResult == DialogResult.Yes)
            {
                if (txtID.Text == null || txtID.Text == "")
                {
                    MessageBox.Show("Vui lòng điền ID học viên cần xóa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtID.Focus();
                }
                else
                {
                    XElement emp = xmldoc.Descendants("HocVien").FirstOrDefault(p => (p.Element("ID").Value == txtID.Text && p.Element("MaLop").Value == txtMaLop.Text));
                    if (emp != null)
                    {
                        emp.Remove();
                        xmldoc.Save(File_Hoc_Vien);
                        LayDSNS();
                    }
                }
            }
            
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            LayDSNS();
        }

        private void btnInCNATLD_Click(object sender, EventArgs e)
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            var info = xmldoc.Descendants("HocVien").Where(p => (p.Element("ID").Value == txtID.Text && p.Element("MaLop").Value == txtMaLop.Text));
            if (info.Count() == 0)
            {
                MessageBox.Show("Chưa chọn học viên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                if (info.Count() == 1)
                {
                    var xmlDaoTao = XDocument.Load(File_Lop_Dao_Tao);
                    var lopDaoTao = xmlDaoTao.Descendants("LopDaoTao").First(p => (p.Element("MaLop").Value == txtMaLop.Text));
                    var data = xmldoc.Descendants("HocVien").Where(p => (p.Element("ID").Value == txtID.Text && p.Element("MaLop").Value == txtMaLop.Text &&p.Element("SoCNATLD").Value != "")).Select(p => new
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
                    MessageBox.Show("Có học viên khác trong lớp trùng ID.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

        }

        private void btnInTheATLD_Click(object sender, EventArgs e)
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            var info = xmldoc.Descendants("HocVien").Where(p => (p.Element("ID").Value == txtID.Text && p.Element("MaLop").Value == txtMaLop.Text));
            if(info.Count() == 0)
            {
                MessageBox.Show("Chưa chọn học viên.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                if (info.Count() == 1)
                {
                    var data = xmldoc.Descendants("HocVien").Where(p => (p.Element("ID").Value == txtID.Text && p.Element("MaLop").Value == txtMaLop.Text && p.Element("SoTheATLD").Value != "")).Select(p => new
                    {
                        HoTen = p.Element("HoTen").Value,
                        NgaySinh = p.Element("NgaySinh").Value,
                        CCCD = p.Element("CCCD").Value,
                        CongViec = p.Element("CongViec").Value,
                        DonVi = p.Element("DonVi").Value,
                        SoTheATLD = p.Element("SoTheATLD").Value,
                        KhoaHL = p.Element("KhoaHuanLuyen").Value,
                        TuNgayDD = p.Element("TuNgay").Value,
                        TuNgayMM = p.Element("TuNgay").Value.ToString().Split('/')[1],
                        TuNgayYY = (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[2]) - 2000).ToString(),
                        DenNgayDD = p.Element("DenNgay").Value,
                        DenNgayMM = p.Element("DenNgay").Value.ToString().Split('/')[1],
                        DenNgayYY = (Int32.Parse(p.Element("DenNgay").Value.ToString().Split('/')[2]) - 2000).ToString(),
                        NgayCapTheDD = p.Element("NgayCapThe").Value.ToString().Split('/')[0],
                        NgayCapTheMM = p.Element("NgayCapThe").Value.ToString().Split('/')[1],
                        NgayCapTheY = p.Element("NgayCapThe").Value.ToString().Split('/')[2],
                        HieuLucThe = p.Element("HieuLucThe").Value,
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
    }
}
