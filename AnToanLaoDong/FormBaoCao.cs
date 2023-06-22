using AnToanLaoDong.Dataset;
using AnToanLaoDong.ReportTemplate;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace AnToanLaoDong
{
    public partial class FormBaoCao : Form
    {
        private XDocument xmldoc;
        private string File_Lop_Dao_Tao = "LopDaoTao.xml";
        private string File_Hoc_Vien = "HocVien.xml";
        public FormBaoCao()
        {
            InitializeComponent();
            loadCombobox();
        }

        public void loadCombobox()
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            var DsDonVi = xmldoc.Descendants("HocVien").Select(p => p.Element("DonVi").Value).Distinct().ToList();
            lstDonViBaoCao9.DataSource = DsDonVi;
            var SelectYear = new List<String>();
            int i = DateTime.Now.Year;
            do
            {
                SelectYear.Add(i.ToString());
                i--;

            } while (i > DateTime.Now.Year - 10);
            lstNamBaoCao5.DataSource = SelectYear;
            lstNamBaoCao9.DataSource = SelectYear;
        }

        private void btnInBaoCao9_Click(object sender, EventArgs e)
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            var xmlDaoTao = XDocument.Load(File_Lop_Dao_Tao);
            var dsLopDaoTao = xmlDaoTao.Descendants("LopDaoTao").Where(p => (p.Element("Nam").Value == lstNamBaoCao9.SelectedItem.ToString())).Select(p => p.Element("MaLop").Value).ToList();
            var dsHocVien = xmldoc.Descendants("HocVien").Where(p => (p.Element("DonVi").Value == lstDonViBaoCao9.SelectedItem.ToString() && dsLopDaoTao.Contains(p.Element("MaLop").Value)));


            var nhom1 = dsHocVien.Where(p => p.Element("DoiTuong").Value == "Nhóm 1").Select(p => new
            {
                HoTen = p.Element("HoTen").Value,
                NamSinh = p.Element("NgaySinh").Value.ToString().Split('/')[2],
                CongViec = p.Element("CongViec").Value,
                DonVi = p.Element("DonVi").Value,
                TuNgay = String.Concat(p.Element("TuNgay").Value, " - ", p.Element("DenNgay").Value),
                DenNgay = p.Element("DenNgay").Value,
                XepLoai = p.Element("XepLoai").Value,
                Nam = lstNamBaoCao9.SelectedItem.ToString(),
                NgayTiepTheo = String.Concat(1, "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[1]) - 1).ToString(), "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[2]) + 2).ToString()),
            }).ToList();

            var nhom2 = dsHocVien.Where(p => p.Element("DoiTuong").Value == "Nhóm 2").Select(p => new
            {
                HoTen = p.Element("HoTen").Value,
                NamSinh = p.Element("NgaySinh").Value.ToString().Split('/')[2],
                CongViec = p.Element("CongViec").Value,
                DonVi = p.Element("DonVi").Value,
                TuNgay = String.Concat(p.Element("TuNgay").Value, " - ", p.Element("DenNgay").Value),
                DenNgay = p.Element("DenNgay").Value,
                XepLoai = p.Element("XepLoai").Value,
                Nam = lstNamBaoCao9.SelectedItem.ToString(),
                NgayTiepTheo = String.Concat(1, "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[1]) - 1).ToString(), "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[2]) + 2).ToString()),
            }).ToList();
            var nhom3 = dsHocVien.Where(p => p.Element("DoiTuong").Value == "Nhóm 3").Select(p => new
            {
                HoTen = p.Element("HoTen").Value,
                NamSinh = p.Element("NgaySinh").Value.ToString().Split('/')[2],
                CongViec = p.Element("CongViec").Value,
                DonVi = p.Element("DonVi").Value,
                TuNgay = String.Concat(p.Element("TuNgay").Value, " - ", p.Element("DenNgay").Value),
                DenNgay = p.Element("DenNgay").Value,
                XepLoai = p.Element("XepLoai").Value,
                Nam = lstNamBaoCao9.SelectedItem.ToString(),
                NgayTiepTheo = String.Concat(1, "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[1]) - 1).ToString(), "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[2]) + 2).ToString()),
            }).ToList();
            var nhom4 = dsHocVien.Where(p => p.Element("DoiTuong").Value == "Nhóm 4").Select(p => new
            {
                HoTen = p.Element("HoTen").Value,
                NamSinh = p.Element("NgaySinh").Value.ToString().Split('/')[2],
                CongViec = p.Element("CongViec").Value,
                DonVi = p.Element("DonVi").Value,
                TuNgay = String.Concat(p.Element("TuNgay").Value, " - ", p.Element("DenNgay").Value),
                DenNgay = p.Element("DenNgay").Value,
                XepLoai = p.Element("XepLoai").Value,
                Nam = lstNamBaoCao9.SelectedItem.ToString(),
                NgayTiepTheo = String.Concat(1, "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[1]) - 1).ToString(), "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[2]) + 2).ToString()),
            }).ToList();
            var nhom5 = dsHocVien.Where(p => p.Element("DoiTuong").Value == "Nhóm 5").Select(p => new
            {
                HoTen = p.Element("HoTen").Value,
                NamSinh = p.Element("NgaySinh").Value.ToString().Split('/')[2],
                CongViec = p.Element("CongViec").Value,
                DonVi = p.Element("DonVi").Value,
                TuNgay = String.Concat(p.Element("TuNgay").Value, " - ", p.Element("DenNgay").Value),
                DenNgay = p.Element("DenNgay").Value,
                XepLoai = p.Element("XepLoai").Value,
                Nam = lstNamBaoCao9.SelectedItem.ToString(),
                NgayTiepTheo = String.Concat(1, "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[1]) - 1).ToString(), "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[2]) + 2).ToString()),
            }).ToList();
            var nhom6 = dsHocVien.Where(p => p.Element("DoiTuong").Value == "Nhóm 6").Select(p => new
            {
                HoTen = p.Element("HoTen").Value,
                NamSinh = p.Element("NgaySinh").Value.ToString().Split('/')[2],
                CongViec = p.Element("CongViec").Value,
                DonVi = p.Element("DonVi").Value,
                TuNgay = String.Concat(p.Element("TuNgay").Value, " - ", p.Element("DenNgay").Value),
                DenNgay = p.Element("DenNgay").Value,
                XepLoai = p.Element("XepLoai").Value,
                Nam = lstNamBaoCao9.SelectedItem.ToString(),
                NgayTiepTheo = String.Concat(1, "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[1]) - 1).ToString(), "/", (Int32.Parse(p.Element("TuNgay").Value.ToString().Split('/')[2]) + 2).ToString()),
            }).ToList();

            Mau9Template thongTin = new Mau9Template();
            var thongTinChung = new ThongTinChungMau9
            {
                NamBaoCao = lstNamBaoCao9.SelectedItem.ToString(),
                DonViBaoCao = lstDonViBaoCao9.SelectedItem.ToString(),
            };
            var listThongTinChung = new List<ThongTinChungMau9>();
            listThongTinChung.Add(thongTinChung);
            thongTin.Database.Tables["Nhom1"].SetDataSource(nhom1);
            thongTin.Database.Tables["Nhom2"].SetDataSource(nhom2);
            thongTin.Database.Tables["Nhom3"].SetDataSource(nhom3);
            thongTin.Database.Tables["Nhom4"].SetDataSource(nhom4);
            thongTin.Database.Tables["Nhom5"].SetDataSource(nhom5);
            thongTin.Database.Tables["Nhom6"].SetDataSource(nhom6);
            thongTin.Database.Tables["ThongTin"].SetDataSource(listThongTinChung);


            FormInBaoCao f = new FormInBaoCao();
            f.crystalReportViewer1.ReportSource = null;
            f.crystalReportViewer1.ReportSource = thongTin;
            f.ShowDialog();
        }
        public class ThongTinChungMau9
        {
            public string NamBaoCao { get; set; }
            public string DonViBaoCao { get; set; }
        }

        private void btnInBaoCao5_Click(object sender, EventArgs e)
        {
            xmldoc = XDocument.Load(File_Hoc_Vien);
            var xmlDaoTao = XDocument.Load(File_Lop_Dao_Tao);
            var dsLopDaoTao = xmlDaoTao.Descendants("LopDaoTao").Where(p => (p.Element("Nam").Value == lstNamBaoCao5.SelectedItem.ToString())).Select(p => p.Element("MaLop").Value).ToList();
            var dsHocVien = xmldoc.Descendants("HocVien").Where(p => dsLopDaoTao.Contains(p.Element("MaLop").Value));


            var nhom1 = new KetQuaHuanLuyenMau5
            {
                DoiTuongHuanLuyen = "Nhóm 1",
                SoNguoiDuocHuanLuyen = xmldoc.Descendants("HocVien").Where(p => p.Element("DoiTuong").Value == "Nhóm 1").Count(),
                SoNguoiDuocCapCN = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 1" && p.Element("SoCNATLD").Value != "")).Count(),
                SoNguoiDuocCapThe = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 1" && p.Element("SoTheATLD").Value != "")).Count(),
            };
            var nhom2 = new KetQuaHuanLuyenMau5
            {
                DoiTuongHuanLuyen = "Nhóm 2",
                SoNguoiDuocHuanLuyen = xmldoc.Descendants("HocVien").Where(p => p.Element("DoiTuong").Value == "Nhóm 2").Count(),
                SoNguoiDuocCapCN = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 2" && p.Element("SoCNATLD").Value != "")).Count(),
                SoNguoiDuocCapThe = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 2" && p.Element("SoTheATLD").Value != "")).Count(),
            };
            var nhom3 = new KetQuaHuanLuyenMau5
            {
                DoiTuongHuanLuyen = "Nhóm 3",
                SoNguoiDuocHuanLuyen = xmldoc.Descendants("HocVien").Where(p => p.Element("DoiTuong").Value == "Nhóm 3").Count(),
                SoNguoiDuocCapCN = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 3" && p.Element("SoCNATLD").Value != "")).Count(),
                SoNguoiDuocCapThe = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 3" && p.Element("SoTheATLD").Value != "")).Count(),
            };
            var nhom4 = new KetQuaHuanLuyenMau5
            {
                DoiTuongHuanLuyen = "Nhóm 4",
                SoNguoiDuocHuanLuyen = xmldoc.Descendants("HocVien").Where(p => p.Element("DoiTuong").Value == "Nhóm 4").Count(),
                SoNguoiDuocCapCN = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 4" && p.Element("SoCNATLD").Value != "")).Count(),
                SoNguoiDuocCapThe = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 4" && p.Element("SoTheATLD").Value != "")).Count(),
            };
            var nhom5 = new KetQuaHuanLuyenMau5
            {
                DoiTuongHuanLuyen = "Nhóm 5",
                SoNguoiDuocHuanLuyen = xmldoc.Descendants("HocVien").Where(p => p.Element("DoiTuong").Value == "Nhóm 5").Count(),
                SoNguoiDuocCapCN = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 5" && p.Element("SoCNATLD").Value != "")).Count(),
                SoNguoiDuocCapThe = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 5" && p.Element("SoTheATLD").Value != "")).Count(),
            };
            var nhom6 = new KetQuaHuanLuyenMau5
            {
                DoiTuongHuanLuyen = "Nhóm 6",
                SoNguoiDuocHuanLuyen = xmldoc.Descendants("HocVien").Where(p => p.Element("DoiTuong").Value == "Nhóm 6").Count(),
                SoNguoiDuocCapCN = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 6" && p.Element("SoCNATLD").Value != "")).Count(),
                SoNguoiDuocCapThe = xmldoc.Descendants("HocVien").Where(p => (p.Element("DoiTuong").Value == "Nhóm 6" && p.Element("SoTheATLD").Value != "")).Count(),
            };
            var tongCong = new KetQuaHuanLuyenMau5
            {
                DoiTuongHuanLuyen = "Tổng cộng",
                SoNguoiDuocHuanLuyen = nhom1.SoNguoiDuocHuanLuyen + nhom2.SoNguoiDuocHuanLuyen + nhom3.SoNguoiDuocHuanLuyen + nhom4.SoNguoiDuocHuanLuyen + nhom5.SoNguoiDuocHuanLuyen + nhom6.SoNguoiDuocHuanLuyen,
                SoNguoiDuocCapCN = nhom1.SoNguoiDuocCapCN + nhom2.SoNguoiDuocCapCN + nhom3.SoNguoiDuocCapCN + nhom4.SoNguoiDuocCapCN + nhom5.SoNguoiDuocCapCN + nhom6.SoNguoiDuocCapCN,
                SoNguoiDuocCapThe = nhom1.SoNguoiDuocCapThe + nhom2.SoNguoiDuocCapThe + nhom3.SoNguoiDuocCapThe + nhom4.SoNguoiDuocCapThe + nhom5.SoNguoiDuocCapThe + nhom6.SoNguoiDuocCapThe
            };

            var KetQuaHuanLuyen = new List<KetQuaHuanLuyenMau5> { nhom1, nhom2, nhom3, nhom4, nhom5, nhom6, tongCong };
            var listThongTinChung = new List<ThongTinChungMau5>();
            var ThongTinChung = new ThongTinChungMau5
            {
                TenDoanhNghiep = txtTenDoanhNghiep.Text,
                DiaDiemBaoCao = txtDiaDiemBaoCao.Text,
                NgayBaoCao = txtNgayBaoCao.Text,
                Nam = lstNamBaoCao5.SelectedItem.ToString(),
                DiaChi = txtDiaChi.Text,
                KinhGui = txtKinhGui.Text,
                DienThoai = txtDienThoai.Text,
                Fax = txtFax.Text,
                Email = txtEmail.Text,
                DiaChiChiNhanh = txtDiaChiChiNhanh.Text,
                HoatDongHuanLuyen = txtHoatDongHuanLuyen.Text,
                DeXuatKienNghi = txtDeXuatKienNghi.Text,
                NgayBaoCaoDD = txtNgayBaoCao.Text.Split('/')[0],
                NgayBaoCaoMM = txtNgayBaoCao.Text.Split('/')[1],
                NgayBaoCaoYYYY = txtNgayBaoCao.Text.Split('/')[2],
            };
            listThongTinChung.Add(ThongTinChung);
            Mau5Template thongTin = new Mau5Template();
            thongTin.Database.Tables["ThongTinChung"].SetDataSource(listThongTinChung);
            thongTin.Database.Tables["KetQuaHuanLuyen"].SetDataSource(KetQuaHuanLuyen);
            FormInBaoCao f = new FormInBaoCao();
            f.crystalReportViewer1.ReportSource = thongTin;
            f.ShowDialog();
        }

        public class ThongTinChungMau5
        {
            public string TenDoanhNghiep { get; set; }
            public string DiaDiemBaoCao { get; set; }
            public string NgayBaoCao { get; set; }
            public string NgayBaoCaoDD { get; set; }
            public string NgayBaoCaoMM { get; set; }
            public string NgayBaoCaoYYYY { get; set; }
            public string Nam { get; set; }
            public string DiaChi { get; set; }
            public string KinhGui { get; set; }
            public string DienThoai { get; set; }
            public string Fax { get; set; }
            public string Email { get; set; }
            public string DiaChiChiNhanh { get; set; }
            public string HoatDongHuanLuyen { get; set; }
            public string DeXuatKienNghi { get; set; }
        }
        public class KetQuaHuanLuyenMau5
        {
            public string DoiTuongHuanLuyen { get; set; }
            public int SoNguoiDuocHuanLuyen { get; set; }
            public int SoNguoiDuocCapCN { get; set; }
            public int SoNguoiDuocCapThe { get; set; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormLopDaoTao f = new FormLopDaoTao();
            this.Hide();
            f.ShowDialog();
        }
    }
}
