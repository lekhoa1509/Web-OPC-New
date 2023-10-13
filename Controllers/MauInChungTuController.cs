using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using web4.Models;
using System.Web.Mvc;
using System.Net;
using System.Data.SqlClient;
using System.Data;
using System.Drawing;
using OfficeOpenXml.Drawing;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using OfficeOpenXml.Table;
using Newtonsoft.Json;
namespace web4.Controllers
{
    public class MauInChungTuController : Controller
    {

        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoTienVeCN

        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        // GET: DanhMuc



        public ActionResult Index(MauInChungTu MauIn)
        {
            DataSet ds = new DataSet();
            connectSQL();

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTu_SAP]";
            Response.Cookies["From_date"].Value = MauIn.From_date.ToString();
            Response.Cookies["To_Date"].Value = MauIn.To_date.ToString();
            MauIn.UserName = Request.Cookies["UserName"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_username", MauIn.UserName);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }

        public ActionResult MauInNLCB(MauInChungTu MauIn)
        {
            DataSet ds = new DataSet();
            connectSQL();

            MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTuDetail_SAP]";



            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                MauIn.From_date = Request.Cookies["From_date"].Value;
                MauIn.To_date = Request.Cookies["To_Date"].Value;
                MauIn.UserName = Request.Cookies["UserName"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_So_Ct", MauIn.So_Ct);
                    cmd.Parameters.AddWithValue("@_username", MauIn.So_Ct);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult MauInNLCB_Fill()
        {
            return View();
        }

        public List<MauInChungTu> LoadDmDt(string Ma_dvcs)
        {
            connectSQL();

            Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<MauInChungTu> dataItems = new List<MauInChungTu>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DmDtTdv_SAP_MauIn]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@_Ma_Dvcs", Ma_dvcs + "_01");
                    command.CommandTimeout = 950;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            MauInChungTu dataItem = new MauInChungTu
                            {
                                Ma_Dt = reader["ma_dt"].ToString(),
                                Ten_Dt = reader["ten_dt"].ToString(),
                                Dia_Chi = reader["Dia_Chi"].ToString(),
                                Dvcs = reader["Dvcs"].ToString(),
                                Dvcs1 = reader["Dvcs1"].ToString()
                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }

        public ActionResult PhieuNhapXNTT_Fill()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;
            return View();
        }

        public ActionResult PhieuNhapXNTT(MauInChungTu MauIn)
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            DataSet ds = new DataSet();
            connectSQL();

            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_XacNhanCKTT_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {


                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_Ma_dt", MauIn.Ma_Dt);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult PhieuXuatKho_Fill(MauInChungTu MauIn)
        {
            DataSet ds = new DataSet();
            connectSQL();

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTuSO_Detail_SAP]";

            MauIn.From_date = Request.Cookies["From_date"].Value;
            MauIn.To_date = Request.Cookies["To_Date"].Value;
            MauIn.UserName = Request.Cookies["UserName"].Value;


            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_so_Ct", MauIn.So_Ct);
                    cmd.Parameters.AddWithValue("@_username", MauIn.UserName);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }

        public ActionResult PhieuXuatKho_SO(MauInChungTu MauIn)
        {
            return View();
        }

        public ActionResult PhieuXuatKho_Data(MauInChungTu MauIn)
        {

            DataSet ds = new DataSet();
            connectSQL();

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_MauInChungTuSO_SAP]";
            Response.Cookies["From_date"].Value = MauIn.From_date.ToString();
            Response.Cookies["To_Date"].Value = MauIn.To_date.ToString();
            MauIn.UserName = Request.Cookies["UserName"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_username", MauIn.UserName);
                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult ThongBaoNoQH_Fill()
        {
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;
            return View();
        }
        public ActionResult ThongBaoNoQH_In(MauInChungTu MauIn)
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            DataSet ds = new DataSet();
            connectSQL();

            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_ThongBaoNoQuaHan_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {


                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_Ma_dt", MauIn.Ma_Dt);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }
        public ActionResult BangDoiChieuCongNo_Fill()

        {
            List<MauInChungTu> dmDlist = LoadDmDt("");

            ViewBag.DataItems = dmDlist;
            return View();
            
        }
        public ActionResult BangDoiChieuCongNo( MauInChungTu MauIn)
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            DataSet ds = new DataSet();
            connectSQL();

            //MauIn.So_Ct = Request.Cookies["SoCt"].Value;

            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DoiChieuDoanhThuCongNo_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                //MauIn.From_date = Request.Cookies["From_date"].Value;
                //MauIn.To_date = Request.Cookies["To_Date"].Value;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {


                    cmd.Parameters.AddWithValue("@_Tu_Ngay", MauIn.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", MauIn.To_date);
                    cmd.Parameters.AddWithValue("@_Ma_dt", MauIn.Ma_Dt);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    cmd.Parameters.AddWithValue("@_Ngay_TT", MauIn.Ngay_TT);
                    cmd.Parameters.AddWithValue("@_Ngay_Ky", MauIn.Ngay_Ky);
                    cmd.Parameters.AddWithValue("@_So", MauIn.So);


                    sda.Fill(ds);

                }
            }
            return View(ds);
        }


        //public ActionResult ExportToExcel(List<List<string>> tableData)
        //{
        //    var fileName = $"MauThongBaoNoQH{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        //    var userDownloadsFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
        //    var filePath = Path.Combine(userDownloadsFolder, fileName);

        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //    using (var package = new ExcelPackage(new FileInfo(filePath)))
        //    {
        //        var worksheet = package.Workbook.Worksheets.Add("MySheet");
        //        worksheet.View.ShowGridLines = false;
        //        // Đường dẫn đến hình ảnh trong thư mục 'image'
        //        var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
        //                                                                    // Lấy giá trị từ biến Dvcs
        //        string Dvcs = Request.Cookies["Dvcs"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs"].Value) : "";
        //        string Dvcs1 = Request.Cookies["Dvcs1"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs1"].Value) : "";
        //        string ten_dt = Request.Cookies["ten_dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["ten_dt"].Value) : "";
        //        string denngay = Request.Cookies["DenNgayCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["DenNgayCookie"].Value) : "";
        //        string tongno = Request.Cookies["TongNo"] != null ? HttpUtility.UrlDecode(Request.Cookies["TongNo"].Value) : "";
        //        string QuaHan = Request.Cookies["QuaHan"] != null ? HttpUtility.UrlDecode(Request.Cookies["QuaHan"].Value) : "";
        //        // Đặt font chữ "Arial" cho toàn bộ tệp Excel
        //        worksheet.Cells.Style.Font.Name = "Times New Roman";

        //        // Chèn hình ảnh từ tệp hình vào ô A1
        //        ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
        //        picture.SetSize(55, 45); // Đặt kích thước cho hình ảnh
        //        picture.From.Row = 1;
        //        picture.From.Column = 0;
        //        worksheet.Column(1).Width = 8;

        //        // Đặt văn bản vào ô A2
        //        worksheet.Cells["B1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
        //        var cellB1 = worksheet.Cells["B1"];
        //        cellB1.Style.Font.Bold = true;
        //        worksheet.Cells["B2"].Value = Dvcs;
        //        worksheet.Cells["B3"].Value = $"Số:............................/KT-{Dvcs1}";
        //        worksheet.Cells["I1"].Value = "Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam";
        //        worksheet.Cells["I2"].Value = "Độc Lập - Tự Do - Hạnh Phúc";
        //        worksheet.Cells["I2"].Style.Indent = 4;
        //        worksheet.Cells["I2"].Style.Font.UnderLine = true;
        //        worksheet.Cells["E4"].Value = "THÔNG BÁO NỢ QUÁ HẠN";
        //        worksheet.Cells["E4"].Style.Font.Bold = true;
        //        worksheet.Cells["E4"].Style.Font.Size = 16;
        //        worksheet.Cells["A6"].Value = $"Kính gửi: {ten_dt}";
        //        worksheet.Cells["A6"].Style.Font.Bold = true;
        //        worksheet.Cells["A8"].Value = $"{Dvcs} - Công ty Cổ Phần Dược Phẩm OPC trân trọng thông báo đến quý khách hàng có số dư nợ mà Quý Khách";
        //        worksheet.Cells["A9"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là:";
        //        worksheet.Cells["G9"].Value = $"{tongno}";
        //        worksheet.Cells["G9"].Style.Indent = 0;
        //        worksheet.Cells["G9"].Style.Font.Bold = true;
        //        worksheet.Cells["A11"].Value = "Trong đó nợ quá hạn là: ";
        //        worksheet.Cells["A11"].Style.Indent = 4;
        //        worksheet.Cells["D11"].Value = $"{QuaHan},";
        //        worksheet.Cells["D11"].Style.Font.Bold = true;

        //        // Xử lý dữ liệu bảng từ Ajax
        //        var startRow = 13;
        //        var startColumn = 1;

        //        // Mảng chứa chiều rộng của từng cột (theo đơn vị đo bạn mong muốn)
        //        var columnWidths = new int[] { 7, 20, 20, 20, 20, 20 }; // Ví dụ: Chiều rộng của 6 cột là 5, 10, 15, 12, 15, và 10 đơn vị

        //        for (int row = 0; row < tableData.Count; row++)
        //        {
        //            var rowData = tableData[row];
        //            for (int col = 0; col < rowData.Count; col++)
        //            {
        //                worksheet.Cells[startRow + row, startColumn + col].Value = rowData[col];
        //            }
        //        }

        //        var tableRange = worksheet.Cells[startRow, startColumn, startRow + tableData.Count - 1, startColumn + tableData[0].Count - 1];
        //        var table = worksheet.Tables.Add(tableRange, "MyTable");
        //        table.TableStyle = TableStyles.Light1;

        //        package.Save();
        //    }

        //    byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

        //    // Tạo đối tượng ContentResult để trả về file Excel và mã JavaScript
        //    var result = new FileContentResult(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        //    {
        //        FileDownloadName = fileName
        //    };

        //    // Trả về FileContentResult
        //    return result;
        //}
        //public ActionResult ExportToExcel()
        //{
        //    // Lấy dữ liệu từ cookie
        //    string jsonData = Request.Cookies["tableDataCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie"].Value) : "";

        //    // Kiểm tra xem có dữ liệu từ cookie không
        //    if (!string.IsNullOrEmpty(jsonData))
        //    {
        //        // Parse chuỗi JSON thành mảng JavaScript
        //        List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(jsonData);

        //        //var fileName = $"MauThongBaoNoQH{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        //        //string userDownloadsFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)+ "\\Downloads";

        //        //if (!Directory.Exists(userDownloadsFolder))
        //        //{
        //        //    Directory.CreateDirectory(userDownloadsFolder);
        //        //}
        //        //var filePath = Path.Combine(userDownloadsFolder, fileName);

        //        // Khởi tạo tệp Excel
        //        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //        using (var package = new ExcelPackage())
        //        {
        //            var worksheet = package.Workbook.Worksheets.Add("MySheet");
        //            worksheet.View.ShowGridLines = false;

        //            // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
        //            // Đường dẫn đến hình ảnh trong thư mục 'image'
        //            var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
        //                                                                        // Lấy giá trị từ biến Dvcs
        //            string Dvcs = Request.Cookies["Dvcs"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs"].Value) : "";
        //            string Dvcs1 = Request.Cookies["Dvcs1"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs1"].Value) : "";
        //            string ten_dt = Request.Cookies["ten_dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["ten_dt"].Value) : "";
        //            string denngay = Request.Cookies["DenNgayCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["DenNgayCookie"].Value) : "";
        //            string tongno = Request.Cookies["TongNo"] != null ? HttpUtility.UrlDecode(Request.Cookies["TongNo"].Value) : "";
        //            string QuaHan = Request.Cookies["QuaHan"] != null ? HttpUtility.UrlDecode(Request.Cookies["QuaHan"].Value) : "";
        //            string HanNgay = Request.Cookies["HanNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["HanNgay"].Value) : "";
        //            string CN = Request.Cookies["CN"] != null ? HttpUtility.UrlDecode(Request.Cookies["CN"].Value) : "";
        //            string TK = Request.Cookies["TK"] != null ? HttpUtility.UrlDecode(Request.Cookies["TK"].Value) : "";
        //            string LH = Request.Cookies["LH"] != null ? HttpUtility.UrlDecode(Request.Cookies["LH"].Value) : "";
        //            // Đặt font chữ "Arial" cho toàn bộ tệp Excel
        //            worksheet.Cells.Style.Font.Name = "Times New Roman";

        //            // Chèn hình ảnh từ tệp hình vào ô A1
        //            ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
        //            picture.SetSize(55, 45); // Đặt kích thước cho hình ảnh
        //            picture.From.Row = 1;
        //            picture.From.Column = 0;
        //            worksheet.Column(1).Width = 8;

        //            // Đặt văn bản vào ô A2
        //            worksheet.Cells["B1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
        //            var cellB1 = worksheet.Cells["B1"];
        //            cellB1.Style.Font.Bold = true;
        //            worksheet.Cells["B1"].Style.Indent = 3;
        //            worksheet.Cells["B2"].Style.Indent = 3;
        //            worksheet.Cells["B3"].Style.Indent = 3;
        //            worksheet.Cells["B2"].Value = Dvcs;
        //            worksheet.Cells["B3"].Value = $"Số:............................/KT-{Dvcs1}";
        //            worksheet.Cells["H1"].Value = "Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam";
        //            worksheet.Cells["H2"].Value = "Độc Lập - Tự Do - Hạnh Phúc";
        //            worksheet.Cells["H2"].Style.Indent = 4;
        //            worksheet.Cells["H2"].Style.Font.UnderLine = true;
        //            worksheet.Cells["E4"].Value = "THÔNG BÁO NỢ QUÁ HẠN";
        //            worksheet.Cells["E4"].Style.Font.Bold = true;
        //            worksheet.Cells["E4"].Style.Font.Size = 16;
        //            worksheet.Cells["A6"].Value = $"Kính gửi: {ten_dt}";
        //            worksheet.Cells["A6"].Style.Font.Bold = true;
        //            worksheet.Cells["A8"].Value = $"{Dvcs} - Công ty Cổ Phần Dược Phẩm OPC trân trọng thông báo đến quý khách hàng có số dư nợ mà Quý Khách";
        //            worksheet.Cells["A9"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là: {tongno}";
        //            worksheet.Cells["B11"].Value = $"Trong đó nợ quá hạn là: {QuaHan} bao gồm các hóa đơn sau:";
        //            var startRow = 13;
        //            var startColumn = 1;
        //            worksheet.Cells[startRow - 1, startColumn].Value = "STT";
        //            worksheet.Cells[startRow - 1, startColumn + 1].Value = "SỐ HÓA ĐƠN";
        //            worksheet.Cells[startRow - 1, startColumn + 2].Value = "NGÀY XUẤT";
        //            worksheet.Cells[startRow - 1, startColumn + 3].Value = "TIỀN NỢ";
        //            worksheet.Cells[startRow - 1, startColumn + 4].Value = "HẠN THANH TOÁN";
        //            worksheet.Cells[startRow - 1, startColumn + 5].Value = "NGÀY QUÁ HẠN";
        //            for (int col = 0; col < 6; col++)
        //            {
        //                var columnHeaderCell = worksheet.Cells[startRow - 1, startColumn + col];
        //                columnHeaderCell.Style.Font.Bold = true;
        //                columnHeaderCell.Style.Font.Size = 10;
        //                columnHeaderCell.Style.Font.Color.SetColor(Color.Black);
        //                columnHeaderCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //                columnHeaderCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //                columnHeaderCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                columnHeaderCell.Style.Fill.BackgroundColor.SetColor(Color.White);
        //            }
        //            var columnHeaderStyle = worksheet.Cells[startRow - 1, startColumn, startRow - 1, startColumn + 5].Style;
        //            columnHeaderStyle.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); // Đóng khung solid đen
        //            worksheet.Column(startColumn).Width = 5; // Độ rộng cột cho "STT"
        //            worksheet.Column(startColumn + 1).Width = 15; // Độ rộng cột cho "SỐ HÓA ĐƠN"
        //            worksheet.Column(startColumn + 2).Width = 15; // Độ rộng cột cho "NGÀY XUẤT"
        //            worksheet.Column(startColumn + 3).Width = 15; // Độ rộng cột cho "TIỀN NỢ"
        //            worksheet.Column(startColumn + 4).Width = 18; // Độ rộng cột cho "HẠN THANH TOÁN"
        //            worksheet.Column(startColumn + 5).Width = 15; // 

        //            // Đảm bảo rằng có dữ liệu trong biến tableData
        //            if (tableData != null && tableData.Any())
        //            {
        //                // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
        //                for (int row = 0; row < tableData.Count; row++)
        //                {
        //                    var rowData = tableData[row];
        //                    for (int col = 0; col < rowData.Count; col++)
        //                    {
        //                        worksheet.Cells[startRow + row, startColumn + col].Value = rowData[col];
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
        //            }
        //            worksheet.Cells[startRow + tableData.Count, startColumn + 1].Value = "Tổng cộng";
        //            worksheet.Cells[startRow + tableData.Count, startColumn + 1].Style.Font.Bold = true;
        //            worksheet.Cells[startRow + tableData.Count, startColumn + 3].Value = $"{QuaHan}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
        //            worksheet.Cells[startRow + tableData.Count, startColumn + 3].Style.Font.Bold = true;
        //            int defaultHeaderRowIndex = 13;
        //            // Xóa hàng tiêu đề mặc định
        //            worksheet.DeleteRow(defaultHeaderRowIndex);
        //            var dataRowStyle = worksheet.Cells[startRow, startColumn, startRow, startColumn + 5].Style;
        //            dataRowStyle.Font.Bold = false;
        //            dataRowStyle.Font.Color.SetColor(Color.Black);
        //            dataRowStyle.Fill.PatternType = ExcelFillStyle.None;
        //            // Tạo bảng trong tệp Excel
        //            var endRow = startRow + tableData.Count;
        //            var endColumn = 6;
        //            worksheet.DeleteRow(endRow, 1);
        //            var tableRange = worksheet.Cells[startRow, startColumn, endRow, endColumn];
        //            var table = worksheet.Tables.Add(tableRange, "MyTable");
        //            table.TableStyle = TableStyles.Light1;
        //            int nextRow = endRow + 1;
        //            worksheet.Cells[nextRow, startColumn].Value = $"Kính đề nghị Quý khách vui lòng đối chiếu và xác nhận số tiền gửi về {Dvcs} - Công Ty Cổ Phần Dược Phẩm OPC";
        //            worksheet.Cells[nextRow + 1, startColumn].Value = $"trước ngày {HanNgay}.Đồng thời sớm thanh toán số dư nợ quá hạn trên cho Chi Nhánh chúng tôi bằng tiền mặt hoặc chuyển vào";
        //            worksheet.Cells[nextRow + 2, startColumn].Value = $" tài khoản: Chi nhánh Công Ty Cổ Phẩn Dược Phẩm OPC tại {CN}.";
        //            worksheet.Cells[nextRow + 3, startColumn].Value = $"Số tài khoản: {TK}";
        //            worksheet.Cells[nextRow + 3, startColumn].Style.Indent = 2;
        //            worksheet.Cells[nextRow + 4, startColumn].Value = $"Khi cần đối chiếu xin liên hệ {LH}";
        //            worksheet.Cells[nextRow + 4, startColumn].Style.Indent = 2;
        //            worksheet.Cells[nextRow + 6, startColumn].Value = "Trân trọng!";
        //            worksheet.Cells[nextRow + 6, startColumn].Style.Indent = 2;
        //            worksheet.Cells[nextRow + 6, startColumn].Style.Font.Italic = true;
        //            worksheet.Cells[nextRow + 8, startColumn + 1].Value = "Khách Hàng Xác Nhận";
        //            worksheet.Cells[nextRow + 8, startColumn + 1].Style.Font.Bold = true;
        //            worksheet.Cells[nextRow + 8, startColumn + 4].Value = "Giám Đốc";
        //            worksheet.Cells[nextRow + 8, startColumn + 4].Style.Font.Bold = true;
        //            worksheet.Cells[nextRow + 8, startColumn + 7].Value = "Kế Toán";
        //            worksheet.Cells[nextRow + 8, startColumn + 7].Style.Font.Bold = true;
        //            worksheet.Cells[nextRow + 9, startColumn].Value = "(Ký, đóng dấu, ghi rõ họ tên)";
        //            worksheet.Cells[nextRow + 9, startColumn].Style.Indent = 4;
        //            worksheet.Cells[nextRow + 9, startColumn].Style.Font.Italic = true;

        //            package.Save();
        //            byte[] fileBytes = package.GetAsByteArray();

        //            // Trả về tệp Excel dưới dạng dữ liệu binary
        //            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MauThongBaoNoQH1.xlsx");

        //        }


        //    }
        //    else
        //    {
        //        return Content("Không có dữ liệu từ cookie.");
        //    }
        //    return View("ThongBaoNoQH_In");
        //}
        public ActionResult ExportToExcel()
        {
            var fileName = $"MauThongBaoNoQH{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            // Lấy dữ liệu từ cookie
            string jsonData = Request.Cookies["tableDataCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie"].Value) : "";

            // Kiểm tra xem có dữ liệu từ cookie không
            if (!string.IsNullOrEmpty(jsonData))
            {
                // Parse chuỗi JSON thành mảng JavaScript
                List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(jsonData);

                //var fileName = $"MauThongBaoNoQH{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                //string userDownloadsFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)+ "\\Downloads";

                //if (!Directory.Exists(userDownloadsFolder))
                //{
                //    Directory.CreateDirectory(userDownloadsFolder);
                //}
                //var filePath = Path.Combine(userDownloadsFolder, fileName);

                // Khởi tạo tệp Excel
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("MySheet");
                    worksheet.View.ShowGridLines = false;

                    // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                    // Đường dẫn đến hình ảnh trong thư mục 'image'
                    var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
                                                                                // Lấy giá trị từ biến Dvcs
                    string Dvcs = Request.Cookies["Dvcs"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs"].Value) : "";
                    string Dvcs1 = Request.Cookies["Dvcs1"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs1"].Value) : "";
                    string ten_dt = Request.Cookies["ten_dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["ten_dt"].Value) : "";
                    string denngay = Request.Cookies["DenNgayCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["DenNgayCookie"].Value) : "";
                    string tongno = Request.Cookies["TongNo"] != null ? HttpUtility.UrlDecode(Request.Cookies["TongNo"].Value) : "";
                    string QuaHan = Request.Cookies["QuaHan"] != null ? HttpUtility.UrlDecode(Request.Cookies["QuaHan"].Value) : "";
                    string HanNgay = Request.Cookies["HanNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["HanNgay"].Value) : "";
                    string CN = Request.Cookies["CN"] != null ? HttpUtility.UrlDecode(Request.Cookies["CN"].Value) : "";
                    string TK = Request.Cookies["TK"] != null ? HttpUtility.UrlDecode(Request.Cookies["TK"].Value) : "";
                    string LH = Request.Cookies["LH"] != null ? HttpUtility.UrlDecode(Request.Cookies["LH"].Value) : "";
                    // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                    worksheet.Cells.Style.Font.Name = "Times New Roman";

                    // Chèn hình ảnh từ tệp hình vào ô A1
                    ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
                    picture.SetSize(55, 45); // Đặt kích thước cho hình ảnh
                    picture.From.Row = 1;
                    picture.From.Column = 0;
                    worksheet.Column(1).Width = 8;

                    // Đặt văn bản vào ô A2
                    worksheet.Cells["B1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
                    var cellB1 = worksheet.Cells["B1"];
                    cellB1.Style.Font.Bold = true;
                    worksheet.Cells["B1"].Style.Indent = 3;
                    worksheet.Cells["B2"].Style.Indent = 3;
                    worksheet.Cells["B3"].Style.Indent = 3;
                    worksheet.Cells["B2"].Value = Dvcs;
                    worksheet.Cells["B3"].Value = $"Số:............................/KT-{Dvcs1}";
                    worksheet.Cells["H1"].Value = "Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam";
                    worksheet.Cells["H2"].Value = "Độc Lập - Tự Do - Hạnh Phúc";
                    worksheet.Cells["H2"].Style.Indent = 4;
                    worksheet.Cells["H2"].Style.Font.UnderLine = true;
                    worksheet.Cells["E4"].Value = "THÔNG BÁO NỢ QUÁ HẠN";
                    worksheet.Cells["E4"].Style.Font.Bold = true;
                    worksheet.Cells["E4"].Style.Font.Size = 16;
                    worksheet.Cells["A6"].Value = $"Kính gửi: {ten_dt}";
                    worksheet.Cells["A6"].Style.Font.Bold = true;
                    worksheet.Cells["A8"].Value = $"{Dvcs} - Công ty Cổ Phần Dược Phẩm OPC trân trọng thông báo đến quý khách hàng có số dư nợ mà Quý Khách";
                    worksheet.Cells["A9"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là: {tongno}";
                    worksheet.Cells["B11"].Value = $"Trong đó nợ quá hạn là: {QuaHan} bao gồm các hóa đơn sau:";
                    var startRow = 13;
                    var startColumn = 1;
                    worksheet.Cells[startRow - 1, startColumn].Value = "STT";
                    worksheet.Cells[startRow - 1, startColumn + 1].Value = "SỐ HÓA ĐƠN";
                    worksheet.Cells[startRow - 1, startColumn + 2].Value = "NGÀY XUẤT";
                    worksheet.Cells[startRow - 1, startColumn + 3].Value = "TIỀN NỢ";
                    worksheet.Cells[startRow - 1, startColumn + 4].Value = "HẠN THANH TOÁN";
                    worksheet.Cells[startRow - 1, startColumn + 5].Value = "NGÀY QUÁ HẠN";
                    for (int col = 0; col < 6; col++)
                    {
                        var columnHeaderCell = worksheet.Cells[startRow - 1, startColumn + col];
                        columnHeaderCell.Style.Font.Bold = true;
                        columnHeaderCell.Style.Font.Size = 10;
                        columnHeaderCell.Style.Font.Color.SetColor(Color.Black);
                        columnHeaderCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        columnHeaderCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        columnHeaderCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        columnHeaderCell.Style.Fill.BackgroundColor.SetColor(Color.White);
                    }
                    var columnHeaderStyle = worksheet.Cells[startRow - 1, startColumn, startRow - 1, startColumn + 5].Style;
                    columnHeaderStyle.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); // Đóng khung solid đen
                    worksheet.Column(startColumn).Width = 5; // Độ rộng cột cho "STT"
                    worksheet.Column(startColumn + 1).Width = 15; // Độ rộng cột cho "SỐ HÓA ĐƠN"
                    worksheet.Column(startColumn + 2).Width = 15; // Độ rộng cột cho "NGÀY XUẤT"
                    worksheet.Column(startColumn + 3).Width = 15; // Độ rộng cột cho "TIỀN NỢ"
                    worksheet.Column(startColumn + 4).Width = 18; // Độ rộng cột cho "HẠN THANH TOÁN"
                    worksheet.Column(startColumn + 5).Width = 15; // 

                    // Đảm bảo rằng có dữ liệu trong biến tableData
                    if (tableData != null && tableData.Any())
                    {
                        // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                        for (int row = 0; row < tableData.Count; row++)
                        {
                            var rowData = tableData[row];
                            for (int col = 0; col < rowData.Count; col++)
                            {
                                worksheet.Cells[startRow + row, startColumn + col].Value = rowData[col];
                            }
                        }
                    }
                    else
                    {
                        worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                    }
                    worksheet.Cells[startRow + tableData.Count, startColumn + 1].Value = "Tổng cộng";
                    worksheet.Cells[startRow + tableData.Count, startColumn + 1].Style.Font.Bold = true;
                    worksheet.Cells[startRow + tableData.Count, startColumn + 3].Value = $"{QuaHan}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                    worksheet.Cells[startRow + tableData.Count, startColumn + 3].Style.Font.Bold = true;
                    int defaultHeaderRowIndex = 13;
                    // Xóa hàng tiêu đề mặc định
                    worksheet.DeleteRow(defaultHeaderRowIndex);
                    var dataRowStyle = worksheet.Cells[startRow, startColumn, startRow, startColumn + 5].Style;
                    dataRowStyle.Font.Bold = false;
                    dataRowStyle.Font.Color.SetColor(Color.Black);
                    dataRowStyle.Fill.PatternType = ExcelFillStyle.None;
                    // Tạo bảng trong tệp Excel
                    var endRow = startRow + tableData.Count;
                    var endColumn = 6;
                    worksheet.DeleteRow(endRow, 1);
                    var tableRange = worksheet.Cells[startRow, startColumn, endRow, endColumn];
                    var table = worksheet.Tables.Add(tableRange, "MyTable");
                    table.TableStyle = TableStyles.Light1;
                    int nextRow = endRow + 1;
                    worksheet.Cells[nextRow, startColumn].Value = $"Kính đề nghị Quý khách vui lòng đối chiếu và xác nhận số tiền gửi về {Dvcs} - Công Ty Cổ Phần Dược Phẩm OPC";
                    worksheet.Cells[nextRow + 1, startColumn].Value = $"trước ngày {HanNgay}.Đồng thời sớm thanh toán số dư nợ quá hạn trên cho Chi Nhánh chúng tôi bằng tiền mặt hoặc chuyển vào";
                    worksheet.Cells[nextRow + 2, startColumn].Value = $" tài khoản: Chi nhánh Công Ty Cổ Phẩn Dược Phẩm OPC tại {CN}.";
                    worksheet.Cells[nextRow + 3, startColumn].Value = $"Số tài khoản: {TK}";
                    worksheet.Cells[nextRow + 3, startColumn].Style.Indent = 2;
                    worksheet.Cells[nextRow + 4, startColumn].Value = $"Khi cần đối chiếu xin liên hệ {LH}";
                    worksheet.Cells[nextRow + 4, startColumn].Style.Indent = 2;
                    worksheet.Cells[nextRow + 6, startColumn].Value = "Trân trọng!";
                    worksheet.Cells[nextRow + 6, startColumn].Style.Indent = 2;
                    worksheet.Cells[nextRow + 6, startColumn].Style.Font.Italic = true;
                    worksheet.Cells[nextRow + 8, startColumn + 1].Value = "Khách Hàng Xác Nhận";
                    worksheet.Cells[nextRow + 8, startColumn + 1].Style.Font.Bold = true;
                    worksheet.Cells[nextRow + 8, startColumn + 4].Value = "Giám Đốc";
                    worksheet.Cells[nextRow + 8, startColumn + 4].Style.Font.Bold = true;
                    worksheet.Cells[nextRow + 8, startColumn + 7].Value = "Kế Toán";
                    worksheet.Cells[nextRow + 8, startColumn + 7].Style.Font.Bold = true;
                    worksheet.Cells[nextRow + 9, startColumn].Value = "(Ký, đóng dấu, ghi rõ họ tên)";
                    worksheet.Cells[nextRow + 9, startColumn].Style.Indent = 4;
                    worksheet.Cells[nextRow + 9, startColumn].Style.Font.Italic = true;

                    package.Save();
                    byte[] fileBytes = package.GetAsByteArray();

                    // Trả về tệp Excel dưới dạng dữ liệu binary
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

                }


            }
            else
            {
                return Content("Không có dữ liệu từ cookie.");
            }
            return View("ThongBaoNoQH_In");
        }

        public ActionResult ExportBaoCaoCongNo()
        {

            var fileName = $"BangDoiChieuCongNo{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            // Lấy dữ liệu từ cookie
            string jsonData = Request.Cookies["tableDataCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie"].Value) : "";
            string jsonData2 = Request.Cookies["tableData2Cookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableData2Cookie"].Value) : "";
            // Kiểm tra xem có dữ liệu từ cookie không
            if (!string.IsNullOrEmpty(jsonData) &&!string.IsNullOrEmpty(jsonData2))
            {
                // Parse chuỗi JSON thành mảng JavaScript
                List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(jsonData);
                List<List<string>> tableData2 = JsonConvert.DeserializeObject<List<List<string>>>(jsonData2);

                //var fileName = $"MauThongBaoNoQH{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                //string userDownloadsFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)+ "\\Downloads";

                //if (!Directory.Exists(userDownloadsFolder))
                //{
                //    Directory.CreateDirectory(userDownloadsFolder);
                //}
                //var filePath = Path.Combine(userDownloadsFolder, fileName);

                // Khởi tạo tệp Excel
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("MySheet");
                    worksheet.View.ShowGridLines = false;
                    var startRow = 13;
                    var startColumn = 1;
                    // ... (Các bước tạo nội dung tệp Excel như bạn đã làm)
                    // Đường dẫn đến hình ảnh trong thư mục 'image'
                    var imagePath = Server.MapPath("~/assets/images/logo.png"); // Thay thế bằng đường dẫn thật
                                                                                // Lấy giá trị từ biến Dvcs
                    string Dvcs = Request.Cookies["Dvcs"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs"].Value) : "";
                    string TuNgay = Request.Cookies["tungay"] != null ? HttpUtility.UrlDecode(Request.Cookies["tungay"].Value) : "";
                    string TuThang = Request.Cookies["tuthang"] != null ? HttpUtility.UrlDecode(Request.Cookies["tuthang"].Value) : "";
                    string DenNgay = Request.Cookies["denngay"] != null ? HttpUtility.UrlDecode(Request.Cookies["denngay"].Value) : "";
                    string DenThang = Request.Cookies["denthang"] != null ? HttpUtility.UrlDecode(Request.Cookies["denthang"].Value) : "";
                    string Nam = Request.Cookies["nam"] != null ? HttpUtility.UrlDecode(Request.Cookies["nam"].Value) : "";
                    string DiaChi = Request.Cookies["Dia_Chi"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dia_Chi"].Value) : "";
                    string NoDauKy = Request.Cookies["NoDauKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NoDauKy"].Value) : "";
                    string TienHD = Request.Cookies["TienHD"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienHD"].Value) : "";
                    string TonNo = Request.Cookies["TonNo"] != null ? HttpUtility.UrlDecode(Request.Cookies["TonNo"].Value) : "";
                    string TienChu = Request.Cookies["TienChu"] != null ? HttpUtility.UrlDecode(Request.Cookies["TienChu"].Value) : "";
                    string TonNo2 = Request.Cookies["TonNo2"] != null ? HttpUtility.UrlDecode(Request.Cookies["TonNo2"].Value) : "";
                    string NgayTT = Request.Cookies["NgayTT"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayTT"].Value) : "";
                    string ChiNhanh = Request.Cookies["ChiNhanh"] != null ? HttpUtility.UrlDecode(Request.Cookies["ChiNhanh"].Value) : "";
                    string DiaChi2 = Request.Cookies["DiaChi2"] != null ? HttpUtility.UrlDecode(Request.Cookies["DiaChi2"].Value) : "";
                    //string Dvcs1 = Request.Cookies["Dvcs1"] != null ? HttpUtility.UrlDecode(Request.Cookies["Dvcs1"].Value) : "";
                    string ten_kh = Request.Cookies["Ten_Dt"] != null ? HttpUtility.UrlDecode(Request.Cookies["Ten_Dt"].Value) : "";
                    string NgayKy = Request.Cookies["NgayKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NgayKy"].Value) : "";
                    string ThangKy = Request.Cookies["ThangKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["ThangKy"].Value) : "";
                    string NamKy = Request.Cookies["NamKy"] != null ? HttpUtility.UrlDecode(Request.Cookies["NamKy"].Value) : "";
                    //string HanNgay = Request.Cookies["HanNgay"] != null ? HttpUtility.UrlDecode(Request.Cookies["HanNgay"].Value) : "";
                    string CN = Request.Cookies["CN"] != null ? HttpUtility.UrlDecode(Request.Cookies["CN"].Value) : "";
                    string TK = Request.Cookies["TK"] != null ? HttpUtility.UrlDecode(Request.Cookies["TK"].Value) : "";
                    string LH = Request.Cookies["LH"] != null ? HttpUtility.UrlDecode(Request.Cookies["LH"].Value) : "";
                    // Đặt font chữ "Arial" cho toàn bộ tệp Excel
                    worksheet.Cells.Style.Font.Name = "Times New Roman";

                    // Chèn hình ảnh từ tệp hình vào ô A1
                    ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
                    picture.SetSize(70, 50); // Đặt kích thước cho hình ảnh
                    picture.From.Row = 1;
                    picture.From.Column = 0;
                
                    worksheet.Column(1).Width = 8;

                    // Đặt văn bản vào ô A2
                    worksheet.Cells["B1"].Value = "CTY CỔ PHẦN DƯỢC PHẨM OPC";
                    var cellB1 = worksheet.Cells["B1"];
                    cellB1.Style.Font.Bold = true;
                    //worksheet.Cells["B1"].Style.Indent = 3;
                    //worksheet.Cells["B2"].Style.Indent = 3;
                    //worksheet.Cells["B3"].Style.Indent = 3;
                    worksheet.Cells["B2"].Value = Dvcs;
                    worksheet.Cells["B3"].Value = $"Số:";
                    worksheet.Cells["H1"].Value = "Cộng Hòa Xã Hội Chủ Nghĩa Việt Nam";
                    worksheet.Cells["H2"].Value = "Độc Lập - Tự Do - Hạnh Phúc";
                    worksheet.Cells["H2"].Style.Indent = 4;

                    worksheet.Cells["E4"].Value = "BẢNG ĐỐI CHIẾU DOANH THU CÔNG NỢ";
                    worksheet.Cells["E4"].Style.Font.Bold = true;
                    worksheet.Cells["E4"].Style.Font.Size = 16;
                    worksheet.Cells["E5"].Value = $"Từ ngày {TuNgay} tháng {TuThang} đến ngày {DenNgay} tháng {DenThang} năm {Nam}";
                    worksheet.Cells["E5"].Style.Indent = 4;
                    worksheet.Cells["A7"].Value = $"Tên khách hàng: {ten_kh}";
                    worksheet.Cells["A7"].Style.Font.Bold = true;
                    worksheet.Cells["A8"].Value = $"Địa chỉ khách hàng: {DiaChi}";
                    worksheet.Cells["A8"].Style.Font.Bold = true;
                    worksheet.Cells["A9"].Value = $"I.Số dư nợ trước ngày: {TuNgay}/{TuThang}/{Nam}";
                    worksheet.Cells["A9"].Style.Font.Bold = true;
                    worksheet.Cells["E9"].Value = $"mang sang {NoDauKy} đồng";
                    worksheet.Cells["E9"].Style.Font.Bold = true;
                    worksheet.Cells["A10"].Value = "II.Doanh thu và công nợ phát sinh trong kỳ đối chiếu này: ";
                    worksheet.Cells["A10"].Style.Font.Bold = true;
                    //worksheet.Cells["A8"].Value = $"{Dvcs} - Công ty Cổ Phần Dược Phẩm OPC trân trọng thông báo đến quý khách hàng có số dư nợ mà Quý Khách";
                    //worksheet.Cells["A9"].Value = $"hàng chưa thanh toán cho chúng tôi tính đến ngày {denngay} là: {tongno}";
                    //worksheet.Cells["B11"].Value = $"Trong đó nợ quá hạn là: {QuaHan} bao gồm các hóa đơn sau:";
                  
                    var sttCell = worksheet.Cells[startRow - 1, startColumn, startRow, startColumn];
                    sttCell.Merge = true;
                    sttCell.Value = "STT";
                    sttCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sttCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                    sttCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
                    sttCell.Style.Font.Bold = true;
                    sttCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    var KHMua = worksheet.Cells[startRow - 1, startColumn + 1, startRow - 1, startColumn + 3];
                    KHMua.Merge = true;
                    KHMua.Value = "KHÁCH HÀNG MUA";
                    KHMua.Style.Font.Bold = true;
                    KHMua.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    KHMua.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    KHMua.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 1].Value = "SỐ";
                    worksheet.Cells[startRow, startColumn + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 1].Style.Font.Bold = true;
                    worksheet.Cells[startRow, startColumn + 2].Value = "NGÀY";
                    worksheet.Cells[startRow, startColumn + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 2].Style.Font.Bold = true;
                    worksheet.Cells[startRow, startColumn + 3].Value = "SỐ TIỀN";
                    worksheet.Cells[startRow, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 3].Style.Font.Bold = true;
                    var KHTT = worksheet.Cells[startRow - 1, startColumn + 4, startRow - 1, startColumn + 8];
                    KHTT.Merge = true;
                    KHTT.Value = "KHÁCH HÀNG THANH TOÁN/TRẢ HÀNG BÙ TRỪ";
                    KHTT.Style.Font.Bold = true;
                    KHTT.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    //worksheet.Column(startColumn + 4).Width = 20;
                    worksheet.Cells[startRow - 1, startColumn + 9].Value = "";
                    worksheet.Cells[startRow - 1, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 4].Value = "SỐ";
                    worksheet.Cells[startRow, startColumn + 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 4].Style.Font.Bold = true;

                    worksheet.Cells[startRow, startColumn + 5].Value = "NGÀY";
                    worksheet.Cells[startRow, startColumn + 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 5].Style.Font.Bold = true;
                    worksheet.Cells[startRow, startColumn + 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells[startRow, startColumn + 6].Value = "SỐ TIỀN";
                    worksheet.Cells[startRow, startColumn + 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 6].Style.Font.Bold = true;

                    worksheet.Cells[startRow, startColumn + 7].Value = "CKTT";
                    worksheet.Cells[startRow, startColumn + 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 7].Style.Font.Bold = true;

                    worksheet.Cells[startRow, startColumn + 8].Value = "TỔNG TIỀN";
                    worksheet.Cells[startRow, startColumn + 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 8].Style.Font.Bold = true;
                    worksheet.Cells[startRow, startColumn + 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells[startRow, startColumn + 9].Value = "GHI CHÚ";
                    worksheet.Cells[startRow, startColumn + 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 9].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRow, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow, startColumn + 9].Style.Font.Bold = true;
                    if (tableData != null && tableData.Any())
                    {
                        // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                        for (int row = 0; row < tableData.Count; row++)
                        {
                            var rowData = tableData[row];
                            for (int col = 0; col < rowData.Count; col++)
                            {
                                if(col == 4 &&col ==7 &&col==9)
                                {
                                    worksheet.Cells[startRow - 1 + row, startColumn + col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                    worksheet.Cells[startRow - 1 + row, startColumn + col].Value = rowData[col];
                                    worksheet.Cells[startRow - 1 + row, startColumn + col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                                }
                                else
                                {
                                    worksheet.Cells[startRow - 1 + row, startColumn + col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[startRow - 1 + row, startColumn + col].Value = rowData[col];
                                    worksheet.Cells[startRow - 1 + row, startColumn + col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                                }

                               
                            }
                        }
                    }
                    else
                    {
                        worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                    }
                    var TC = worksheet.Cells[startRow - 1 + tableData.Count, startColumn, startRow - 1 + tableData.Count, startColumn + 2];
                    TC.Merge = true;
                    TC.Value = "Tổng cộng: ";
                    TC.Style.Font.Bold = true;
                    TC.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    TC.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                    TC.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 3].Value = $"{TienHD}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 3].Style.Font.Bold = true;
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRow - 1 + tableData.Count, startColumn + 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                    var endII = startRow + tableData.Count;
                    var nextII = endII + 1;
                    worksheet.Cells[nextII, startColumn].Value = $"III. Số tiền khách hàng chưa thanh toán, tính đến cuối ngày: {DenNgay}/{DenThang}/{Nam} là: {TonNo} đồng";
                    worksheet.Cells[nextII, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[nextII + 1, startColumn].Value = $"Số tiền bằng chữ: {TienChu}";
                    worksheet.Cells[nextII + 1, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[nextII + 2, startColumn].Value = "Chi tiết các hóa đơn chưa thanh toán: ";

                    var startRowIII = nextII + 3;

                    var sttIIICell = worksheet.Cells[startRowIII + 1, startColumn, startRowIII + 2, startColumn];
                    sttIIICell.Merge = true;
                    worksheet.Column(startColumn).Width = 15; // Đặt chiều rộng của cột chứa ô "STT" thành 15 đơn vị.

                    sttIIICell.Value = "STT";
                    sttIIICell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                    sttIIICell.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc
                    sttIIICell.Style.Font.Bold = true;
                    sttIIICell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                    var HD = worksheet.Cells[startRowIII + 1, startColumn + 1, startRowIII + 1, startColumn + 4];
                    HD.Merge = true;
                    HD.Value = "HÓA ĐƠN";
                    HD.Style.Font.Bold = true;
                    HD.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    HD.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    HD.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                    worksheet.Cells[startRowIII + 2, startColumn + 1].Value = "SỐ";
                    worksheet.Column(startColumn + 1).Width = 15;
                    worksheet.Cells[startRowIII + 2, startColumn + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRowIII + 2, startColumn + 1].Style.Font.Bold = true;
                    worksheet.Cells[startRowIII + 2, startColumn + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRowIII + 2, startColumn + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells[startRowIII + 2, startColumn + 2].Value = "NGÀY";
                    worksheet.Column(startColumn + 2).Width = 15;
                    worksheet.Cells[startRowIII + 2, startColumn + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRowIII + 2, startColumn + 2].Style.Font.Bold = true;
                    worksheet.Cells[startRowIII + 2, startColumn + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRowIII + 2, startColumn + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Cells[startRowIII + 2, startColumn + 3].Value = "SỐ TIỀN HD";
                    worksheet.Column(startColumn + 3).Width = 15;
                    worksheet.Cells[startRowIII + 2, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRowIII + 2, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRowIII + 2, startColumn + 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRowIII + 2, startColumn + 3].Style.Font.Bold = true;

                    worksheet.Cells[startRowIII + 2, startColumn + 4].Value = "GHI CHÚ";
                    worksheet.Column(startColumn + 4).Width = 15;
                    worksheet.Column(startColumn + 5).Width = 15;
                    worksheet.Column(startColumn + 6).Width = 15;
                    worksheet.Column(startColumn + 8).Width = 15;
                    worksheet.Cells[startRowIII + 2, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRowIII + 2, startColumn + 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRowIII + 2, startColumn + 4].Style.Font.Bold = true;
                    worksheet.Cells[startRowIII + 2, startColumn + 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    if (tableData2 != null && tableData2.Any())
                    {
                        // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                        for (int row = 0; row < tableData2.Count; row++)
                        {
                            var rowData = tableData2[row];
                            for (int col = 0; col < rowData.Count; col++)
                            {
                                worksheet.Cells[startRowIII + 1 + row, startColumn + col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[startRowIII + 1 + row, startColumn + col].Value = rowData[col];
                                worksheet.Cells[startRowIII + 1 + row, startColumn + col].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                            }
                        }
                    }
                    else
                    {
                        worksheet.Cells[startRowIII, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                    }
                    var TC2 = worksheet.Cells[startRowIII + 1 + tableData2.Count, startColumn, startRowIII + 1 + tableData2.Count, startColumn + 2];
                    TC2.Merge = true;
                    TC2.Value = "Tổng cộng: ";
                    TC2.Style.Font.Bold = true;
                    TC2.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    TC2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Đặt canh giữa ngang
                    TC2.Style.VerticalAlignment = ExcelVerticalAlignment.Center; // Đặt canh giữa dọc

                    worksheet.Cells[startRowIII + 1 + tableData2.Count, startColumn + 3].Value = $"{TonNo2}";
                    worksheet.Cells[startRowIII + 1 + tableData2.Count, startColumn + 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[startRowIII + 1 + tableData2.Count, startColumn + 3].Style.Font.Bold = true;
                    worksheet.Cells[startRowIII + 1 + tableData2.Count, startColumn + 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    worksheet.Cells[startRowIII + 1 + tableData2.Count, startColumn + 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                    var endRowIII = startRowIII + 1 + tableData2.Count + 1;
                    worksheet.Cells[endRowIII, startColumn].Value = $"Xin vui lòng xác nhận và gửi lại cho {Dvcs} trước ngày {NgayTT}";
                    worksheet.Cells[endRowIII, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 1, startColumn].Value = $"Nơi nhận: {ChiNhanh}";
                    worksheet.Cells[endRowIII + 1, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 2, startColumn].Value = $"Địa chỉ: {DiaChi2}";
                    worksheet.Cells[endRowIII + 2, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 3, startColumn].Value = $"Khi cần đối chiếu số liệu liên hệ: {LH}";
                    worksheet.Cells[endRowIII + 3, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 4, startColumn].Value = $"Số tiền còn nợ đề nghị Quý khách hàng thanh toán bằng tiền mặt hoặc chuyển khoản vào tài khoản {CN}, số";
                    worksheet.Cells[endRowIII + 4, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 5, startColumn].Value = $"tài khoản: {TK}";
                    worksheet.Cells[endRowIII + 5, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 7, startColumn].Value = "Trân trọng cảm ơn!";
                    worksheet.Cells[endRowIII + 7, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 7, startColumn].Style.Font.Italic = true;
                    worksheet.Cells[endRowIII + 8, startColumn + 7].Value = $"Ngày {NgayKy} tháng {ThangKy} năm {NamKy}";
                    worksheet.Cells[endRowIII + 8, startColumn + 7].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 9, startColumn].Value = "ĐẠI DIỆN KHÁCH HÀNG";
                    worksheet.Cells[endRowIII + 9, startColumn].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 9, startColumn+7].Value = "ĐẠI DIỆN CHI NHÁNH";
                    worksheet.Cells[endRowIII + 9, startColumn+7].Style.Font.Bold = true;
                    worksheet.Cells[endRowIII + 9, startColumn + 7].Style.Indent = 2;
                    
                    //for (int col = 0; col < 6; col++)
                    //{
                    //    var columnHeaderCell = worksheet.Cells[startRow - 1, startColumn + col];
                    //    columnHeaderCell.Style.Font.Bold = true;
                    //    columnHeaderCell.Style.Font.Size = 10;
                    //    columnHeaderCell.Style.Font.Color.SetColor(Color.Black);
                    //    columnHeaderCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //    columnHeaderCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    //    columnHeaderCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //    columnHeaderCell.Style.Fill.BackgroundColor.SetColor(Color.White);
                    //}
                    //var columnHeaderStyle = worksheet.Cells[startRow - 1, startColumn, startRow - 1, startColumn + 5].Style;
                    //columnHeaderStyle.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); // Đóng khung solid đen
                    //worksheet.Column(startColumn).Width = 5; // Độ rộng cột cho "STT"
                    //worksheet.Column(startColumn + 1).Width = 15; // Độ rộng cột cho "SỐ HÓA ĐƠN"
                    //worksheet.Column(startColumn + 2).Width = 15; // Độ rộng cột cho "NGÀY XUẤT"
                    //worksheet.Column(startColumn + 3).Width = 15; // Độ rộng cột cho "TIỀN NỢ"
                    //worksheet.Column(startColumn + 4).Width = 18; // Độ rộng cột cho "HẠN THANH TOÁN"
                    //worksheet.Column(startColumn + 5).Width = 15; // 

                    // Đảm bảo rằng có dữ liệu trong biến tableData
                    //if (tableData != null && tableData.Any())
                    //{
                    //    // Lặp qua từng hàng dữ liệu trong tableData và ghi vào tệp Excel
                    //    for (int row = 0; row < tableData.Count; row++)
                    //    {
                    //        var rowData = tableData[row];
                    //        for (int col = 0; col < rowData.Count; col++)
                    //        {
                    //            worksheet.Cells[startRow + row, startColumn + col].Value = rowData[col];
                    //        }
                    //    }
                    //}
                    //else
                    //{
                    //    worksheet.Cells[startRow, startColumn].Value = "Không có dữ liệu bảng từ cookie.";
                    //}
                    //worksheet.Cells[startRow + tableData.Count, startColumn + 1].Value = "Tổng cộng";
                    //worksheet.Cells[startRow + tableData.Count, startColumn + 1].Style.Font.Bold = true;
                    //worksheet.Cells[startRow + tableData.Count, startColumn + 3].Value = $"{QuaHan}"; // Ví dụ: Ghi giá trị tổng vào cột thứ 4
                    //worksheet.Cells[startRow + tableData.Count, startColumn + 3].Style.Font.Bold = true;
                    //int defaultHeaderRowIndex = 13;
                    //// Xóa hàng tiêu đề mặc định
                    //worksheet.DeleteRow(defaultHeaderRowIndex);
                    //var dataRowStyle = worksheet.Cells[startRow, startColumn, startRow, startColumn + 5].Style;
                    //dataRowStyle.Font.Bold = false;
                    //dataRowStyle.Font.Color.SetColor(Color.Black);
                    //dataRowStyle.Fill.PatternType = ExcelFillStyle.None;
                    // Tạo bảng trong tệp Excel
                    //var endRow = startRow + tableData.Count;
                    //var endColumn = 6;
                    //worksheet.DeleteRow(endRow, 1);
                    //var tableRange = worksheet.Cells[startRow, startColumn, endRow, endColumn];
                    //var table = worksheet.Tables.Add(tableRange, "MyTable");
                    //table.TableStyle = TableStyles.Light1;
                    //int nextRow = endRow + 1;
                    //worksheet.Cells[nextRow, startColumn].Value = $"Kính đề nghị Quý khách vui lòng đối chiếu và xác nhận số tiền gửi về {Dvcs} - Công Ty Cổ Phần Dược Phẩm OPC";
                    //worksheet.Cells[nextRow + 1, startColumn].Value = $"trước ngày {HanNgay}.Đồng thời sớm thanh toán số dư nợ quá hạn trên cho Chi Nhánh chúng tôi bằng tiền mặt hoặc chuyển vào";
                    //worksheet.Cells[nextRow + 2, startColumn].Value = $" tài khoản: Chi nhánh Công Ty Cổ Phẩn Dược Phẩm OPC tại {CN}.";
                    //worksheet.Cells[nextRow + 3, startColumn].Value = $"Số tài khoản: {TK}";
                    //worksheet.Cells[nextRow + 3, startColumn].Style.Indent = 2;
                    //worksheet.Cells[nextRow + 4, startColumn].Value = $"Khi cần đối chiếu xin liên hệ {LH}";
                    //worksheet.Cells[nextRow + 4, startColumn].Style.Indent = 2;
                    //worksheet.Cells[nextRow + 6, startColumn].Value = "Trân trọng!";
                    //worksheet.Cells[nextRow + 6, startColumn].Style.Indent = 2;
                    //worksheet.Cells[nextRow + 6, startColumn].Style.Font.Italic = true;
                    //worksheet.Cells[nextRow + 8, startColumn + 1].Value = "Khách Hàng Xác Nhận";
                    //worksheet.Cells[nextRow + 8, startColumn + 1].Style.Font.Bold = true;
                    //worksheet.Cells[nextRow + 8, startColumn + 4].Value = "Giám Đốc";
                    //worksheet.Cells[nextRow + 8, startColumn + 4].Style.Font.Bold = true;
                    //worksheet.Cells[nextRow + 8, startColumn + 7].Value = "Kế Toán";
                    //worksheet.Cells[nextRow + 8, startColumn + 7].Style.Font.Bold = true;
                    //worksheet.Cells[nextRow + 9, startColumn].Value = "(Ký, đóng dấu, ghi rõ họ tên)";
                    //worksheet.Cells[nextRow + 9, startColumn].Style.Indent = 4;
                    //worksheet.Cells[nextRow + 9, startColumn].Style.Font.Italic = true;

                    package.Save();
                    byte[] fileBytes = package.GetAsByteArray();

                    // Trả về tệp Excel dưới dạng dữ liệu binary
                    return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

                }


            }
            else
            {
                return Content("Không có dữ liệu từ cookie.");
            }
            return View("ThongBaoNoQH_In");

        }

    }
}