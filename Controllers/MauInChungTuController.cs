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
        public ActionResult ExportToExcel()
        {
            // Lấy dữ liệu từ cookie
            string jsonData = Request.Cookies["tableDataCookie"] != null ? HttpUtility.UrlDecode(Request.Cookies["tableDataCookie"].Value) : "";

            // Kiểm tra xem có dữ liệu từ cookie không
            if (!string.IsNullOrEmpty(jsonData))
            {
                // Parse chuỗi JSON thành mảng JavaScript
                List<List<string>> tableData = JsonConvert.DeserializeObject<List<List<string>>>(jsonData);

                var fileName = $"MauThongBaoNoQH{DateTime.Now:yyyyMMddHHmmss}.xml";
                string userDownloadsFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)+ "\\Downloads";
             
                if (!Directory.Exists(userDownloadsFolder))
                {
                    Directory.CreateDirectory(userDownloadsFolder);
                }
                var filePath = Path.Combine(userDownloadsFolder, fileName);

                // Khởi tạo tệp Excel
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
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

                    // Trả về tệp Excel dưới dạng phản hồi HTTP
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