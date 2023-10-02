using DevExpress.Office.Import.OpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using web4.Models;

namespace web4.Controllers
{
    public class BKHoaDonGiaoHangController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BKHoaDonGiaoHang
        public ActionResult Index()
        {
            return View();
        }
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public List<BKHoaDonGiaoHang> LoadDmHD(string fromDate, string toDate)
        {
         
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            connectSQL();

            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_BKHoaDonGiaoHang_SAP]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    command.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                  // Thêm tham số Ma_TDV vào truy vấn

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            BKHoaDonGiaoHang dataItem = new BKHoaDonGiaoHang
                            {
                                So_Ct_E = reader["So_Ct_E"].ToString(),
                                Ten_TDV = reader["Ten_TDV"].ToString(),
                                Ma_TDV = reader["Ma_TDV"].ToString(),
                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }


        public ActionResult BangKeHoaDonGiaoHang()
        {
            string fromDate = "20230901"; // Thay đổi giá trị ngày theo nhu cầu
            string toDate = "20230926";   // Thay đổi giá trị đến ngày theo nhu cầu
            //string Ma_TDV = Request.Cookies["Ma_TDV"].Value; // Sử dụng giá trị selectedValue
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
         
      
            // Gọi LoadDmHD với Ma_TDV để lấy dữ liệu đã lọc theo Ma_TDV
            List<BKHoaDonGiaoHang> dmDList = LoadDmHD(fromDate, toDate);

            var distinctDataTDV = dmDList
                .GroupBy(x => x.Ten_TDV)
                .Select(x => x.First())
                .ToList();

            var distinctDataItems = dmDList
           .GroupBy(x => x.So_Ct_E)
           .Select(x => x.First())
           .ToList();


            ViewBag.DataTDV = distinctDataTDV;
            ViewBag.DataItems = distinctDataItems;

            DataSet ds = new DataSet();
            connectSQL();
            string Pname = "[usp_BKHoaDonGiaoHang_SAP]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;
                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);
                }
            }
            return View();
        }








        public List<BKHoaDonGiaoHang> LoadDmHDWithMaTDV(string selectedValue)
        {
            string fromDate = "20230901"; // Thay đổi giá trị ngày theo nhu cầu
            string toDate = "20230926";
            string Ma_TDV = Request.Cookies["selectedValue"].Value; 
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            connectSQL();
            System.Diagnostics.Debug.WriteLine("Ma_TDV có trong hàm load là: " + Ma_TDV);
            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_BKHoaDonGiaoHang_SAP]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    command.Parameters.AddWithValue("@_Den_Ngay", toDate);
                    command.Parameters.AddWithValue("@_Ma_CbNv", Ma_TDV); // Thêm tham số Ma_TDV vào truy vấn

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            BKHoaDonGiaoHang dataItem = new BKHoaDonGiaoHang
                            {
                                So_Ct_E = reader["So_Ct_E"].ToString(),
                                Ten_TDV = reader["Ten_TDV"].ToString(),
                                Ma_TDV = reader["Ma_TDV"].ToString(),
                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }
            
            return dataItems;
        }
        public ActionResult BangKeHoaDonGiaoHang_Main(string selectedValue)
        {
            string Ma_TDV = Request.Cookies["selectedValue"].Value; 
            List<BKHoaDonGiaoHang> dmDList = LoadDmHDWithMaTDV(Ma_TDV);
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            string fromDate = "20230901"; // Thay đổi giá trị ngày theo nhu cầu
            string toDate = "20230926";
       

          
            var distinctDataTDV = dmDList
                .GroupBy(x => x.Ten_TDV)
                .Select(x => x.First())
                .ToList();
            var distinctDataItems = dmDList
    .GroupBy(x => x.So_Ct_E)
    .Select(x => x.First())
    .ToList();
            ViewBag.DataTDV = distinctDataTDV;  
            ViewBag.DataItems = distinctDataItems;

            DataSet ds = new DataSet();
            connectSQL();
            string Pname = "[usp_BKHoaDonGiaoHang_SAP]";
                
            System.Diagnostics.Debug.WriteLine("Ma TDV trong controller: " + Ma_TDV);

            // Lọc danh sách dựa trên Ma_TDV
           

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;
                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", fromDate);
                    cmd.Parameters.AddWithValue("@_Den_Ngay",toDate);
                    cmd.Parameters.AddWithValue("@_Ma_CbNv", Ma_TDV);
                    //cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);

                    sda.Fill(ds);
                }
            }

            return View(ds);
        }
        public ActionResult FilterHD(Account Acc)
        {
            DataSet ds = new DataSet();
            string Ma_TDV = Request.Cookies["selectedValue"].Value;
            connectSQL();
            List<BKHoaDonGiaoHang> dmDList = LoadDmHDWithMaTDV(Ma_TDV);
            string Pname = "[usp_BKHoaDonGiaoHang_SAP]";
            Response.Cookies["From_date"].Value = Acc.From_date.ToString();
            Response.Cookies["To_Date"].Value = Acc.To_date.ToString();
            string SelectHD = Request.Cookies["selectedHD"].Value;
            var distinctDataTDV = dmDList
              .GroupBy(x => x.Ten_TDV)
              .Select(x => x.First())
              .ToList();
            var distinctDataItems = dmDList
  .GroupBy(x => x.So_Ct_E)
  .Select(x => x.First())
  .ToList();
            ViewBag.DataTDV = distinctDataTDV;
            ViewBag.DataItems = distinctDataItems;
            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);

                    sda.Fill(ds);

                }
            }

            // Sử dụng LINQ để lọc các giá trị có cột So_Ct_E bằng với SelectHD
            var filteredRows = ds.Tables[0].AsEnumerable()
          .Where(row => row.Field<string>("So_Ct_E") == SelectHD);

            if (filteredRows.Any())
            {
                var filteredDataTable = filteredRows.CopyToDataTable();
                ViewBag.DataSet = filteredDataTable;
            }
            else
            {
                // Không có dữ liệu thỏa mãn điều kiện, bạn có thể xử lý thông báo hoặc thực hiện hành động khác ở đây
                ViewBag.DataSet = new DataTable(); // Tạo DataTable trống để tránh lỗi
            }

            return View("BangKeHoaDonGiaoHang_Main");
        }




    }
}