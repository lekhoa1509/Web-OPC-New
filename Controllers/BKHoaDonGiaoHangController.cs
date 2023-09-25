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
        public List<BKHoaDonGiaoHang> LoadDmHD(string fromDate, string toDate,string selectedValue)
        {
            string Ma_TDV = selectedValue;
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


        public ActionResult BangKeHoaDonGiaoHang(string selectedValue)
        {
            string fromDate = "20230901"; // Thay đổi giá trị ngày theo nhu cầu
            string toDate = "20230925";   // Thay đổi giá trị đến ngày theo nhu cầu
            //string Ma_TDV = Request.Cookies["Ma_TDV"].Value; // Sử dụng giá trị selectedValue
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
         
            string Ma_TDV = selectedValue;
            System.Diagnostics.Debug.WriteLine("Ma_TDV 1: " + Ma_TDV);
            // Gọi LoadDmHD với Ma_TDV để lấy dữ liệu đã lọc theo Ma_TDV
            List<BKHoaDonGiaoHang> dmDList = LoadDmHD(fromDate, toDate,selectedValue);

            var distinctDataTDV = dmDList
                .GroupBy(x => x.Ten_TDV)
                .Select(x => x.First())
                .ToList();

            var distinctDataItems = dmDList
           .GroupBy(x => x.So_Ct_E)
           .Select(x => x.First())
           .ToList();


            ViewBag.DataTDV = distinctDataTDV;
            ViewBag.DataItems = distinctDataTDV;

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
                    cmd.Parameters.AddWithValue("@_Ma_CbNv", Ma_TDV);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);
                }
            }
            return View();
        }








        public List<BKHoaDonGiaoHang> LoadDmHDWithMaTDV(string selectedValue)
        {
            string fromDate = "20230901"; // Thay đổi giá trị ngày theo nhu cầu
            string toDate = "20230925";
            string Ma_TDV = selectedValue;
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
                    command.Parameters.AddWithValue("@_Ma_CbNv", selectedValue); // Thêm tham số Ma_TDV vào truy vấn

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
            List<BKHoaDonGiaoHang> dmDList = LoadDmHDWithMaTDV(selectedValue);
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            string fromDate = "20230901"; // Thay đổi giá trị ngày theo nhu cầu
            string toDate = "20230925";
            string Ma_TDV = selectedValue;

          
            var distinctDataTDV = dmDList
                .GroupBy(x => x.Ten_TDV)
                .Select(x => x.First())
                .ToList();
            var distinctDataItems = dmDList
    .GroupBy(x => x.So_Ct_E)
    .Select(x => x.First())
    .ToList();
            ViewBag.DataTDV = distinctDataTDV;  
            ViewBag.DataItems = dmDList;

            DataSet ds = new DataSet();
            connectSQL();
            string Pname = "[usp_BKHoaDonGiaoHang_SAP]";

            System.Diagnostics.Debug.WriteLine("Ma TDV la " + Ma_TDV);

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
                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);

                    sda.Fill(ds);
                }
            }

            return View(ds);
        }



    }
}