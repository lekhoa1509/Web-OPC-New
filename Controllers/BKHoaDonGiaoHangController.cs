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
        public List<BKHoaDonGiaoHang> LoadDmHD(string Ma_dvcs)
        {
            connectSQL();

            //Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_BKHoaDonGiaoHang_SAP]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    //command.Parameters.AddWithValue("@_Ma_Dvcs", Ma_dvcs + "_01");

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            BKHoaDonGiaoHang dataItem = new BKHoaDonGiaoHang
                            {
                                //Ma_Dt = reader["Ma_Dt"].ToString(),
                                //So_Ct = reader["So_Ct"].ToString(),
                                So_Ct_E = reader["So_Ct_E"].ToString(),
                                //Ten_Dt = reader["Ten_Dt"].ToString(),
                              
                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }
        public ActionResult BangKeHoaDonGiaoHang(Account Acc)
        {
            List<BKHoaDonGiaoHang> dmDList = LoadDmHD("");
            ViewBag.DataItems = dmDList;
            return View();
            //DataSet ds = new DataSet();
            //connectSQL();
            //Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            ////string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            //string Pname = "[usp_BKHoaDonGiaoHang_SAP]";
            //ViewBag.ProcedureName = Pname;

            //using (SqlCommand cmd = new SqlCommand(Pname, con))
            //{
            //    cmd.CommandTimeout = 950;

            //    cmd.Connection = con;
            //    cmd.CommandType = CommandType.StoredProcedure;
            //    Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
            //    {

            //        cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
            //        cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
            //        cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs_1);
            //        sda.Fill(ds);

            //    }

            //}
            //return View(ds);
        }
    }
}