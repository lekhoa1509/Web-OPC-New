using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using web4.Models;
using System.Web.Mvc;
using System.Data;

namespace web4.Controllers
{
    public class CongTacVienController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoTienVeCN

        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        // GET: CongTacVien
        public ActionResult index()
        {
            return View();
        }
        public List<CTV> LoadDmDt(string Ma_dvcs)
        {
            connectSQL();

            //Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<CTV> dataItems = new List<CTV>();
            string appendedString = Ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DmDtTdv_SAP_MauIn]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@_Ma_Dvcs", "OPC_HN_01");
                    command.CommandTimeout = 950;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CTV dataItem = new CTV
                            {
                               
                                Ma_Dt = reader["ma_dt"].ToString(),
                                Ten_Dt = reader["ten_dt"].ToString(),
                                Dvcs = reader["Dvcs"].ToString()
                              
                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }

        public List<CTV> LoadDmVt()
        {
            connectSQL();

            //Ma_dvcs = Request.Cookies["ma_dvcs"].Value;
            List<CTV> dataItems = new List<CTV>();
            //string appendedString = Ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DanhMucVatTu]", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                   // command.Parameters.AddWithValue("@_Ma_Dvcs", "OPC_HN_01");
                    command.CommandTimeout = 950;
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CTV dataItem = new CTV
                            {

                                Ma_vt = reader["Ma_Vt"].ToString(),
                                ten_vt = reader["Ten_Vt"].ToString()
                               

                            };
                            dataItems.Add(dataItem);
                        }
                    }
                }
            }

            return dataItems;
        }
        public ActionResult InputCTV()
        {
            List<CTV> dmDlist = LoadDmDt("");
            List<CTV> DmVt = LoadDmVt();

            ViewBag.DataItems = dmDlist;
            ViewBag.DataItems2 = DmVt;

            return View();
        }
        public ActionResult SaveCtV(CTV CTV)
        {
            

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();
                try
                {

                    using (SqlCommand command = new SqlCommand("InsertB30CtvData", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.Parameters.AddWithValue("@_Ngay_Ct", CTV.Ngay_Ct);
                        command.Parameters.AddWithValue("@_so_Ct", CTV.So_Ct);
                        command.Parameters.AddWithValue("@_Dvcs", CTV.Dvcs);
                        command.Parameters.AddWithValue("@_Ma_Dt", CTV.Ma_Dt);
                        command.Parameters.AddWithValue("@_Ma_vt", CTV.Ma_vt);
                        command.Parameters.AddWithValue("@_Ten_Vt", CTV.ten_vt);
                        command.Parameters.AddWithValue("@_Han_Muc", CTV.Han_muc);

                        command.ExecuteNonQuery();
                        
                    }
                }
                catch (Exception ex)
                {
                    if(con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                    ex.Message.ToString();
                }
                return View();
            }
           
        }

       
      
        
        //[HttpPost]
        //public ActionResult InputCTV(DateTime Ngay, string So, string DonViCoSo, string DoiTuong, string MaVt, string TenVt, int HanMuc)
        //{
        //    try
        //    {
        //        // Validation (add your validation logic here)

        //        connectSQL();
        //        List<CTV> dmDlist = LoadDmDt("");
        //        List<CTV> DmVt = LoadDmVt();

        //        ViewBag.DataItems = dmDlist;
        //        ViewBag.DataItems2 = DmVt;

        //        using (SqlConnection connection = new SqlConnection(con.ConnectionString))
        //        {
        //            connection.Open();

        //            using (SqlCommand command = new SqlCommand("InsertB30CtvData", connection))
        //            {
        //                command.CommandType = CommandType.StoredProcedure;
        //                command.Parameters.AddWithValue("@_Ngay_Ct", Ngay);
        //                command.Parameters.AddWithValue("@_so_Ct", So);
        //                command.Parameters.AddWithValue("@_Dvcs", DonViCoSo);
        //                command.Parameters.AddWithValue("@_Ma_Dt", DoiTuong);
        //                command.Parameters.AddWithValue("@_Ma_vt", MaVt);
        //                command.Parameters.AddWithValue("@_Ten_Vt", TenVt);
        //                command.Parameters.AddWithValue("@_Han_Muc", HanMuc);

        //                command.ExecuteNonQuery();
        //            }
        //        }

        //        // Success message
        //        ViewBag.Message = "Lưu Thành Công";
        //    }
        //    catch (Exception ex)
        //    {
        //        // Log the exception or handle it as needed
        //        ViewBag.Message = "Save failed: " + ex.Message;
        //    }

        //    return View();
        //}


    }

}