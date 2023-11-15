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
    public class BangKeHoaDonController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BangKeHoaDon
        public ActionResult Index()
        {
            return View();
        }
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public ActionResult bangkehoadon(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_BKSaleOrder_SAP]";
            Acc.UserName = Request.Cookies["UserName"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
              
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs_1);
                    cmd.Parameters.AddWithValue("@_UserName", Acc.UserName);
                    sda.Fill(ds);

                }
            }


            return View(ds);
           

        }
        public ActionResult bangkehoadon_Fill()
        {
            return View();
        }


        public ActionResult danhsachhoadon_SAP(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            // Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
            //Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DanhSachHoaDon_SAP]";
            Acc.UserName = Response.Cookies["UserName"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs_1 = Request.Cookies["MA_DVCS"].Value;
               
                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs_1);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", Acc.Ma_dt);
                    cmd.Parameters.AddWithValue("@_Tinh_Trang", Acc.Tinh_Trang);
                    cmd.Parameters.AddWithValue("@_username", Acc.UserName);

                    sda.Fill(ds);

                }
            }


            return View(ds);

        }

        public ActionResult danhsachhoadon_SAP_CN(Account Acc)
        {
            DataSet ds = new DataSet();
            connectSQL();
            
            //Acc.UserName = Request.Cookies["UserName"].Value;
            //string query = "exec usp_Vth_BC_BHCNTK_CN @_ngay_Ct1 = '" + Acc.From_date + "',@_Ngay_Ct2 ='"+ Acc.To_date+"',@_Ma_Dvcs='"+ Acc.Ma_DvCs_1+"'";
            string Pname = "[usp_DanhSachHoaDon_SAP]";
            Acc.UserName = Request.Cookies["UserName"].Value;

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;

                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;
                Acc.Ma_DvCs = Request.Cookies["Ma_DvCs_2"].Value;

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_Tu_Ngay", Acc.From_date);
                    cmd.Parameters.AddWithValue("@_Den_Ngay", Acc.To_date);
                    cmd.Parameters.AddWithValue("@_ma_dvcs", Acc.Ma_DvCs);
                    cmd.Parameters.AddWithValue("@_Ma_Dt", Acc.Ma_dt);
                    cmd.Parameters.AddWithValue("@_Tinh_Trang", Acc.Tinh_Trang);
                    cmd.Parameters.AddWithValue("@_username", Acc.UserName);

                    sda.Fill(ds);

                }
            }
            

            return View(ds);

        }
        public ActionResult DanhSachHoaDon_Fill()
        {
            return View();
        }
        public ActionResult DanhSachHoaDon_Fill_HCM()
        {
            return View();
        }
        public ActionResult bangkehoadon_Fill_HCM()
        {
            return View();
        }
    }
}