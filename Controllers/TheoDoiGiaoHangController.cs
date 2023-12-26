﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using web4.Models;

namespace web4.Controllers
{
    public class TheoDoiGiaoHangController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        // GET: BaoCaoTienVeCN


        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }

        // GET: TheoDoiGiaoHang

        public List<TheoDoiGiaoHang> LoadDmTDV()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            connectSQL();
            
            List<TheoDoiGiaoHang> dataItems = new List<TheoDoiGiaoHang>();

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DanhSachTDV]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Ma_Dvcs", "OPC_HN");

                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[0];

                            foreach (DataRow row in dt.Rows)
                            {
                                TheoDoiGiaoHang dataItem = new TheoDoiGiaoHang
                                {
                                    Ma_NVGH = row["Ma_CbNv"].ToString(),
                                    Ten_NVGH = row["hoten"].ToString(),
                                    Dvcs    = row["Ma_Dvcs"].ToString()
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }


        public List<TheoDoiGiaoHang> LoadHD()
        {
            string ma_dvcs = Request.Cookies["MA_DVCS"] != null ? Request.Cookies["MA_DVCS"].Value : "";
            connectSQL();

            List<TheoDoiGiaoHang> dataItems = new List<TheoDoiGiaoHang>();

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DanhSachHoaDonGiaoHang_SAP]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Ma_Dvcs", "OPC_CT");
                    command.Parameters.AddWithValue("@_Ma_CbNv", "101460");


                    using (SqlDataAdapter sda = new SqlDataAdapter(command))
                    {
                        DataSet ds = new DataSet();
                        sda.Fill(ds);

                        // Kiểm tra xem DataSet có bảng dữ liệu hay không
                        if (ds.Tables.Count > 0)
                        {
                            DataTable dt = ds.Tables[0];

                            foreach (DataRow row in dt.Rows)
                            {
                                TheoDoiGiaoHang dataItem = new TheoDoiGiaoHang
                                {
                                    So_HD = row["so_ct"].ToString(),
                                    Ngay_HD = Convert.ToDateTime(row["Ngay_Ct1"]),
                                    
                                    Ma_Dt = row["Ma_dt"].ToString(),
                                    Ten_Dt = row["Ten_Dt"].ToString(),
                                    Ma_NVGH = row["Ma_nvgh"].ToString(),
                                    Tien_HD = float.Parse(row["tien"].ToString())


                            };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }
        public ActionResult Index()
        {
            
            return View();
        }
        public ActionResult InsertGiaoHang()
        {
           
            return View();
        }
        public ActionResult InsetGiaoHangLoadHD()
        {
            List<TheoDoiGiaoHang> dmDlistTDV = LoadDmTDV();
            List<TheoDoiGiaoHang> dmListHD = LoadHD();
            ViewBag.DataTDV = dmDlistTDV;
            ViewBag.DataHD = dmListHD;
            return View();
        }
    }
}