﻿using System;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Web.Caching;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using web4.Models;
namespace web4.Controllers
{
    public class TheoDoiHDThauController : Controller
    {
        // GET: BaoCao
        SqlConnection con = new SqlConnection();
        SqlCommand sqlc = new SqlCommand();
        SqlDataReader dt;
        public ActionResult Index()
        {
            return View();
        }
        public void connectSQL()
        {
            con.ConnectionString = "Data source= " + "118.69.109.109" + ";database=" + "SAP_OPC" + ";uid=sa;password=Hai@thong";
        }
        public List<BKHoaDonGiaoHang> LoadDmDt()
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            connectSQL();

            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();
            string appendedString = ma_dvcs == "OPC_B1" ? "_010203" : "_01"; // Dòng này cộng chuỗi dựa trên giá trị của Ma_dvcs
            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DmDtTdv_SAP_MauIn]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_Ma_Dvcs", ma_dvcs+ appendedString);

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
                                BKHoaDonGiaoHang dataItem = new BKHoaDonGiaoHang
                                {
                                    Ma_CbNv = row["Ma_Dt"].ToString(),
                                    hoten = row["Ten_Dt"].ToString(),
                                    Ma_Dvcs = row["Dvcs"].ToString()
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }
        public ActionResult TheoDoiHopDongThau_Fill()
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            

            // Gọi LoadDmHD với Ma_TDV để lấy dữ liệu đã lọc theo Ma_TDV
            List<BKHoaDonGiaoHang> dmDList = LoadDmTDV();
            List<BKHoaDonGiaoHang> dmDListDt = LoadDmDt();
            var distinctDataTDV = dmDList
                .GroupBy(x => x.Ten_TDV)
                .Select(x => x.First())
                .ToList();

            var distinctDataItems = dmDList
           .GroupBy(x => x.So_Ct_E)
           .Select(x => x.First())
           .ToList();


            ViewBag.DataTDV = dmDList;
            ViewBag.DataDt = dmDListDt;
            ViewBag.DataItems = distinctDataItems;

            DataSet ds = new DataSet();
            connectSQL();
            string Pname = "[usp_DanhSachTDV]";

            using (SqlCommand cmd = new SqlCommand(Pname, con))
            {
                cmd.CommandTimeout = 950;
                cmd.Connection = con;
                cmd.CommandType = CommandType.StoredProcedure;

                con.Open();

                using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                {

                    cmd.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);
                    sda.Fill(ds);
                }
            }
            return View();
        }
        public List<BKHoaDonGiaoHang> LoadDmTDV()
        {
            string ma_dvcs = Request.Cookies["Ma_dvcs"].Value;
            connectSQL();

            List<BKHoaDonGiaoHang> dataItems = new List<BKHoaDonGiaoHang>();

            using (SqlConnection connection = new SqlConnection(con.ConnectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("[usp_DanhSachTDV]", connection))
                {
                    command.CommandTimeout = 950;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.AddWithValue("@_ma_dvcs", ma_dvcs);

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
                                BKHoaDonGiaoHang dataItem = new BKHoaDonGiaoHang
                                {
                                    Ma_CbNv = row["Ma_CbNv"].ToString(),
                                    hoten = row["hoten"].ToString(),
                                    Ma_Dvcs = row["Ma_Dvcs"].ToString()
                                };

                                dataItems.Add(dataItem);
                            }
                        }
                    }
                }
            }

            return dataItems;
        }
    }
}