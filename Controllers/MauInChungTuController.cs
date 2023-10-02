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
            return View();
        }
        public ActionResult BangDoiChieuCongNo()
        {
            return View();
        }
        public ActionResult ExportToExcel()
        {
            var fileName = $"MauThongBaoNoQH{DateTime.Now.ToString("_yyyyMMddHHmmss")}.xlsx";
            var userDownloadsFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
            var filePath = Path.Combine(userDownloadsFolder, fileName);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("MySheet");

                // Đường dẫn đến hình ảnh trong thư mục 'image'
                var imagePath = "D:/New OPC/Web-OPC-New/Images/opc.png"; // Thay thế bằng đường dẫn thật

                // Chèn hình ảnh từ tệp hình vào ô A1
                ExcelPicture picture = worksheet.Drawings.AddPicture("MyPicture", new FileInfo(imagePath));
                picture.From.Column = 1; // Cột A
                picture.From.Row = 1;    // Dòng 1
                picture.SetSize(30, 30); // Đặt kích thước cho hình ảnh

                package.Save();
            }

            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

            // Tạo đối tượng ContentResult để trả về file Excel và mã JavaScript
            var result = new ContentResult
            {
                Content = $"<!DOCTYPE html><html><head><script>showSuccessAlert();</script></head><body></body></html>",
                ContentType = "text/html",
            };

            // Trả về ContentResult
            return result;
        }



    }
}