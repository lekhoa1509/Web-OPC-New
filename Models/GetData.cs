using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
namespace web4.Models
{
    public class GetData
    {
        //Mau So Ton No Chi Tiet Phai Thu
        public int Stt { get; set; }
        public string TenDt { get; set; }
        public string NgayCt { get; set; }
        public string SoCtEinv { get; set; }
        public string NgayDenHan { get; set; }
        public string GhiChu { get; set; }
        public decimal CongNoTT { get; set; }
        public decimal TienThue { get; set; }
        public decimal CongNo { get; set; }
        public decimal TotalCongNoTT { get; set; }
        public decimal TotalCongNoST { get; set; }
        public decimal TotalCongNo { get; set; }
        //Mau Thong Bao No QH
        public string  SoHD { get; set; }
        public string NgayXuat { get; set; }
        public decimal TienNo { get; set; }
        public string HanTT { get; set; }
        public int NgayQH { get; set; }
        //Mau Doi Chieu Cong No
         public string So { get; set; }
        public string Ngay { get; set; }
        public decimal TienHD { get; set; }

    }
}