using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace web4.Models
{
    public class GetData
    {
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
    }
}