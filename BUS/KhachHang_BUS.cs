using DAL;
using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BUS
{
    public class KhachHang_BUS
    {
        DAL.KhachHang_DAL khachhang = new DAL.KhachHang_DAL();
        public DataTable Load_KhachHang()
        {
            return khachhang.Load_KhachHang_DAL();
        }
        public void Insert_KhachHang(KhachHang_DTO ob)
        {
            khachhang.Insert_KhachHang(ob);
        }
        public void Update_KhachHang(KhachHang_DTO ob)
        {
            khachhang.Update_khachHang(ob);
        }
        public void Delete_khachhang(string ID)
        {
            khachhang.Delete_KhachHang(ID);
        }
    }
}
