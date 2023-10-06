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
    public class NhanVien_BUS
    {
        NhanVien_DAL nhanvien = new NhanVien_DAL();
        public DataTable Load_Nhanvien()
        {
            return nhanvien.Load_NhanVien_DAL();
        }
        public void Insert_KhachHang(NhanVien_DTO ob)
        {
            nhanvien.Insert_NhanVien(ob);
        }
        public void Update_KhachHang(NhanVien_DTO ob)
        {
            nhanvien.Update_NhanVien(ob);
        }
        public void Delete_khachhang(string ID)
        {
            nhanvien.Detele_NhanVien(ID);
        }
    }
}
