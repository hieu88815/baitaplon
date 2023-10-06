using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class KhachHang_DAL : Connect
    {
        public KhachHang_DAL() { }
        public DataTable Load_KhachHang_DAL()
        {
            return Load_DL("select * from khachhang");
        }
        public void Update_khachHang(KhachHang_DTO ob)
        {
            string sql = "update khachhang set hovaten = N'" + ob.hovaten + "', ngaysinh ='" + ob.ngaysinh + "', sdt = '" + ob.sdt + "' where ID = '" + ob.ID + "'";
            Excecute(sql);
        }
        public void Insert_KhachHang(KhachHang_DTO ob)
        {
            string sql = "insert into khachhang values('" + ob.ID + "', N'" + ob.hovaten + "', '" + ob.ngaysinh + "', '" + ob.sdt + "')";
            Excecute(sql);
        }
        public void Delete_KhachHang(string ID)
        {
            string sql = "delete from khachhang where ID = '" + ID + "'";
            Excecute(sql);
        }
    }
}
