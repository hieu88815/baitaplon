using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class PhongDangThue_DAL : Connect
    {
        public PhongDangThue_DAL() { }
        public DataTable Load_PhongDangThue_DAL()
        {
            return Load_DL("select * from phongdangthue");
        }
        public void Insert_PhongDangThue(PhongDangThue_DTO ob)
        {
            string sql = "insert into phongdangthue values ('" + ob.ID + "', '" + ob.maphong + "', '" + ob.ngaynhanphong + "', '" + ob.ngaytraphong + "')";
            Excecute(sql);
        }
        public void UpDate_PhongDangThue(PhongDangThue_DTO ob)
        {
            string sql = "update phongdangthue set ngaynhanphong = '" + ob.ngaynhanphong + "', ngaytraphong = '" + ob.ngaytraphong + "' where ID = '" + ob.ID + "' and maphong = '" + ob.maphong + "'";
            Excecute(sql);
        }
        public void Delete_PhongDangThue(string ID, string maphong)
        {
            string sql = "delete from phongdangthue where ID = '" + ID + "' and maphong = '" + maphong + "'";
            Excecute(sql);
        }
    }
}
