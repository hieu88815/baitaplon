using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class GiaPhong_DAL : Connect
    {
        public GiaPhong_DAL() { }
        public DataTable Load_GiaPhong_DAL()
        {
            return Load_DL("select * from giaphong");
        }
        public void Update_GiaPhong(GiaPhong_DTO ob)
        {
            string sql = "update giaphong set gia = " + ob.gia + " where loaiphong = N'" + ob.loaiphong + "'";
            Excecute(sql);
        }
        public void Insert_GiaPhong(GiaPhong_DTO ob)
        {
            string sql = "insert into giaphong values(N'" + ob.loaiphong + "', " + ob.gia + ")";
            Excecute(sql);
        }
        public void Delete_GiaPhong(string loaiphong)
        {
            string sql = "delete from giaphong where loaiphong = N'" + loaiphong + "'";
            Excecute(sql);
        }
    }
}
