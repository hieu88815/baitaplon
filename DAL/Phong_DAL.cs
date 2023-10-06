using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class Phong_DAL : Connect
    {
        public Phong_DAL() { }
        public DataTable Load_Phong_DAL()
        {
            return Load_DL("select * from phong");
        }
        public void Insert_Phong(Phong_DTO ob)
        {
            string sql = "insert into phong values ('" + ob.maphong + "', '" + ob.tenphong + "', N'" + ob.loaiphong + "')";
            Excecute(sql);
        }
        public void Update_Phong(Phong_DTO ob)
        {
            string sql = "update phong set tenphong = '" + ob.tenphong + "', loaiphong = N'" + ob.loaiphong + "' where maphong = '" + ob.maphong + "'";

            Excecute(sql);
        }
        public void detele_Phong(string maphong)
        {
            string sql = "delete from phong where maphong = '" + maphong + "'";
            Excecute(sql);
        }
    }
}
