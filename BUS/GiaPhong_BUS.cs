using DTO;
using DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BUS
{
    public class GiaPhong_BUS
    {
        GiaPhong_DAL giaphong = new GiaPhong_DAL();
        public DataTable Load_GiaPhong()
        {
            return giaphong.Load_GiaPhong_DAL();
        }
        public void Insert_Giaphong(GiaPhong_DTO ob)
        {
            giaphong.Insert_GiaPhong(ob);
        }
        public void Update_Giaphong(GiaPhong_DTO ob)
        {
            giaphong.Update_GiaPhong(ob);
        }
        public void Delete_GiaPhong(string loaiphong)
        {
            giaphong.Delete_GiaPhong(loaiphong);
        }

    }
}
