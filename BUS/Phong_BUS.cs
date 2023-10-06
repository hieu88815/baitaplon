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
    public class Phong_BUS
    {
        Phong_DAL phongdt = new Phong_DAL();
        public DataTable Load_Phong()
        {
            return phongdt.Load_Phong_DAL();
        }
        public void Insert_Phong(Phong_DTO ob)
        {
            phongdt.Insert_Phong(ob);
        }
        public void Update_Phong(Phong_DTO ob)
        {
            phongdt.Update_Phong(ob);
        }
        public void Delete_Phong(string maphong)
        {
            phongdt.detele_Phong(maphong);
        }
    }
}
