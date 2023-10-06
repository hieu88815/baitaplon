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
    public class PhongDangThue_BUS
    {
        PhongDangThue_DAL phong = new PhongDangThue_DAL();
        public DataTable Load_PhongDangThue()
        {
            return phong.Load_PhongDangThue_DAL();
        }
        public void Insert_phongdangthue(PhongDangThue_DTO ob)
        {
            phong.Insert_PhongDangThue(ob);
        }
        public void Update_phongdangthue(PhongDangThue_DTO ob)
        {
            phong.UpDate_PhongDangThue(ob);
        }
        public void Delete_phongdangthue(string ID, string maphong)
        {
            phong.Delete_PhongDangThue(ID, maphong);
        }
    }
}
