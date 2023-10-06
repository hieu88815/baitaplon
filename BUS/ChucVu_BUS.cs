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
    public class ChucVu_BUS
    {
        ChucVu_DAL chucvu = new ChucVu_DAL();
        public DataTable Load_Chucvu()
        {
            return chucvu.Load_ChucVu_DAL();
        }

        public void Insert_Chucvu(ChucVu_DTO ob)
        {
            chucvu.Insert_chucvu(ob);
        }
        public void Update_Chucvu(ChucVu_DTO ob)
        {
            chucvu.Update_chucvu(ob);
        }
        public void Delete_Chucvu(string machucvu)
        {
            chucvu.Delete_chucvu(machucvu);
        }
    }
}
