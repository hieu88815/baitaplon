using DTO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class NhanVien_DAL : Connect
    {
        public NhanVien_DAL() { }
        public DataTable Load_NhanVien_DAL()
        {
            return Load_DL("select * from nhanvien");
        }
        public void Insert_NhanVien(DTO.NhanVien_DTO ob)
        {
            string sql = "insert into nhanvien values('" + ob.ID + "', N'" + ob.hovaten + "', '" + ob.ngaysinh + "', '" + ob.sdt + "', N'" + ob.machucvu + "', " + ob.sogiocong + ")";
            Excecute(sql);
        }
        public void Update_NhanVien(DTO.NhanVien_DTO ob)
        {
            string sql = "update nhanvien set hovaten = N'" + ob.hovaten + "', ngaysinh = '" + ob.ngaysinh + "', sdt = '" + ob.sdt + "', machucvu = '" + ob.machucvu + "', sogiocong = " + ob.sogiocong + " where ID = '" + ob.ID + "'";
            Excecute(sql);
        }
        public void Detele_NhanVien(string ID)
        {
            string sql = "delete from nhanvien where ID = '" + ID + "'";
            Excecute(sql);
        }
    }
}
