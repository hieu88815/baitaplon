using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class NhanVien_DTO
    {
        public string ID;
        public string hovaten;
        public string ngaysinh;
        public string sdt;
        public string machucvu;
        public int sogiocong;
        public NhanVien_DTO() { }
        public NhanVien_DTO(string id, string hovaten, string ngaysinh, string std, string machucvu, int sogiocong)
        {
            this.ID = id;
            this.hovaten = hovaten;
            this.ngaysinh = ngaysinh;
            this.sdt = std;
            this.machucvu = machucvu;
            this.sogiocong = sogiocong;
        }
    }
}
