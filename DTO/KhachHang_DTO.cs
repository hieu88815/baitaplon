using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class KhachHang_DTO
    {
        public string ID;
        public string hovaten;
        public string ngaysinh;
        public string sdt;
        public KhachHang_DTO() { }
        public KhachHang_DTO(string ID, string hovaten, string ngaysinh, string sdt)
        {
            this.ID = ID;
            this.hovaten = hovaten;
            this.sdt = sdt;
            this.ngaysinh = ngaysinh;
        }

    }
}
