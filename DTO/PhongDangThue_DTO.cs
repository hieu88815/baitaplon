using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class PhongDangThue_DTO
    {
        public string ID;
        public string maphong;
        public string ngaynhanphong;
        public string ngaytraphong;
        public PhongDangThue_DTO() { }
        public PhongDangThue_DTO(string ID, string maphong, string ngaynhanphong, string ngaytraphong)
        {
            this.ID = ID;
            this.maphong = maphong;
            this.ngaynhanphong = ngaynhanphong;
            this.ngaytraphong = ngaytraphong;
        }
    }
}
