using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class Phong_DTO
    {
        public string maphong;
        public string tenphong;
        public string loaiphong;
        public Phong_DTO() { }
        public Phong_DTO(string maphong, string tenphong, string loaiphong)
        {
            this.loaiphong = loaiphong;
            this.maphong = maphong;
            this.tenphong = tenphong;
        }
    }
}
