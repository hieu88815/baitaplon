using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class GiaPhong_DTO
    {
        public string loaiphong;
        public double gia;
        public GiaPhong_DTO() { }
        public GiaPhong_DTO(string loaiphong, double gia)
        {
            this.loaiphong = loaiphong;
            this.gia = gia;
        }

    }
}
