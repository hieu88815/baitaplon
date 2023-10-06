using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class ChucVu_DTO
    {
        public string machucvu;
        public string tenchucvu;
        public double luongcung;
        public ChucVu_DTO() { }
        public ChucVu_DTO(string machucvu, string tenchucvu, double luongcung)
        {
            this.machucvu = machucvu;
            this.tenchucvu = tenchucvu;
            this.luongcung = luongcung;
        }

    }
}
