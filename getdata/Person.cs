using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace getdata
{
    public class Person
    {
        private string name;
        private string dc;
        private string sdt;
        private string email;

        public string Name { get => name; set => name = value; }
        public string Dc { get => dc; set => dc = value; }
        public string Sdt { get => sdt; set => sdt = value; }
        public string Email { get => email; set => email = value; }
    }
}
