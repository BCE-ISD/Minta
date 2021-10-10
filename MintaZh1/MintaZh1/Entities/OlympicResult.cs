using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MintaZh1.Entities
{
    class OlympicResult
    {
        public int Year { get; set; }
        public string Country { get; set; }
        public int[] Medals { get; set; }
        public int Position { get; set; }
    }
}
