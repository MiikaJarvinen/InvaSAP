using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvaSAP
{
    // Kevyt työtilausluokka, jota käytetään esimerkiksi avoimien töiden listauksessa.
    public class OpenWorkOrder
    {
        public string id { get; set; }
        public string kuvaus { get; set; }
        public string laite { get; set; }
        public string laiteKuvaus { get; set; }
    }


}
