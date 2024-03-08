using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Validador.Clases
{
    internal class Publicacion
    {
        public string Id { get; set; }
        public decimal Price { get; set; }
        public string Base_price { get; set; }
        public string Status { get; set; }
        public string Catalog_listing { get; set; }
        public List<Variacion> Variations { get; set; }       
    }
}
