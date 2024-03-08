using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Validador.Clases
{
    internal class Variacion
    {
        public string Id { get; set; }
        public List<Atributo> Attributes { get; set; }
        public void AddAttibute(Atributo attribute)
        {
            if (Attributes == null)
            {
                Attributes = new List<Atributo>();
            }
            Attributes.Add(attribute);
        }
        public string SKU { get; set; }
    }
}
