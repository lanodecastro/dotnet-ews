using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorporativoWebApi.Infra.ActiveDirectory
{
    public class CompromissoDto
    {
        public DateTime Inicio { get; set; }
        public DateTime Termino { get; set; }
        public string Assunto { get; set; }
        public string Local { get; set; }


    }
}
