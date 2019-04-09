using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorporativoWebApi.Infra.ActiveDirectory
{
    public class ConsultarAgendaDeEventos:AgendaExchange
    {

        public ConsultarAgendaDoControlador(string usuario,string senha,string url):base(usuario,senha,url)
        {

        }

        public object Executar(string contaCompartilhada,DateTime? dataEspecifica)
        {
            IList<CompromissoDto> compromissosAlternativos = new List<CompromissoDto>();
            IList<CompromissoDto> compromissosDoDia = new List<CompromissoDto>();

            string compromissosDoDiaReferencia = string.Empty;
            string compromissosAlternativosReferencia = string.Empty;

            if (dataEspecifica.HasValue)
            {
                compromissosAlternativosReferencia = dataEspecifica.Value.ToShortDateString();
                compromissosAlternativos = RecuperarCompromissos(dataEspecifica.Value, dataEspecifica.Value.AddHours(23).AddMinutes(59), 50, contaCompartilhada);
            }
            else
            {
                DateTime dataHoje = Convert.ToDateTime(DateTime.Now.ToShortDateString());
                DateTime proximoDia = DiaUtil(dataHoje.AddDays(1));
                compromissosDoDiaReferencia = DateTime.Now.ToShortDateString();
                compromissosAlternativosReferencia = proximoDia.ToShortDateString();

                compromissosAlternativos = RecuperarCompromissos(proximoDia, proximoDia.AddHours(23).AddMinutes(59), 50, contaCompartilhada);
                compromissosDoDia = RecuperarCompromissos(dataHoje, dataHoje.AddHours(23).AddMinutes(59), 50, contaCompartilhada);
            }

            return new
            {
                CompromissosDoDia =new  {Referencia=compromissosDoDiaReferencia,Itens= compromissosDoDia } ,
                CompromissosAlternativos =new {Referencia=compromissosAlternativosReferencia,Itens=compromissosAlternativos}
            };
        }

        private DateTime DiaUtil(DateTime dt)
        {
            while (true)
            {
                if (dt.DayOfWeek == DayOfWeek.Saturday)
                {
                    dt = dt.AddDays(2);
                    return DiaUtil(dt);
                }
                else if (dt.DayOfWeek == DayOfWeek.Sunday)
                {
                    dt = dt.AddDays(1);
                    return DiaUtil(dt);
                }
                else return dt;
            }
        }
    }
}
