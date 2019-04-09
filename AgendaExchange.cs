using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CorporativoWebApi.Infra.ActiveDirectory
{
    public class AgendaExchange
    {
        private string usuario;
        private string senha;
        private string url;
        private ExchangeService service;


        /// <summary>
        /// Inicializa a conexão com o Exchange.
        /// </summary>
        /// <param name="usuario">Usuário do AD com a devida permissão no calendário que será exibido. Recomenda-se utilizar uma conta de serviço para não expor a senha de um usuário</param>
        /// <param name="senha">Senha do referido usuário</param>
        /// <param name="url">Url do EWS</param>
        public AgendaExchange(string usuario, string senha, string url)
        {
            this.usuario = usuario;
            this.senha = senha;
            this.url = url;
           

            EstabelecerConexao();
        }

        private void EstabelecerConexao()
        {
            service = new ExchangeService(ExchangeVersion.Exchange2010);
            service.Url = new Uri(this.url);
            service.Credentials = new NetworkCredential(this.usuario, this.senha, "Corregedoria.df");
        }

        protected IList<CompromissoDto> RecuperarCompromissos(DateTime inicio, DateTime fim, int qtdItens,string contaCompartilhada)
        {
            CalendarFolder calendar = CalendarFolder.Bind(service, new FolderId(WellKnownFolderName.Calendar, contaCompartilhada));

            CalendarView cView = new CalendarView(inicio, fim, qtdItens);
            cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Location);
            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

            IList<CompromissoDto> lst = new List<CompromissoDto>();
            foreach (var item in appointments)
            {
                var compromisso = new CompromissoDto()
                {
                    Assunto = item.Subject,
                    Inicio = item.Start,
                    Termino = item.End,
                    Local = item.Location
                };
                lst.Add(compromisso);
            }
            return lst;
        }
    }
}
