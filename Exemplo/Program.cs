using System;
using NPAP;

namespace Exemplo
{
    class Program
    {
        static void Main(string[] args)
        {
            var impressora = new NpapClient("192.168.1.3");
            impressora.EventAlertaJob += impressora_EventAlertaJob;
            if (impressora.SolicitarAlertasDeJob())
            {
                // Executar esta tarefa em um BackgroundWorker
                impressora.MonitorarAlertas();
            }
            impressora.CancelarAlertasDeJob();
        }

        static void impressora_EventAlertaJob(object sender, AlertaJob alerta)
        {
            Console.WriteLine("JobId: " + alerta.JobId);
            Console.WriteLine("Duração: " + alerta.Duração / 1000 + " segundos");
            Console.WriteLine("Páginas: " + alerta.Páginas);
            Console.WriteLine("Usuário: " + alerta.Usuário);
            Console.WriteLine("Serviço: " + alerta.Serviço);
            Console.WriteLine("Bandeja: " + alerta.Bandeja);
            Console.WriteLine();
        }
    }
}
