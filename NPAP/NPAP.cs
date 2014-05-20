using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;

namespace NPAP
{
    public class AlertaJob
    {
        public string JobId;
        public int Duração;
        public int Páginas;
        public string Usuário;
        public string Serviço;
        public string Bandeja;
    }

    public class NpapClient
    {
        #region Campos privados

        private const int RemotePort = 9300;
        private const int Timeout = 5000;

        private static readonly byte[] Comando = { 0x01, 0x00 };
        private static readonly byte[] Resposta = { 0x01, 0x01 };
        private static readonly byte[] Alerta = { 0x01, 0x02 };
        private static readonly byte[] Ack = { 0x01, 0x03 };

        private readonly UdpClient _clienteUdp;
        private IPEndPoint _impressora;
        private int _idPacote;
        private byte[] _ackNumber = { 0x00, 0x31 };

        #endregion

        #region Construtores

        /// <summary>
        /// Construtor do cliente NPAP
        /// </summary>
        /// <param name="ipImpressora">Endereço IP da impressora a ser monitorada</param>
        public NpapClient(string ipImpressora)
        {
            _impressora = new IPEndPoint(IPAddress.Parse(ipImpressora), RemotePort);
            _clienteUdp = new UdpClient();
            _idPacote = 1;
        }

        /// <summary>
        /// Solicita cancelamento de alertas e encerra conexões
        /// </summary>
        ~NpapClient()
        {
            CancelarAlertasDeJob();
            _clienteUdp.Close();
        }

        #endregion

        #region Métodos públicos

        /// <summary>
        /// Envia solicitação à impressora para que ela envie alertas de jobs realizados
        /// </summary>
        /// <returns>
        /// Retorna true caso comando tenha sido executado com sucesso na impressora
        /// </returns>
        public bool SolicitarAlertasDeJob()
        {
            var resposta = EnviarComando(new byte[] { 0xA5, 0x00, 0x04, 0x50, 0xE0, 0x73, 0x04 });
            if (resposta.Length > 10 && resposta[4] == 0xE0 && resposta[5] == 0x73 && resposta[6] == 0x04)
            {
                EnviarComando(new byte[] { 0xA5, 0x00, 0x05, 0x50, 0x03, 0x05, 0x02, 0x00 });
                var flag = (byte)(resposta[9] & 0x11);
                if (flag != 0)
                {
                    EnviarComando(new byte[] { 0xA5, 0x00, 0x0D, 0x50, 0xE0, 0x73, 0x01, 0x08, 
                                               0x00, flag, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 });
                }
                if ((flag & 0x10) == 0)
                {
                    EnviarComando(new byte[] { 0xA5, 0x00, 0x05, 0x50, 0xE0, 0xE1, 0x03, 0x01 });
                }
                if ((flag & 0x01) == 0)
                {
                    EnviarComando(new byte[] { 0xA5, 0x00, 0x04, 0x50, 0xE0, 0x03, 0x09 });
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// Envia solicitação à impressora para que ela não envie alertas de jobs realizados.
        /// </summary>
        /// <returns>
        /// Retorna true caso comando tenha sido executado com sucesso na impressora
        /// </returns>
        public bool CancelarAlertasDeJob()
        {
            byte[] solicitação = {
                0xA5, 0x00, 0x0D, 0x40, 0xE0, 0x73, 0x01, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
            };
            return EnviarComando(solicitação) != null;
        }

        /// <summary>
        /// Coleta pacotes enviados pela impressora e chama o método GerenciarAlerta caso seja um
        /// pacote de alerta. Esse método lança excessão caso algum erro ocorra!
        /// </summary>
        public void MonitorarAlertas()
        {
            while (true)
            {
                var pacote = _clienteUdp.Receive(ref _impressora);
                //try
                //{
                if (pacote[1] == Alerta[1])
                {
                    var sequence = pacote.Skip(4).Take(2).ToArray();
                    var ack = Ack.Concat(_ackNumber).Concat(sequence).ToArray();
                    _clienteUdp.Send(ack, ack.Length, _impressora);
                    var dados = pacote.Skip(12).Take(pacote.Length - 12).ToArray();
                    if ((dados[4] == 0xE0 || dados[4] == 0xF0) && dados[5] == 0x73 && (dados[6] == 0x01 || dados[6] == 0x03))
                    {
                        GerenciarAlerta(dados);
                    }
                }
                //}
                //catch (Exception)
                //{
                //    // IMPORTANTE - Logar excessão
                //}
            }
        }

        #endregion

        #region Métodos privados

        /// <summary>
        /// Envia comandos para a impressora
        /// </summary>
        /// <param name="comando">Byte array contendo comando</param>
        /// <returns>
        /// Retorna resposta à solicitação
        /// </returns>
        private byte[] EnviarComando(byte[] comando)
        {
            var idPacote = BitConverter.GetBytes(IPAddress.HostToNetworkOrder(_idPacote++));
            var mensagem = Comando.Concat(_ackNumber).Concat(idPacote).Concat(comando).ToArray();
            _clienteUdp.Send(mensagem, mensagem.Length, _impressora);

            try
            {
                _clienteUdp.Client.ReceiveTimeout = Timeout;
                var resposta = _clienteUdp.Receive(ref _impressora);
                _clienteUdp.Client.ReceiveTimeout = 0;

                var idResposta = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(resposta, 8));
                if (resposta[1] == Resposta[1] && idResposta == _idPacote - 1)
                {
                    _ackNumber = resposta.Skip(2).Take(2).ToArray();
                    var sequence = resposta.Skip(4).Take(2).ToArray();
                    var ack = Ack.Concat(_ackNumber).Concat(sequence).ToArray();
                    _clienteUdp.Send(ack, ack.Length, _impressora);
                    return resposta.Skip(12).Take(resposta.Length - 12).ToArray();
                }
            }
            catch (Exception)
            {
                return null;
            }
            return null;
        }

        /// <summary>
        /// Tarefa responsável por gerenciar alertas recebidos. Sua função é interpretar o pacote
        /// e disparar um evento.
        /// </summary>
        /// <param name="pacote">Byte array com pacote NPAP</param>
        /// <returns>Sempre retorna true</returns>
        public void GerenciarAlerta(byte[] pacote)
        {
            AlertaJob alerta = Interpretar(Analisar(pacote));
            if (alerta != null)
            {
                OnEventAlertaJob(alerta);
            }
        }

        /// <summary>
        /// Converte os dados do pacote para um formato amigável
        /// </summary>
        /// <param name="pacote">Byte array com pacote NPAP</param>
        /// <returns>Dicionário contendo dados do pacote</returns>
        private Dictionary<string, string> Analisar(byte[] pacote)
        {
            var dados = new Dictionary<string, string>();
            int j = pacote[7];
            int posição = 8 + j;
            int k = pacote[posição++];

            for (int i = 0; i < k; i++)
            {
                posição += 10;
                var array = pacote.Skip(posição).Take(8).ToArray();
                if ((array[1] & 0x01) == 1)
                {
                    int tipo = LerNumero(pacote, ref posição, 2);
                    if (tipo == 1)
                    {
                        posição += 34;
                        AnáliseRecursiva(pacote, ref posição, "", dados);
                    }
                }
            }
            return dados;
        }

        /// <summary>
        /// Analisa recursivamente os nós do pacote NPAP
        /// </summary>
        /// <param name="pacote">Byte array com pacote NPAP</param>
        /// <param name="posição">Ponteiro contendo posição atual no pacote</param>
        /// <param name="parâmetro">Parametro pai do nó</param>
        /// <param name="dados">Dicionário contendo dados coletados</param>
        private void AnáliseRecursiva(byte[] pacote, ref int posição, string parâmetro, Dictionary<string, string> dados)
        {
            if (posição >= pacote.Length)
            {
                return;
            }
            var pid = LerNumero(pacote, ref posição, 1);
            var pdt = LerNumero(pacote, ref posição, 1);
            parâmetro = parâmetro + "." + pid.ToString();
            var valor = "";
            int tamanho;

            switch (pdt)
            {
                case 0:
                    var pne = LerNumero(pacote, ref posição, 2);
                    tamanho = LerNumero(pacote, ref posição, 2);
                    dados.Add(parâmetro, pne.ToString());

                    for (int i = 0; i < pne; i++)
                    {
                        AnáliseRecursiva(pacote, ref posição, parâmetro, dados);
                    }
                    return;

                case 1:
                    tamanho = LerNumero(pacote, ref posição, 2);
                    valor = LerNumero(pacote, ref posição, tamanho).ToString();
                    break;

                case 2:
                case 3:
                    tamanho = LerNumero(pacote, ref posição, 2);
                    var texto = new byte[tamanho];
                    try
                    {
                        texto = pacote.Skip(posição).Take(tamanho).ToArray();
                        posição += tamanho;
                    }
                    catch (Exception) { }
                    valor = System.Text.Encoding.UTF8.GetString(texto);
                    break;

                default:
                    valor = "Tipo de dado desconhecido";
                    break;
            }
            dados.Add(parâmetro, valor);
        }

        /// <summary>
        /// Transforma os bytes de uma determinada posição do pacote em um inteiro com sinal
        /// </summary>
        /// <param name="pacote">Byte array com pacote NPAP</param>
        /// <param name="posicao">Ponteiro contendo posição no pacote a ser convertida</param>
        /// <param name="tamanho">Tamanho do byte a ser convertido</param>
        /// <returns>Valor lido</returns>
        private int LerNumero(byte[] pacote, ref int posicao, int tamanho)
        {
            var valor = 0;

            if (posicao + tamanho <= pacote.Length)
            {
                switch (tamanho)
                {
                    case 1:
                        valor = (int)(sbyte)pacote[posicao++];
                        break;

                    case 2:
                        valor = IPAddress.NetworkToHostOrder(BitConverter.ToInt16(pacote, posicao));
                        posicao += 2;
                        break;

                    case 4:
                        valor = IPAddress.NetworkToHostOrder(BitConverter.ToInt32(pacote, posicao));
                        posicao += tamanho;
                        break;
                }
            }
            return valor;
        }

        /// <summary>
        /// Transforma dados do dicionário em estrutura contendo dados de interesse
        /// </summary>
        /// <param name="dados">Dicionário com todos os dados do pacote</param>
        /// <returns>Estrutura com informações de um alerta de job</returns>
        private AlertaJob Interpretar(Dictionary<string, string> dados)
        {
            if (dados.ContainsKey(".1.1.1.2"))
            {
                var alerta = new AlertaJob();
                alerta.JobId = dados[".1.1.1.2"];
                alerta.Duração = Convert.ToInt32(dados[".1.1.1.7"]);

                alerta.Páginas = 0;
                if (dados.ContainsKey(".1.1.2.1"))
                {
                    for (int i = 1; i <= 5; i++)
                    {
                        var key = ".1.1.2.1." + i.ToString() + ".2.2";
                        if (dados.ContainsKey(key))
                        {
                            alerta.Páginas += Convert.ToInt32(dados[key]);
                        }
                    }
                }
                if (dados.ContainsKey(".1.1.2.2"))
                {
                    for (int i = 1; i <= 2; i++)
                    {
                        var key = ".1.1.2.2." + i.ToString() + ".2.2";
                        if (dados.ContainsKey(key))
                        {
                            alerta.Páginas += Convert.ToInt32(dados[key]);
                        }
                    }
                }
                if (dados.ContainsKey(".1.1.2.3"))
                {
                    for (int i = 1; i <= 3; i++)
                    {
                        var key = ".1.1.2.3." + i.ToString() + ".2.2";
                        if (dados.ContainsKey(key))
                        {
                            alerta.Páginas += Convert.ToInt32(dados[key]);
                        }
                    }
                }
                if ((alerta.Páginas == 0) && (Convert.ToInt32(alerta.Duração) < 1000))
                {
                    return null;
                }
                if (dados.ContainsKey(".1.1.2.1.1.5"))
                {
                    alerta.Bandeja = dados[".1.1.2.1.1.5"];
                }
                if (dados.ContainsKey(".1.1.1.4"))
                {
                    var mensagem = dados[".1.1.1.4"];
                    var match = Regex.Match(mensagem, @"(?<campo>(UR|HT)):\s?(?<valor>[\w\s]+)");
                    while (match.Success)
                    {
                        if (match.Groups["campo"].ToString() == "UR")
                        {
                            alerta.Usuário = match.Groups["valor"].ToString();
                        }
                        else if (match.Groups["campo"].ToString() == "HT")
                        {
                            alerta.Serviço = match.Groups["valor"].ToString();
                        }
                        match = match.NextMatch();
                    }
                    return alerta;
                }
            }
            return null;
        }

        #endregion

        #region Eventos

        /// <summary>
        /// Event Handler para alertas recebidos
        /// </summary>
        public event AlertReceivedEventHandler EventAlertaJob;

        /// <summary>
        /// Método que dispara o event handler EventAlertaJob
        /// </summary>
        /// <param name="alerta">Alerta recebido</param>
        private void OnEventAlertaJob(AlertaJob alerta)
        {
            if (EventAlertaJob != null)
            {
                EventAlertaJob(this, alerta);
            }
        }

        #endregion
    }

    /// <summary>
    /// Protótipo de um evento de alerta de job
    /// </summary>
    /// <param name="sender">Objeto que disparou o evento</param>
    /// <param name="alerta">Alerta recebido</param>
    public delegate void AlertReceivedEventHandler(object sender, AlertaJob alerta);
}