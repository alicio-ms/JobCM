using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace JobCM
{
    class Program
    {
        static void Main(string[] args)
        {
        }
    }

    class AzureData
    {
        
    }

    public class ConstrutivoData
    {
        public static string LogaConstrutivo(string strLogin, string strSenha)
        {
            // Cria a requisição
            RequestState hwState = new RequestState("http://jundiai.construtivo.com/ssf/s/portalLogin");

            // Corpo da Solicitação
            hwState.StrBody = String.Format("j_username={0}&j_password={1}&okBtn=OK", strLogin, strSenha);

            // Seta para o Tipo POST
            hwState.HwrRequest.Method = "POST";
            hwState.HwrRequest.Host = "jundiai.construtivo.com";
            hwState.HwrRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            hwState.HwrRequest.Headers.Add("Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3");
            hwState.HwrRequest.Headers.Add("Accept-Encoding", "gzip, deflate");
            hwState.HwrRequest.Referer = "http://jundiai.construtivo.com/ssf/a/do?p_name=ss_forum&p_action=1&action=__login";
            hwState.HwrRequest.ContentType = "application/x-www-form-urlencoded";
            hwState.HwrRequest.KeepAlive = true;

            // Adiciona um Cookie Container
            hwState.HwrRequest.CookieContainer = new CookieContainer();

            // Seta o cookie de saida            
            hwState.HwrRequest.CookieContainer.SetCookies(new Uri("http://jundiai.construtivo.com"), "");

            // Começa a requisição assincronamente
            hwState.HwrRequest.BeginGetRequestStream(Requisicoes.GetRequestStreamCallback, hwState);

            // Aguarda o temino da requisição
            hwState.AllDone.WaitOne();

            // Retorna o Valor da Sessão
            return hwState.HwrRequest.Headers[6];
        }
    }

    // Classe de Requisiçoes HTML
    public class RequestState
    {
        // Requisição
        public HttpWebRequest HwrRequest { get; set; }

        // Resposta
        public HttpWebResponse HwResponse { get; set; }

        // Resposta em formato de String
        public object OResposta { get; set; }

        // All Done de Témino da Resposta
        public ManualResetEvent AllDone { get; set; }

        // Corpo
        public string StrBody { get; set; }


        // Total Arquivos na consulta
        public int IntTotal { get; set; }

        //Erro na requisição
        public bool BErro { get; set; }

        // Cria um Request State
        public RequestState(string strUrl)
        {
            // Cria a Requisição
            HwrRequest = (HttpWebRequest)WebRequest.Create(strUrl);

            // Seta o AllDone para aguardar
            AllDone = new ManualResetEvent(false);

            //Inicializa o Erro para Falso
            BErro = false;

            // Iniclaiza o Stream
            OResposta = null;
        }
    }

    class HtmlProcess
    {
        
    }

    class AnaliseDiagramas
    {
        
    }

    class AnaliseListas
    {
        
    }

    class AnaliseProjetos
    {
        
    }

    public class Requisicoes
    {
        // Envia Solicitação de Login
        public static void GetRequestStreamCallback(IAsyncResult callbackResult)
        {
            // Cria a requisição
            RequestState hwState = (RequestState)callbackResult.AsyncState;

            // Cria o fluxo de postagem
            Stream postStream = hwState.HwrRequest.EndGetRequestStream(callbackResult);

            // Transforma em Bytes o corpo da requisição
            byte[] byteArray = Encoding.UTF8.GetBytes(hwState.StrBody);

            // Adiciona o corpo da requisição
            postStream.Write(byteArray, 0, byteArray.Length);

            // fecha o fluxo
            postStream.Close();

            // Chava a resposta de modo assincrono
            hwState.HwrRequest.BeginGetResponse(GetResponseStreamCallback, hwState);
        }

        // Retorna Resposta de Solicitação
        static void GetResponseStreamCallback(IAsyncResult callbackResult)
        {
            // Carrega a String Temporária de Resposta
            RequestState hwState = (RequestState)callbackResult.AsyncState;

            // Cria as requisições
            hwState.HwResponse = (HttpWebResponse)hwState.HwrRequest.EndGetResponse(callbackResult);

            // Usando a resposta Salva o Valor em uma String
            using (StreamReader httpWebStreamReader = new StreamReader(hwState.HwResponse.GetResponseStream()))
            {
                //Transforma o fluxo em String
                hwState.OResposta = httpWebStreamReader.ReadToEnd();

                // Seta o All Done
                hwState.AllDone.Set();
            }
        }

    }
}
