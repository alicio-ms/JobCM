using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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
        public string StrSession { get; }

        public bool Logado { get; }
        public ConstrutivoData(string strLogin, string strSenha)
        {
            StrSession = LogaConstrutivo( strLogin, strSenha);

            Logado = true;
        }

        private string LogaConstrutivo(string strLogin, string strSenha)
        {
            // Cria a requisição
            RequestState hwState = new RequestState("http://jundiai.construtivo.com/ssf/s/portalLogin")
            {
                StrBody = String.Format("j_username={0}&j_password={1}&okBtn=OK", strLogin, strSenha)
            };

            // Carrega os cabeçalhos
            hwState.HwrRequest.Referer = "http://jundiai.construtivo.com/ssf/a/do?p_name=ss_forum&p_action=1&action=__login";
            hwState.HwrRequest.Headers.Add("Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3");
            hwState.HwrRequest.Headers.Add("Accept-Encoding", "gzip, deflate");

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

        // Carrega todos os links
        public DataTable GetAllLinksNames(String hyperlink)
        {
            // Cria os indices
            Int32 erros = 0;

            // Cria a requisição
            RequestState hwState = new RequestState(hyperlink);

            // Carrega os cabeçalhos
            hwState.HwrRequest.Headers.Add("Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3");
            hwState.HwrRequest.Headers.Add("Accept-Encoding", "gzip, deflate");

            // Adiciona um Cookie Container
            hwState.HwrRequest.CookieContainer = new CookieContainer();

            // Seta o cookie de saida            
            hwState.HwrRequest.CookieContainer.SetCookies(new Uri("http://jundiai.construtivo.com"), StrSession);

            do
            {
                // Começa a requisição assincronamente
                hwState.HwrRequest.BeginGetRequestStream(new AsyncCallback(Requisicoes.GetRequestStreamCallback), hwState);

                // Aguarda o temino da requisição
                hwState.AllDone.WaitOne();

                // Se houve erro incrementa o contador
                if (hwState.BErro) erros++;

            } while ((hwState.BErro) && (erros < 5));
            

            // Retorna a resposta
            return HtmlProcess.LinkNamesFromHtml(hwState.OResposta as string);
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
            HwrRequest = (HttpWebRequest) WebRequest.Create(strUrl);

            HwrRequest.Method = "POST";
            HwrRequest.Host = "jundiai.construtivo.com";
            HwrRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            HwrRequest.ContentType = "application/x-www-form-urlencoded";
            HwrRequest.KeepAlive = true;

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
        public static DataTable LinkNamesFromHtml(string strResposta)
        {
            // Carrega as Listas
            List<String> links = GetLinksFromHtml(strResposta).ToList<String>();
            List<String> titulos = GetNamesFromHtml(strResposta).ToList<String>();

            // Cria a tabela
            DataTable retorno = new DataTable();

            // Adiciona as colunas
            retorno.Columns.Add("Titulo");
            retorno.Columns.Add("Links");

            int indice;
            int total = links.Count;

            // Ajusta os links
            for (indice = 0; indice < total; indice++)
            {
                object[] obj = { titulos[indice], links[indice] };
                retorno.Rows.Add(obj);
            }

            return retorno;
        }

        // Carrega todos os links de uma página
        private static String[] GetLinksFromHtml(String strHtml)
        {
            HtmlAgilityPack.HtmlDocument htmLdoc = new HtmlAgilityPack.HtmlDocument();
            htmLdoc.LoadHtml(strHtml);

            var temp = from no in htmLdoc.DocumentNode.Descendants("a") select no.GetAttributeValue("href", "");

            return temp.ToArray();
        }


        // Carrega todos os links de uma página
        private static String[] GetNamesFromHtml(String strHtml)
        {
            HtmlAgilityPack.HtmlDocument htmLdoc = new HtmlAgilityPack.HtmlDocument();
            htmLdoc.LoadHtml(strHtml);

            var temp = from no in htmLdoc.DocumentNode.Descendants("a") select no.InnerHtml.Replace("\n", "");

            return temp.ToArray();
        }
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

            if (hwState.StrBody != null)
            {
                // Cria o fluxo de postagem
                Stream postStream = hwState.HwrRequest.EndGetRequestStream(callbackResult);

                // Transforma em Bytes o corpo da requisição
                byte[] byteArray = Encoding.UTF8.GetBytes(hwState.StrBody);

                // Adiciona o corpo da requisição
                postStream.Write(byteArray, 0, byteArray.Length);

                // fecha o fluxo
                postStream.Close();
            }

            // Chama a resposta de modo assincrono
            try
            {
                hwState.HwrRequest.BeginGetResponse(new AsyncCallback(GetResponseStreamCallback), hwState);
            }
            catch
            {
                // Seta o Erro
                hwState.BErro = true;

                // Seta o AllDone
                hwState.AllDone.Set();
            }
        }

        
        // Retorna Resposta de Solicitação
        static void GetResponseStreamCallback(IAsyncResult callbackResult)
        {
            // Carrega a String Temporária de Resposta
            RequestState hwState = (RequestState)callbackResult.AsyncState;

            try
            {
                // Cria as requisições
                hwState.HwResponse = (HttpWebResponse)hwState.HwrRequest.EndGetResponse(callbackResult);

                // Usando a resposta Salva o Valor em uma String
                using (StreamReader httpWebStreamReader = new StreamReader(hwState.HwResponse.GetResponseStream()))
                {
                    //Transforma o fluxo em String
                    hwState.OResposta = httpWebStreamReader.ReadToEnd();

                    // Requisicao completada OK
                    hwState.BErro = false;

                    // Seta o All Done
                    hwState.AllDone.Set();
                }
            }
            catch
            {
                // Erro da Requisicao
                hwState.BErro = true;

                // Seta o All Done
                hwState.AllDone.Set();
            }
        }

    }
}
