using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Excel;
using HtmlAgilityPack;
using Syncfusion.DocIO.DLS;
using static System.Convert;
using static System.Data.SqlDbType;
using static System.IO.Path;
using static System.String;
using DateTime = System.DateTime;

namespace JobCM
{
    class Program
    {
        static void Main(string[] args)
        {
            //Atualiza os diagramas
            //AnaliseDiagramas.AtualizaDiagramas();

            //Atuliza as Listas
            //AnaliseListas.AtualizaListas();

            // Atualiza os dados de projetos
            AnaliseProjetos.AtualizaProjetos();

            // Atualiza os Arquivos direto da base de dados
            Arquivos.AtualizaArquivo();
            
            // Lista os dados de material
            AnaliseMaterial.ListaMaterial();
        }
    }

    // Classe do Azure
    public class AzureData
    {
        private readonly string _strConnection;

        public AzureData(string connection)
        {
            _strConnection = connection;
        }

        // Copia uma tabela do Access através do comando e do arquivo
        public DataTable GetTabelaAzure(string strCommand)
        {
            // Carrega a Concexào
            SqlConnection conConnection = new SqlConnection(_strConnection);

            // Abre c conexão
            conConnection.Open();

            // Cria o Comando
            SqlCommand cmdCommand = new SqlCommand(strCommand, conConnection);

            // Cria a tabela
            DataTable dtTemp = new DataTable();

            //Preenche a tabela
            dtTemp.Load(cmdCommand.ExecuteReader());

            // Fecha a conexão
            conConnection.Close();

            // retorna tabela
            return dtTemp;
        }

        // Insere um diagrama na base de dados
        public void InsereArquivo(MemoryStream msTemp, string nome, string strTabela)
        {
            UpdateArquivo(new Dictionary<string, object> { {"@filename", nome}, {"@storage", msTemp.ToArray()} },
                $"INSERT INTO [{strTabela}] (Nome, Arquivo) VALUES (@filename, @storage);");
        }

        private void UpdateArquivo(Dictionary<string, object> parametros, string strcomando)
        {
            // Carrega a Concexào
            SqlConnection conConnection = new SqlConnection(_strConnection);

            // Abre c conexão
            conConnection.Open();

            // Adiciona a String de Inserção
            SqlCommand cmdInsertCommand = new SqlCommand(strcomando, conConnection) {CommandTimeout = 600};

            // Insere os Parametros
            foreach (KeyValuePair<string, object> par in parametros) cmdInsertCommand.Parameters.AddWithValue(par.Key, par.Value);

            // Insere os Dados
            cmdInsertCommand.ExecuteNonQuery();

            // Fecha a conexão
            conConnection.Close();
        }

        // Copia os dados de todos os arquivos
        public DataTable CarregaDiagramas(string tabela)
        {
            // Cria a Tabela
            DataTable dtTabela = new DataTable();
            
            // Adiciona as colunas
            dtTabela.AddColumns(new[] {"Diagrama", "Revisão", "Código", "Origem", "Destino", "Formação", "Folha"} );

            // Tabela Temporária
            DataTable dtTemp = GetTabelaAzure($"SELECT Nome FROM {tabela} ORDER BY Nome;");

            // Cria uma lista para armazenar as tarefas
            List<Task<DataTable>> dtTabelaTask2 = new List<Task<DataTable>>();
            
            // Carrega os dados
            foreach (DataRow dr in dtTemp.Rows)
            {
                while (dtTabelaTask2.Count(dtT => (dtT.Status != TaskStatus.RanToCompletion) && (dtT.Status != TaskStatus.Faulted)) > 20) Thread.Sleep(100);

                //MemoryStream ms = new MemoryStream();

                //while (ms.Length == 0) ms = GetFile(dr.Field<string>("Nome"), tabela);

                // Carrega os dados
                dtTabelaTask2.Add(Task.Factory.StartNew(() => AnaliseDiagramas.DiagramaFromExcel(GetFile(dr.Field<string>("Nome"), tabela), dr.Field<string>("Nome"))));
            }

            while (dtTabelaTask2.Count(dtT => (dtT.Status != TaskStatus.RanToCompletion) && (dtT.Status != TaskStatus.Faulted)) > 0) Thread.Sleep(1000);

            // Carrega todas as tabelas
            foreach (Task<DataTable> dt in dtTabelaTask2) dtTabela.Merge(dt.Result);

            // Retorna a lista de valores
            return dtTabela;
        }

        // Copia os dados de todos os arquivos
        public DataTable CarregaListas(string tabela)
        {
            // Cria a Tabela
            DataTable dtTabela = new DataTable();
            dtTabela.AddColumns(new[] { "Lista", "Revisao", "Codigo", "Comprimento", "Formacao", "Percurso" });
            dtTabela.Columns["Comprimento"].DataType = typeof(int);

            // Abre a conexão
            SqlConnection conn = new SqlConnection(_strConnection);
            conn.Open();

            // Carrega o Comando
            SqlCommand comm = new SqlCommand($"SELECT Nome, Arquivo FROM [{tabela}] WHERE Nome LIKE '%.doc%' ORDER BY Nome;", conn);

            // Executa a leitura
            using (SqlDataReader dataReader = comm.ExecuteReader())
            {
                // Verifica se não foram encontradas linhas
                if (dataReader.HasRows)
                {
                    while (dataReader.Read()) 
                        dtTabela.Merge(AnaliseListas.ListaFromWord(new MemoryStream((byte[])dataReader["Arquivo"]), dataReader["Nome"] as string));
                }
            }

            // Fecha a conexao
            conn.Close();

            // Retorna a lista de valores
            return dtTabela;
        }

        // Carrega dados das listas
        public DataTable CarregaLMs()
        {
            // Cria uma tabela
            DataTable dtTabela = new DataTable();
            dtTabela.AddColumns(new[] { "Lista", "Revisao", "Item", "Código Copem", "Quantidade" });

            // Abre a conexão
            SqlConnection conn = new SqlConnection(_strConnection);
            conn.Open();

            // Carrega o Comando
            SqlCommand comm =
                new SqlCommand("SELECT Nome, Arquivo FROM [Arquivo Projetos] WHERE Nome IN (SELECT DISTINCT Titulo FROM [Dados Projetos] WHERE Titulo LIKE '%FVH%-LM-%.doc%');",
                    conn) {CommandTimeout = 600};

            // Executa a leitura
            using (var dataReader = comm.ExecuteReader())
            {
                // Verifica se não foram encontradas linhas
                if (dataReader.HasRows)
                {
                    while (dataReader.Read())
                        dtTabela.Merge(AnaliseMaterial.ListaMaterialFromWord(new MemoryStream((byte[])dataReader["Arquivo"]), dataReader["Nome"].ToString()));
                }
            }

            // Fecha a conexao
            conn.Close();

            // Retorna a tabela
            return dtTabela;
        }

        // verificar depois
        private MemoryStream GetFile(string name, string table)
        {
            MemoryStream ms = new MemoryStream();

            try
            {
                // Abre a conexão
                using (SqlConnection conn = new SqlConnection(_strConnection))
                {
                    conn.Open();

                    // Carrega o Comando
                    using (SqlCommand comm = new SqlCommand($"SELECT Arquivo FROM {table} WHERE Nome = '{name}';", conn))
                    {
                        // Austa o timeout
                        comm.CommandTimeout = 6000;

                        // Executa a leitura
                        using (SqlDataReader dataReader = comm.ExecuteReader())
                        {
                            dataReader.Read();
                            ms = new MemoryStream((byte[]) dataReader["Arquivo"]);
                        }
                    }

                    conn.Close();

                    return ms;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro: " + name + " - " + ex.Message);
                return ms;
            }
        }

        // Função que insere dados dos diagramas na base de dados
        public void InsereTabelaDiagramas(string strTabela, DataTable dtTabela)
        {
            UpdateDados($"INSERT INTO [{strTabela}] VALUES (@Diagrama, @Revisao, @Codigo, @Origem, @Destino, @Formacao, @Folha);",
                new[] { "Diagrama", "Revisao", "Codigo", "Origem", "Destino", "Formacao", "Folha" }, new[] { 23, 2, 20, 50, 50, 20, 10 }, dtTabela);
        }

        // Funcao de insercao de dados
        private void UpdateDados(string comando, string[] parametros, int[] tamanhos, DataTable dtTabela)
        {
            // Carrega a Concexào
            SqlConnection conConnection = new SqlConnection(_strConnection);

            // Abre c conexão
            conConnection.Open();

            // Cria o adaptador
            SqlDataAdapter adapter = new SqlDataAdapter
            {
                InsertCommand = new SqlCommand(comando, conConnection)
            };

            // Prepara o Comando
            adapter.InsertCommand.Prepare();

            // Adiciona os parâmetros
            adapter.InsertCommand.AddicionaParametros(parametros, tamanhos);

            DataTable dtDataTable = new DataTable();

            // Adiciona as colunas
            dtDataTable.AddColumns(parametros);

            if (parametros.Contains("Data"))
            {
                adapter.InsertCommand.Parameters["@Data"].SqlDbType = SqlDbType.DateTime;
                dtDataTable.Columns["Data"].DataType = Type.GetType("System.DateTime");
            }

            foreach (DataRow linha in dtTabela.Rows) dtDataTable.Rows.Add(linha.ItemArray);

            // Atualiza os dados na tabela
            adapter.Update(dtDataTable);

            // Fecha a conexão
            conConnection.Close();
        }

        // Função que insere os dados das listas na base de dados
        public void InsereTabelaListas(string strTabela, DataTable dtTabela)
        {
            UpdateDados($"INSERT INTO [{strTabela}] VALUES (@Lista, @Revisao, @Codigo, @Comprimento, @Formacao, @Percurso);",
                new[] { "Lista", "Revisao", "Codigo", "Comprimento", "Formacao", "Percurso" }, new[] { 23, 2, 20, 50, 20, 800 }, dtTabela);
        }

        // Função que remove dados da tabela de diagramas
        public void RemoveTabelaDiagramas(string strTabela, DataTable dtTabela)
        {
            UpdateDados($"DELETE FROM [{strTabela}] WHERE (Diagrama = @Diagrama AND [Revisão] = @Revisao AND [Código] = @Codigo AND [Origem] " +
                "= @Origem AND [Destino] = @Destino AND [Formação] = @Formacao AND Folha = @Folha);",
                new[] { "Diagrama", "Revisao", "Codigo", "Origem", "Destino", "Formacao", "Folha" }, new[] { 23, 2, 20, 50, 50, 20, 10 }, dtTabela);
        }
        
        // Função que remove dados da tabela de listas
        public void RemoveTabelaListas(string strTabela, DataTable dtTabela)
        {
            UpdateDados($"DELETE FROM [{strTabela}] WHERE (Lista = @Lista AND [Revisão] = @Revisao AND [Código Cabo] = @Codigo AND [Comprimento]" +
                " = @Comprimento AND [Formação] = @Formacao AND [Percurso] = @Percurso);",
                new[] { "Lista", "Revisao", "Codigo", "Comprimento", "Formacao", "Percurso" }, new[] { 23, 2, 20, 50, 20, 800 }, dtTabela);
        }

        // Funcao que Remove Arquivos
        public void RemoveArquivos(DataTable dtTabela, string tabela)
        {
            // Verifica os arquivos a remover
            foreach (DataRow dr in dtTabela.Rows)
                    UpdateArquivo(new Dictionary<string, object> {{"@filename", dr.Field<string>("Nome")}},$"DELETE FROM [{tabela}] WHERE Nome=@filename;");
        }

        // Adiciona dados na base do Azure SQL
        public void AdicionaDadosProjetos(DataTable dtTabela, string strTabela)
        {
            UpdateDados(
                $"INSERT INTO [{strTabela}] VALUES (@Nrs, @Titulo, @Descricao, @Autor, @Tamanho, @Estado, @Data, @LinkBaixar, @LinkVisualizar, @objectId, @versionId);",
                new[] { "Nrs", "Titulo", "Descricao", "Autor", "Tamanho", "Estado", "Data", "LinkBaixar", "LinkVisualizar", "objectId", "versionId" },
                new[] { 255, 50, 255, 50, 50, 50, 255, 255, 255, 14, 12 }, 
                dtTabela.AsEnumerable().OrderBy(r => r.Field<DateTime>("Data")).AsEnumerable().GroupBy(x => x.Field<string>("Titulo")).Select(g => g.Last()).CopyToDataTable()
                );
        }

        // Adiciona dados na base do Azure SQL
        public void AdicionaDadosPdf(DataTable dtTabela, string strTabela)
        {
            DataTable dtDataTable = new DataTable();
            dtDataTable.AddColumns(new[] { "LinkVisualizar", "PDF", "PDFLink" });

            foreach (DataRow linha in dtTabela.Rows)
            {
                List<object> obj = new List<object>{linha.Field<string>("LinkVisualizar"),
                    GetFileName(linha.Field<string>("PDFLink")),
                    linha.Field<string>("Link PDF").Contains("///") ? "" : linha.Field<string>("PDFLink")
                };

                dtDataTable.Rows.Add(obj.ToArray());
            }

            UpdateDados($"INSERT INTO [{strTabela}] VALUES (@LinkVisualizar, @PDF, @PDFLink);", new[] { "LinkVisualizar", "PDF", "PDFLink" }, new[] { 255, 50, 255 }, dtDataTable);
        }
        
        // Adiciona dados na base do Azure SQL
        public void UpdateDadosObsoletos(DataTable dtTabela, string strTabela)
        {
            DataTable dtDataTable = new DataTable();
            dtDataTable.AddColumns(new[] { "Titulo", "Link Pai" });

            foreach (DataRow linha in dtTabela.Rows)
            {
                object[] obj = { linha.Field<string>("Titulo"), linha.Field<string>("Link Pai") };
                dtDataTable.Rows.Add(obj);
            }

            UpdateDados( $"UPDATE [{strTabela}] SET [Link Pai] = @LinkPai WHERE Titulo = @Titulo", new[] { "Titulo", "LinkPai" }, new[] { 50, 255}, dtDataTable);
        }

        // Executa um comando
        public int ExecutaComando(string comando)
        {
            // Carrega a Concexào
            SqlConnection conConnection = new SqlConnection(_strConnection);
            conConnection.Open();

            // Cria o Comando
            SqlCommand cmdCommand = new SqlCommand(comando, conConnection) {CommandTimeout = 600};

            int retorno = cmdCommand.ExecuteNonQuery();

            conConnection.Close();

            // Executa o Comando
            return retorno;
        }

        // Adiciona dados na base do Azure SQL
        public void AtualizaDadosProjetos(DataTable dtTabela, string strTabela)
        {

            UpdateDados($"UPDATE [{strTabela}] SET Nrs = @Nrs, Titulo = @Titulo, Descricao = @Descricao, Autor = @Autor, Tamanho = @Tamanho, Estado = @Estado, " +
                    "Data = @Data, [Link Baixar] = @LinkBaixar, [objectId] = @objectId, [versionId] = @versionId WHERE [Link Visualizar] = @LinkVisualizar;",
               new[] { "Nrs", "Titulo", "Descricao", "Autor", "Tamanho", "Estado", "Data", "LinkBaixar", "LinkVisualizar", "objectId", "versionId" },
               new[] { 255, 50, 255, 50, 50, 50, 255, 255, 255, 14, 12 },
               dtTabela.AsEnumerable().OrderBy(r => r.Field<DateTime>("Data")).AsEnumerable().GroupBy(x => x.Field<string>("Titulo")).Select(g => g.Last()).CopyToDataTable()
               );
        }
    }

    // Classe do Construtivo
    public class ConstrutivoData
    {
        public string StrSession { get; }

        public bool Logado { get; }
        public ConstrutivoData(string strLogin, string strSenha)
        {
            StrSession = LogaConstrutivo( strLogin, strSenha);

            Logado = true;
        }

        private static string LogaConstrutivo(string strLogin, string strSenha)
        {
            // Cria a requisição
            RequestState hwState = new RequestState("http://jundiai.construtivo.com/ssf/s/portalLogin")
            {
                StrBody = $"j_username={strLogin}&j_password={strSenha}&okBtn=OK",
                HwrRequest ={Referer = "http://jundiai.construtivo.com/ssf/a/do?p_name=ss_forum&p_action=1&action=__login"}
            };

           hwState.Requisita("");

            // Retorna o Valor da Sessão
            return hwState.HwrRequest.Headers[6];
        }

        // Carrega todos os links
        public DataTable GetAllLinksNames(string hyperlink)
        {
            // Cria a requisição
            RequestState hwState = new RequestState(hyperlink);

            hwState.Requisita(StrSession);

            // Retorna a resposta
            return HtmlProcess.LinkNamesFromHtml(Encoding.UTF8.GetString(hwState.OResposta.ToArray()));
        }

        // Carrega os dados de um arquivo
        public object[] GetDocumentData(string endereco)
        {
            // Variaveis
            RequestState hwState = new RequestState(endereco);
            hwState.Requisita(StrSession);

            // Retorna um array com os dados
            return HtmlProcess.GetDataFromHtml(Encoding.UTF8.GetString(hwState.OResposta.ToArray()), endereco);
        }

        // Função que baixa arquivo utilizando HttpWebRequest
        public MemoryStream BaixaArquivo(string strUrl)
        {
            // Cria a requisição
            RequestState hwState = new RequestState(strUrl)
            {
                HwrRequest = {AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate}
            };
            
            hwState.Requisita(StrSession);
            
            return hwState.OResposta;
        }

        // Carrega todos os links para visualização
        public async Task<List<string>> CopiaApenasLinks(string numero)
        {
            const string Base = "http://jundiai.construtivo.com/ssf/a/do?p_name=ss_forum&p_action=1&newTab=ssNewTabPlaceHolder&action=view_folder_listing&binderId=";

            // Cria lista de endereços
            List<string> enderecos = new List<string>();

            // Adiciona Adiciona os enderecos
            enderecos.AddRange(CopiaEnderecos(Base, numero));

            // Cria uma lista para armazenar as tarefas
            List<Task<List<string>>> dtTabelaTask = new List<Task<List<string>>>();

            // Para cada página
            foreach (string end in enderecos)
            {
                // Verifica se existem slots de download disponível (10 slots no total)      
                while (dtTabelaTask.Count(dtT => dtT.Status != TaskStatus.RanToCompletion) > 10)  Thread.Sleep(100);
                
                // Adiciona a tarefa
                dtTabelaTask.Add(Task<List<string>>.Factory.StartNew(() => GetAllLinks($"http://jundiai.construtivo.com/voith/buscaRapida.jsp?pasta={end}&nome=")));
            }

            // Aguarda o Término de todas as tabelas
            List<string>[] dtTabelas = await Task.WhenAll(dtTabelaTask);

            // Cria alista de strings
            List<string> aUrls = new List<string>();

            // Retorna a tabela            
            return dtTabelas.Aggregate(aUrls, (current, dt) => current.Union(dt).ToList());
        }

        // carrega os codigos de binder e entrada dos arquivos
        public DataTable GetCodes(DataTable dtInput)
        {
            // Cria a tabela
            DataTable dtTabela = new DataTable();

            // Cria as colunas
            dtTabela.AddColumns(new[] { "Link Visualizar", "Binder", "Entry" });

            //Carrega os dados para obtenção dos obsoletos
            foreach (DataRow dr in dtInput.Rows)
            {
                string temp = dr.Field<string>("Link Visualizar").Replace("http://jundiai.construtivo.com/ssf/a/c/p_name/ss_forum/p_action/1/action/view_folder_entry/binderId/", "");

                object[] obj = { dr.Field<string>("Link Visualizar"), temp.Split('/')[0], temp.Split('/')[2] };
                dtTabela.Rows.Add(obj);
            }

            return dtTabela;
        }

        // Carrega todos os links
        public List<string> GetAllLinks(string hyperlink)
        {
            // Cria a requisição
            RequestState hwState = new RequestState(hyperlink);
            
            hwState.Requisita(StrSession);

            // Carrega a Lista
            List<string> lista = HtmlProcess.GetLinksFromHtml(Encoding.UTF8.GetString(hwState.OResposta.ToArray())).ToList();

            // Retorna a resposta
            return lista.Select(str => str.Replace("/4083/", "/" + hyperlink.Replace("http://jundiai.construtivo.com/voith/buscaRapida.jsp?pasta=", "")
                        .Replace("&nome=", "") + "/")).ToList();
        }

        // Copia a lista de pastas
        private string[] CopiaEnderecos(string strUrl, string strNumber)
        {
            //Variaveis
            RequestState hwState = new RequestState(strUrl + strNumber);
            
            hwState.Requisita(StrSession);

            // Cria um htmldoc
            HtmlDocument htmLdoc = new HtmlDocument();

            // Carrega a resposta
            htmLdoc.LoadHtml(Encoding.UTF8.GetString(hwState.OResposta.ToArray()));

            // Carrega os Links de pastas
            HtmlNode[] enderecos = htmLdoc.GetElementbyId("sidebarWsTree_ss_forum_div" + strNumber).Descendants("a")
                .Where(d => d.GetAttributeValue("class", "").Contains("ss_tree_highlight_not")).ToArray();

            // Retorna o Array
            return enderecos.Select(no => no.GetAttributeValue("onclick", "").Replace("if (self.ss_treeShowId) {return ss_treeShowId('", "")
                        .Replace("', this,'view_folder_listing', '_ss_forum_');}", "")).Where(temp => temp.Length < 5).ToArray();
        }

        // Carrega os Valores de Links para PDFs
        public DataTable GetPdFs(DataTable dtTabela)
        {
            DataTable dtTemp = dtTabela.Clone();

            dtTemp.Columns.Add("Link PDF");

            // Cria uma lista para armazenar as tarefas
            List<Task<object[]>> dtTabelaTask2 = new List<Task<object[]>>();

            // Para as demais páginas
            foreach (DataRow dr in dtTabela.Rows)
            {
                // Verifica se existem slots de download disponível (10 slots no total)      
                while (dtTabelaTask2.Cast<Task>().Count(dtT => (dtT.Status != TaskStatus.RanToCompletion) && (dtT.Status != TaskStatus.Faulted)) > 10)
                    Thread.Sleep(500);

                // Copia os dados
                string str = dr.Field<string>("Link Visualizar").Replace("http://jundiai.construtivo.com/ssf/a/c/p_name/ss_forum/p_action/1/action/view_folder_entry"
                    + "/binderId/", "").Replace("entryId/", "");

                // Adiciona a tarefa
                dtTabelaTask2.Add(Task.Factory.StartNew(() => AddPdfLink(str.Split('/')[0], str.Split('/')[1], dr.Field<string>("Titulo"), dr.ItemArray)));
            }

            // Aguarda o Término de todas as tabelas
            Task.WaitAll(dtTabelaTask2.ToArray());

            foreach (Task<object[]> tsk in dtTabelaTask2.ToArray()) dtTemp.Rows.Add(tsk.Result);

            return dtTemp;
        }

        // Carrega o link dos PDFs
        private object[] AddPdfLink(string binder, string entry, string titulo, IReadOnlyList<object> linha)
        {
            return new[] { linha[0], linha[1], linha[2], linha[3], linha[4], linha[5], linha[6], linha[7], linha[8], GetPdfLink(binder, entry, titulo)};
        }

        // Carrega o link do pdf do construtivo
        private string GetPdfLink(string binder, string entry, string titulo)
        {
            // Cria a requisição
            RequestState hwState = new RequestState("http://jundiai.construtivo.com/voith/viewItemPdf.jsp")
            {
                StrBody = $"binder={binder}&entry={entry}&titulo={titulo}",
                HwrRequest = {Referer = $"http://jundiai.construtivo.com/ssf/a/c/p_name/ss_forum/p_action/1/action/view_folder_entry/binderId/{binder}/entryId/{entry}"}
            };
            
            hwState.Requisita(StrSession);

            // Remove /n e separa a string
            string[] temp = Encoding.UTF8.GetString(hwState.OResposta.ToArray()).Replace("\n", "").Split('#');

            // Retorna a resposta
            return $"http://jundiai.construtivo.com/ssf/s/readFile/folderEntry/{temp[2]}/{temp[3]}/0/last/{temp[4]}";

        }

        // Carrega o link do pdf do construtivo
        public string[] GetObsoleteLinks(string binder, string entry)
        {
            // Cria os indices
            int indice;

            // Cria a lista
            List<string> lista = new List<string>();

            // Cria a requisição
            RequestState hwState = new RequestState("http://jundiai.construtivo.com/voith/viewItemProjetoRevAnt.jsp")
            {
                StrBody = $"binder={binder}&entry={entry}&user=721",
                HwrRequest = {Referer = $"http://jundiai.construtivo.com/ssf/a/c/p_name/ss_forum/p_action/1/action/view_folder_entry/binderId/{binder}/entryId/{entry}"}
            };
            
            hwState.Requisita(StrSession);

            // Remove /n e os vazios
            string[] temp = Encoding.UTF8.GetString(hwState.OResposta.ToArray()).Replace("\n", "").Split('|').Where(x => x != "").ToArray();

            // Carrega o total
            int total = temp.Length / 3;

            // Carrega os links
            for (indice = 0; indice < total; indice++)
                lista.Add($"http://jundiai.construtivo.com/ssf/a/c/p_name/ss_forum/p_action/1/action/view_folder_entry/binderId/{temp[indice*3 + 2]}/entryId/{temp[indice*3 + 3]}");

            // Retorna a resposta
            return lista.ToArray();
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
        public MemoryStream OResposta { get; set; }

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

            // Cabeçalhos
            HwrRequest.Method = "POST";
            HwrRequest.Host = "jundiai.construtivo.com";
            HwrRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            HwrRequest.ContentType = "application/x-www-form-urlencoded";
            HwrRequest.KeepAlive = true;
            HwrRequest.Headers.Add("Accept-Language", "pt-BR,pt;q=0.8,en-US;q=0.5,en;q=0.3");
            HwrRequest.Headers.Add("Accept-Encoding", "gzip, deflate");
            
            // Adiciona um Cookie Container
            HwrRequest.CookieContainer = new CookieContainer();

            // Seta o AllDone para aguardar
            AllDone = new ManualResetEvent(false);

            //Inicializa o Erro para True
            BErro = true;

            // Iniclaiza o Stream
            OResposta = new MemoryStream();
        }

        public void Requisita(string session)
        {
            // Seta o cookie de saida            
            HwrRequest.CookieContainer.SetCookies(new Uri("http://jundiai.construtivo.com"), session);

            int erros = 0;
            do
            {
                // Começa a requisição assincronamente
                HwrRequest.BeginGetRequestStream(Requisicoes.GetRequestStreamCallback, this);

                // Aguarda o temino da requisição
                AllDone.WaitOne();

                // Se houve erro incrementa o contador
                if (BErro) erros++;

            } while (BErro && (erros < 5));
        }
    }

    // Classe que extracao de dados HTML
    public class HtmlProcess
    {
        public static DataTable LinkNamesFromHtml(string strResposta)
        {
            // Carrega as Listas
            List<string> links = GetLinksFromHtml(strResposta).ToList();
            List<string> titulos = GetNamesFromHtml(strResposta).ToList();

            // Cria a tabela
            DataTable retorno = new DataTable();
            retorno.AddColumns(new[] { "Titulo", "Links" });

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
        public static string[] GetLinksFromHtml(string strHtml)
        {
            HtmlDocument htmLdoc = new HtmlDocument();
            htmLdoc.LoadHtml(strHtml);

            return htmLdoc.DocumentNode.Descendants("a").Select(no => no.GetAttributeValue("href", "")).ToArray();
        }


        // Carrega todos os links de uma página
        public static string[] GetNamesFromHtml(string strHtml)
        {
            HtmlDocument htmLdoc = new HtmlDocument();
            htmLdoc.LoadHtml(strHtml);
            
            return htmLdoc.DocumentNode.Descendants("a").Select(no => no.InnerHtml.Replace("\n", "")).ToArray();
        }

        // Carrega os dados de um arquivo do html
        public static object[] GetDataFromHtml(string strHtml, string endereco)
        {
            HtmlDocument htmLdoc = new HtmlDocument();
            htmLdoc.LoadHtml(strHtml);
            
            try
            {
                // Cultura
                CultureInfo cultureBr = new CultureInfo("pt-BR");

                // Carrega os Links
                string temp = Regex.Replace(htmLdoc.DocumentNode.Descendants("span").First(d => d.Attributes.Contains("class") && d.Attributes["class"].Value.Contains("ss_entryTitle")).InnerText, @"[^\w\.-]", "");

                // Cria o objeto
                return new object[] {
                temp.Substring(0,temp.IndexOf(".", StringComparison.Ordinal) + 1),
                temp.Substring(temp.IndexOf(".", StringComparison.Ordinal) + 1),
                Regex.Replace(htmLdoc.DocumentNode.Descendants("div").Where(d => d.Attributes.Contains("class")&& d.Attributes["class"].Value.Contains("ss_entryContent")).ElementAt(2).InnerText.Replace("Descrição do Documento",""),@"[^\w\x20\.-]",""),
                Regex.Replace(htmLdoc.DocumentNode.Descendants("div").Where(d => d.Attributes.Contains("class")&& d.Attributes["class"].Value.Contains("ss_entryContent")).ElementAt(3).InnerText,@"[^\w\x20\.-]","").Replace("  ",""),
                Regex.Replace(htmLdoc.DocumentNode.Descendants("td").Where(d => d.Attributes.Contains("class")&& d.Attributes["class"].Value.Contains("ss_att_meta")).ElementAt(3).InnerText,@"[^\w\x20\-:]",""),
                htmLdoc.DocumentNode.Descendants("div").First(d => d.Attributes.Contains("class")&& d.Attributes["class"].Value.Contains("ss_workflow")).Descendants("tr").ElementAt(2).Descendants("td").ElementAt(1).InnerText,
                ToDateTime(Regex.Replace(htmLdoc.DocumentNode.Descendants("div").Where(d => d.Attributes.Contains("class")&& d.Attributes["class"].Value.Contains("ss_entryContent")).ElementAt(4).InnerText,@"[^\w\x20\.-:]",""),cultureBr),
                htmLdoc.DocumentNode.Descendants("td").First(d => d.Attributes.Contains("class")&& d.Attributes["class"].Value.Contains("ss_att_title")).Descendants("a").First().GetAttributeValue("href",""),
                endereco
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro: " + ex.Message);
                return null;
            }
        }

        // Carrega o dado de objeto
        public static string GetObjectId(DataTable dtTabela, string filename)
        {
            string strRetorno = "";
            string[] temp = GetFileNameWithoutExtension(filename)?.Split('-');
            strRetorno += GetPartialObjectId(dtTabela, "XX", temp?[0]);
            strRetorno += GetPartialObjectId(dtTabela, "VV", temp?[1].Substring(0, 2));
            strRetorno += ToInt32(temp?[1].Substring(2)).ToString("X2");
            strRetorno += GetPartialObjectId(dtTabela, "WW", temp?[2].Substring(1));
            strRetorno += GetPartialObjectId(dtTabela, "SUU", temp?[3]);
            strRetorno += GetPartialObjectId(dtTabela, "ZZ", temp?[4]);
            strRetorno += ToInt32(temp?[5]).ToString("X4");

            //strRetorno = dtTabela.
            return strRetorno;
        }

        // Retorna codigo especifico
        private static string GetPartialObjectId(DataTable dtTabela, string campo, string valor)
        {
            return dtTabela.Rows.Cast<DataRow>().FirstOrDefault(x => (x.Field<string>("Identificacao") == campo)
                        && (string.Equals(x.Field<string>("Valor"), valor, StringComparison.Ordinal))).Field<string>("Hex");
        }
    }

    // Classe de Analise dos Diagramas
    public class AnaliseDiagramas
    {
        // Atualiza os diagramas
        public static void AtualizaDiagramas()
        {
            // Conexão
            const string strConnection = "Server=tcp:ig1wvlolb4.database.windows.net,1433;Database=CMBM;User ID=alicio@ig1wvlolb4;Password=Sou-Aholyknight86;Trusted_Connection=False;"
                                         + "Encrypt=True;Max Pool Size=250;Connection Timeout=30; Pooling=true;";

            ConstrutivoData construtivo = new ConstrutivoData("Ricardo.gomes", "ricardo.gomes1528");
            AzureData azure = new AzureData(strConnection);

            DataTable dtTabela = construtivo.GetAllLinksNames("http://jundiai.construtivo.com/voith/buscaRapida.jsp?pasta=5061&nome=FVH-ECB-DI");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Dados dos Diagramas Copiados do Construtivo - " + dtTabela.Rows.Count);

            // Seleciona os diagramas da Base do Azure
            DataTable dtTabelaAzure = azure.GetTabelaAzure("SELECT [Nome] FROM [Arquivo];");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Dados Copiados da Base de Dados - " + dtTabelaAzure.Rows.Count);

            //Insere os Diagramas
            var dtTemp = dtTabela.AsEnumerable().Where(x => !dtTabelaAzure.AsEnumerable().Select(dr => dr.Field<string>("Nome")).Contains(x.Field<string>("Titulo")));
            if(dtTemp.Any()) Arquivos.UploadFromServer(azure, construtivo, dtTemp.CopyToDataTable(), "Arquivo");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Inserção dos Arquivos feita com sucesso - Inseridos " + (dtTemp.Any() ? dtTemp.Count().ToString() : "0") + " Arquivos");

            // Remove os diagramas
            var dtTemp2 = dtTabelaAzure.AsEnumerable().Where(x => !dtTabela.AsEnumerable().Select(dr => dr.Field<string>("Titulo")).Contains(x.Field<string>("Nome")));
            if (dtTemp2.Any()) azure.RemoveArquivos(dtTemp2.CopyToDataTable(), "Arquivo");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Remoção dos Arquivos feita com sucesso - Removidos " + (dtTemp2.Any() ? dtTemp2.Count().ToString() : "0") + " Arquivos");

            //Carrega os dados de Todos os arquivos
            DataTable dtDiagramas = azure.CarregaDiagramas("Arquivo");

            // Ajusta os dados dos diagramas
            dtDiagramas = AjustaDiagramas(dtDiagramas);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Carregados os dados dos diagramas - " + dtDiagramas.Rows.Count);
            AtualizaTabelaDiagrama(dtDiagramas, azure);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Atualização dos Arquivos dos Diagramas Concluída com Sucesso");
        }

        private static void AtualizaTabelaDiagrama(DataTable dtDiagramas, AzureData azure)
        {
            // Carrega a tabela do Azure
            DataTable dtTabelaAzure = azure.GetTabelaAzure("SELECT [Diagrama], [Revisão], [Código], [Origem], [Destino], [Formação], [Folha] FROM [Diagramas];");

            // Verifica quais dados não estão na lista e estão na base de dados
            DataTable dtTemp = Utilidades.ComparaTabelas(dtTabelaAzure.Copy(), dtDiagramas.Copy());

            // Remove os registros
            if (dtTemp.Rows.Count > 0) azure.RemoveTabelaDiagramas("Diagramas", dtTemp);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Foram Removidos " + dtTemp.Rows.Count + " Registros");

            // Verifica quais dados estão nos diagramas e não estão na base de dados
            dtTemp = Utilidades.ComparaTabelas(dtDiagramas.Copy(), dtTabelaAzure.Copy());

            // Adiciona os Registros
            if (dtTemp.Rows.Count > 0) azure.InsereTabelaDiagramas("Diagramas", dtTemp);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Foram Adicionados " + dtTemp.Rows.Count + " Registros");
        }
        
        // Copia os dados do diagrama
        public static DataTable DiagramaFromExcel(Stream arquivo, string nome)
        {
            // Cria a Tabela
            DataTable dtTabela = new DataTable();

            // Adiciona as colunas
            dtTabela.AddColumns(new[] { "Diagrama", "Revisão", "Código", "Origem", "Destino", "Formação", "Folha" });

            try
            {
                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(arquivo);
                //...
                //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                // Copia os dados para a tabela de resultado
                foreach (DataTable dt in result.Tables)
                {
                    if (dt.TableName.Contains("FL."))
                    {
                        for (int indice = 11; indice < 94; indice++)
                        {
                            if ((dt.Rows[indice].Field<string>("Column3") != null) && (dt.Rows[indice].Field<string>("Column3") != ""))
                            {
                                // Carrega os valores no array
                                object[] aDados =
                                    {
                                nome.Substring(0, 23), 
                                nome.Substring(24, 2),
                                (dt.Rows[indice].Field<string>("Column6") == "") || (dt.Rows[indice].Field<string>("Column6") == null)? "Vazio": dt.Rows[indice].Field<string>("Column6"),
                                dt.Rows[7].Field<string>("Column1").Contains(" ")?dt.Rows[7].Field<string>("Column1").
                                    Substring(0, dt.Rows[7].Field<string>("Column1").IndexOf(" ", StringComparison.Ordinal)):dt.Rows[7].Field<string>("Column1"),
                                (dt.Rows[indice].Field<string>("Column3") == "")||(dt.Rows[indice].Field<string>("Column3") == null) ? "Vazio":dt.Rows[indice].Field<string>("Column3"),
                                (dt.Rows[indice].Field<string>("Column7") == "")||(dt.Rows[indice].Field<string>("Column7") == null) ? "Vazio":dt.Rows[indice].Field<string>("Column7"),
                                dt.TableName
                            };

                                // Carrega os Dados na Tabela
                                dtTabela.Rows.Add(aDados);
                            }
                        }
                    }
                }

                return dtTabela;
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now + ": " + nome + " - Erro - " + ex.Message);

                return dtTabela;
            }
        }

        private static DataTable AjustaDiagramas(DataTable dtTabela)
        {
            // Remove as celulas vazias
            DataTable dtDiagramas = dtTabela.AsEnumerable().Where(s => s.Field<string>("Código") != "Vazio").CopyToDataTable();

            // Lista de Filtros
            List<string> strFilter = new List<string>
            {
                "QUADRO","CHAVE","SENSORES","BOMBAS","INSTRUMENTOS","PRESSOSTATO","VÁLVULA", "RESISTÊNCIAS","BOMBA","PRESSOSTATOS","TOMADA", "MOTOR","FECHAMENTO",
                    "EXAUSTOR","BARRAMENTO","Caixas","SALA","AR","LÂMINA","TRANSFORMADORES","Ar de Rebaixamento ","HOLD"
            };

            // Carrega os dados filtrados
            DataTable dtFilter = dtDiagramas.AsEnumerable().Where(row => strFilter.Contains(row.Field<string>("Origem"))).CopyToDataTable();

            // Remove os dados filtrados
            dtDiagramas = dtDiagramas.AsEnumerable().Where(row => !strFilter.Contains(row.Field<string>("Origem"))).CopyToDataTable();
            
            foreach (DataRow dr in dtFilter.Rows)
            {
                EnumerableRowCollection<DataRow> z = dtDiagramas.AsEnumerable().Where(row => row.Field<string>("Código") == dr.Field<string>("Código"));
                object[] obj = dr.ItemArray;
                if (z.Any())
                {
                    if (z.First().Field<string>("Origem") == dr.Field<string>("Destino")) obj[3] = z.First().Field<string>("Destino");
                    if (z.First().Field<string>("Destino") == dr.Field<string>("Destino")) obj[3] = z.First().Field<string>("Origem");
                }

                dtDiagramas.Rows.Add(obj);
            }
            
            // Carrega os dados filtrados
            dtFilter = dtDiagramas.AsEnumerable().Where(row => strFilter.Contains(row.Field<string>("Destino"))).CopyToDataTable();

            // Remove os dados filtrados
            dtDiagramas = dtDiagramas.AsEnumerable().Where(row => !strFilter.Contains(row.Field<string>("Destino"))).CopyToDataTable();

            foreach (DataRow dr in dtFilter.Rows)
            {
                EnumerableRowCollection<DataRow> z = dtDiagramas.AsEnumerable().Where(row => row.Field<string>("Código") == dr.Field<string>("Código"));
                object[] obj = dr.ItemArray;
                if (z.Any())
                {
                    if (z.First().Field<string>("Destino") == dr.Field<string>("Origem")) obj[4] = z.First().Field<string>("Origem");
                    if (z.First().Field<string>("Origem") == dr.Field<string>("Origem")) obj[4] = z.First().Field<string>("Destino");
                }

                dtDiagramas.Rows.Add(obj);
            }
            
            return dtDiagramas;
        }
    }

    // Classe de Analise das Listas
    public class AnaliseListas
    {
        //Atualiza as listas
        public static void AtualizaListas()
        {
            // Conexão
            string strConnection = "Server=tcp:ig1wvlolb4.database.windows.net,1433;Database=CMBM;User ID=alicio@ig1wvlolb4;Password=Sou-Aholyknight86;Trusted_Connection=False;"
                + "Encrypt=True;Max Pool Size=250;Connection Timeout=30; Pooling=true;";

            ConstrutivoData construtivo = new ConstrutivoData("Ricardo.gomes", "ricardo.gomes1528");
            AzureData azure = new AzureData(strConnection);


            DataTable dtTabela = construtivo.GetAllLinksNames("http://jundiai.construtivo.com/voith/buscaRapida.jsp?pasta=5061&nome=FVH-ECB-LC");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Dados das Listas Copiados do Construtivo - " + dtTabela.Rows.Count);

            // Seleciona as listas da Base do Azure
            DataTable dtTabelaAzure = azure.GetTabelaAzure("SELECT [Nome] FROM [Arquivo Listas];");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Dados Copiados das Listas da Base de Dados - " + dtTabelaAzure.Rows.Count);

            // Adiciona as listas
            var dtTemp = dtTabela.AsEnumerable().Where(x => !dtTabelaAzure.AsEnumerable().Select(dr => dr.Field<string>("Nome")).Contains(x.Field<string>("Titulo")));
            if(dtTemp.Any()) Arquivos.UploadFromServer(azure, construtivo, dtTemp.CopyToDataTable(), "Arquivo Listas");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Inserção dos Arquivos feita com sucesso - Inseridos " + (dtTemp.Any() ? dtTemp.Count().ToString() : "0") + " Arquivos");

            // Remove as listas
            var dtTemp2 = dtTabelaAzure.AsEnumerable().Where(x => !dtTabela.AsEnumerable().Select(dr => dr.Field<string>("Titulo")).Contains(x.Field<string>("Nome")));
            if (dtTemp2.Any()) azure.RemoveArquivos(dtTemp2.CopyToDataTable(), "Arquivo Listas");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Remoção dos Arquivos feita com sucesso - Removidos " + (dtTemp2.Any() ? dtTemp2.Count().ToString() : "0") + " Arquivos");

            //Carrega os dados de Todos os arquivos
            DataTable dtListas = azure.CarregaListas("Arquivo Listas");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Carregados os dados das Listas - " + dtListas.Rows.Count);

            // Atualiza as Listas
            AtualizaTabelaListas(dtListas, azure);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Atualização dos Arquivos dos Diagramas Concluída com Sucesso");
        }

        // Copia a lista de cabos do Word
        public static DataTable ListaFromWord(MemoryStream arquivo, string nome)
        {
            // Cria uma tabela
            DataTable dtTabela = new DataTable();
            dtTabela.AddColumns(new[] { "Lista", "Revisao", "Codigo", "Comprimento", "Formacao", "Percurso" });
            dtTabela.Columns["Comprimento"].DataType = typeof(int);

            // Carrega a tabela do word
            IWordDocument doc = new WordDocument(arquivo);
            IWTableCollection tabelas = doc.Sections[doc.Sections.Count - 1].Tables;
            IWTable tabela = tabelas[tabelas.Count - 1];

            // Carrega o total de linhas
            int totalLinhas = tabela.Rows.Count;

            // Cria o objeto
            object[] obj = new object[6];

            // Carrega as linhas
            for (int indice = 4; indice < totalLinhas; indice++)
            {
                obj[0] = nome.Substring(0, 23);
                obj[1] = nome.Substring(24, 2);
                obj[2] = Join("", tabela[indice, 0].Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray());
                obj[5] = Join("", tabela[indice, 3].Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray());

                // Verifica a ordem das colunas
                if (tabela[indice, 1].ContainsAny(new[] {"x", "fabricante", "fibra"},false))
                {
                    obj[3] = ToInt32(Join("", tabela[indice, 2].Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray()));
                    obj[4] = Join("", tabela[indice, 1].Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray());
                }
                else
                {
                    obj[3] = ToInt32(Join("", tabela[indice, 1].Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray()));
                    obj[4] = Join("", tabela[indice, 2].Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray());
                }
                // Adciona na tabela
                dtTabela.Rows.Add(obj);
            }

            // Retorna a a tabela
            return dtTabela;
        }

        // Atualiza tabela de lista de cabos
        private static void AtualizaTabelaListas(DataTable dtListas, AzureData azure)
        {
            // Carrega a tabela do Azure
            DataTable dtTabelaAzure = azure.GetTabelaAzure("SELECT [Lista], [Revisão] AS [Revisao], [Código Cabo] AS [Codigo], [Comprimento], [Formação] AS [Formacao]" 
                +", [Percurso] FROM [Listas Lancamento];");

            // Verifica quais dados não estão na lista e estão na base de dados
            DataTable dtTemp = Utilidades.ComparaTabelas(dtTabelaAzure.Copy(), dtListas.Copy());

            // Remove os registros
            if (dtTemp.Rows.Count > 0) azure.RemoveTabelaListas("Listas Lancamento", dtTemp);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Foram Removidos " + dtTemp.Rows.Count + " Registros");

            // Verifica quais dados estão nas listas e não estão na base de dados
            dtTemp = Utilidades.ComparaTabelas(dtListas.Copy(), dtTabelaAzure.Copy());

            // Adiciona os registros
            if (dtTemp.Rows.Count > 0) azure.InsereTabelaListas("Listas Lancamento", dtTemp);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Foram Adicionados " + dtTemp.Rows.Count + " Registros");
        }
    }

    // Classe de Analise de Projetos
    public class AnaliseProjetos
    {
        //Atualiza dados dos Projetos
        public static void AtualizaProjetos()
        {
            // Conexão
            string strConnection = "Server=tcp:ig1wvlolb4.database.windows.net,1433;Database=CMBM;User ID=alicio@ig1wvlolb4;Password=Sou-Aholyknight86;Trusted_Connection=False;"
                + "Encrypt=True;Max Pool Size=250;Connection Timeout=30; Pooling=true;";

            ConstrutivoData construtivo = new ConstrutivoData("Ricardo.gomes", "ricardo.gomes1528");
            AzureData azure = new AzureData(strConnection);

            // Copia os dados de todos os arquivos
            List<string> lista = construtivo.CopiaApenasLinks("255").Result;
            lista = lista.Union(construtivo.CopiaApenasLinks("278").Result).ToList();
            lista = lista.Union(construtivo.CopiaApenasLinks("294").Result).ToList();
            
            // Carrega a tabela
            DataTable dtAzure = azure.GetTabelaAzure("SELECT DISTINCT [Link Visualizar] FROM [Dados Projetos];");

            // Carrega a lista de links
            List<string> listaAzure = dtAzure.Rows.Cast<DataRow>().Select(row => row["Link Visualizar"].ToString()).ToList();

            // Seleciona os valores inexistentes
            IEnumerable<string> t = lista.Where(link => !listaAzure.Contains(link));
            
            DataTable dtCodigos = azure.GetTabelaAzure("SELECT * FROM [Codificacao];");

            // Cria a tabela
            DataTable dtTabela = new DataTable();
            dtTabela.AddColumns(new[] { "Nrs", "Titulo", "Descricao", "Autor", "Tamanho", "Estado", "Data", "Link", "Link Visualizar", "objectId", "versionId" });
            dtTabela.Columns["Data"].DataType = typeof(DateTime);

            // Carrega os dados
            foreach (string str in t) dtTabela.Rows.AddNotNull(construtivo.GetDocumentData(str)?.ToList().AddIds(dtCodigos).ToArray());

            // Adiciona os dados
            if (dtTabela.Rows.Count > 0) azure.AdicionaDadosProjetos(dtTabela, "Dados Projetos");

            // Informa o usuario
            Console.WriteLine(DateTime.Now + ": Atualização dos Projetos feita com sucesso - Inseridos " + dtTabela.Rows.Count + " Registros");

            // Atualiza Lista de PDFs
            AtualizaPdFs(azure, construtivo);

            // Atualiza Lista de Obsoletos
            AtualizaObsoletos(azure, construtivo);

            // Ajusta os Dados
            AjustaDados(construtivo, azure);
        }

        // Ajusta a base de dados
        public static void AjustaDados(ConstrutivoData construtivo, AzureData azure)
        {
            // Executa comando
            azure.ExecutaComando("DELETE FROM [Dados Obsoletos] WHERE [Link Visualizar] IN (SELECT [Link Visualizar] FROM [Dados Projetos] WHERE [Link Visualizar] IN "
                                 + " (SELECT [Link Visualizar] FROM [Dados Obsoletos])); ");

            // Executa Comando
            azure.ExecutaComando("DELETE FROM [Primeira Emissao] WHERE [Link Visualizar] IN (SELECT [Link Visualizar] FROM [Primeira Emissao] WHERE [Link Visualizar] IN"
                                 + " (SELECT [Link Visualizar] FROM [Dados Obsoletos])); ");

            // Carrega a tabela dados
            DataTable dtAzure = azure.GetTabelaAzure("SELECT [Link Visualizar] FROM [Dados Projetos] WHERE Descricao = '' "
                + "UNION SELECT [Link Visualizar] FROM [Dados Projetos] WHERE Autor = '' UNION SELECT [Link Visualizar] FROM [Dados Projetos] WHERE Tamanho= '' "
                    + "UNION SELECT [Link Visualizar] FROM [Dados Projetos] WHERE Estado = '' UNION SELECT [Link Visualizar] FROM [Dados Projetos] WHERE Data= '';");
            
            //Cria a tabela
            DataTable dtTabela = new DataTable();
            dtTabela.AddColumns(new[] { "Nrs", "Titulo", "Descricao", "Autor", "Tamanho", "Estado", "Data", "Link", "Link Visualizar","objectId", "versionId" });
            dtTabela.Columns["Data"].DataType = typeof(DateTime);

            // Carrega os dados
            foreach (DataRow row in dtAzure.Rows) dtTabela.Rows.AddNotNull(construtivo.GetDocumentData(row.Field<string>("Link Visualizar")));

            // Adiciona os dados
            if (dtTabela.Rows.Count > 0) azure.AtualizaDadosProjetos(dtTabela, "Dados Projetos");

            // Informa o usuario
            Console.WriteLine(DateTime.Now + ": Atualização dos Projetos feita com sucesso - Inseridos " + dtTabela.Rows.Count + " Registros");
        }

        // Atualiza os PDFs
        private static void AtualizaPdFs(AzureData azure, ConstrutivoData construtivo)
        {
            // Carrega os Dados do Azure
            DataTable dtAzure = azure.GetTabelaAzure("SELECT Nrs, Titulo, Descricao, Autor, Tamanho, Estado, Data, [Link Baixar], Temp.[Link Visualizar] FROM "
                    + "(SELECT Nrs, Titulo, Descricao, Autor, Tamanho, Estado, Data, [Link Baixar], [Link Visualizar] FROM[Dados Projetos] "
                        + "UNION SELECT Nrs, Titulo, Descricao, Autor, Tamanho, Estado, Data, [Link Baixar], [Link Visualizar] FROM[Dados Obsoletos]) AS Temp "
                            + "LEFT JOIN[Dados PDF] ON Temp.[Link Visualizar] = [Dados PDF].[Link Visualizar] WHERE[Dados PDF].[Link Visualizar] IS NULL;");

            // Carrega os dados dos PDFs
            DataTable dtTEmp2 = construtivo.GetPdFs(dtAzure.Copy());

            // Adiciona os dados 
            azure.AdicionaDadosPdf(dtTEmp2, "Dados PDF");

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Atualização dos PDFs feita com sucesso - Inseridos " + dtTEmp2.Rows.Count + " Registros");
        }

        // Atualiza os Obsoletos
        private static void AtualizaObsoletos(AzureData azure, ConstrutivoData construtivo)
        {
            try
            {
                // Carrega os dados do Azure
                DataTable dtAzure = azure.GetTabelaAzure("SELECT [Dados Projetos].[Link Visualizar] FROM [Dados Projetos] LEFT JOIN "
                + "(SELECT DISTINCT [Link Pai] AS [Link Visualizar] FROM [Dados Obsoletos] UNION SELECT [Link Visualizar] FROM [Primeira Emissao])  AS Junto "
                    + "ON [Dados Projetos].[Link Visualizar] = [Junto].[Link Visualizar] WHERE [Junto].[Link Visualizar] IS NULL;");

                // Carrega a segunda tabela
                DataTable dtAzure2 = azure.GetTabelaAzure("SELECT Titulo FROM [Dados Obsoletos];");
            
                // Carrega os Dados Obsoletos
                DataTable dtTabelaObsoleto = GetAllObsoleteData(construtivo, dtAzure.Copy(), azure).Result;

                // Carrega a lista de dados
                List<string> listaAzure = dtAzure2.Rows.Cast<DataRow>().Select(row => row["Titulo"].ToString()).ToList();

                // Remove os dados não listados
                EnumerableRowCollection<DataRow> y = dtTabelaObsoleto.AsEnumerable().Where(row => listaAzure.Contains(row.Field<string>("Titulo")));

                // Se existem dados atualiza os dados
                if (y.Any()) azure.UpdateDadosObsoletos(y.CopyToDataTable(), "Dados Obsoletos");

                // Remove os dados já listados
                EnumerableRowCollection<DataRow> t = dtTabelaObsoleto.AsEnumerable().Where(row => !listaAzure.Contains(row.Field<string>("Titulo")));

                // Se existem dados insere os dados
                if (t.Any()) azure.AdicionaDadosProjetos(t.CopyToDataTable(), "Dados Obsoletos");

                // Informa o usuario 
                Console.WriteLine(DateTime.Now + ": Atualização dos Obsoletos feita com sucesso - Inseridos " + dtTabelaObsoleto.Rows.Count + " Inseridos");
            }
            catch (Exception ex)
            {
                // Informa o usuario e incrementa
                Console.WriteLine(DateTime.Now + ": Erro - " + ex.Message);
            }
        }

        //Carrega todos os dados obesoletos de uma lista
        private static async Task<DataTable> GetAllObsoleteData(ConstrutivoData construtivo, DataTable dtEntrada, AzureData azure)
        {
            // Cria a tabela
            DataTable dtTabela = new DataTable();
            dtTabela.AddColumns(new[] { "Nrs", "Titulo", "Descricao", "Autor", "Tamanho", "Estado", "Data", "Link", "Link Visualizar", "objectId", "versionId" });
            dtTabela.Columns["Data"].DataType = typeof(DateTime);

            // Cria uma lista para armazenar as tarefas
            List<Task<DataTable>> dtTabelaTask2 = new List<Task<DataTable>>();

            // Carrega os códigos 
            DataTable dtTemp = construtivo.GetCodes(dtEntrada);

            // Carrega os dados obsoletos
            foreach (DataRow dr in dtTemp.Rows)
            {
                // Verifica se existem slots de download disponível (10 slots no total)      
                while (dtTabelaTask2.Count(dtT => (dtT.Status != TaskStatus.RanToCompletion) && (dtT.Status != TaskStatus.Faulted)) > 10)   Thread.Sleep(100);

                // Adiciona a tarefa
                dtTabelaTask2.Add(Task<DataTable>.Factory.StartNew(() => GetObsoleteData(construtivo, dr.Field<string>("Link Visualizar"), azure)));
            }

            // Aguarda o Término de todas as tabelas
            DataTable[] dtTabelas2 = await Task.WhenAll(dtTabelaTask2);

            // Adiciona as novas tabelas
            foreach (DataTable dt in dtTabelas2) dtTabela.Merge(dt);

            // Retorna a tabela
            return dtTabela;
        }

        // Carrega os dados obsoletos
        private static DataTable GetObsoleteData(ConstrutivoData construtivo, string linkpai, AzureData azure)
        {
            DataTable dtCodigos = azure.GetTabelaAzure("SELECT * FROM [Codificacao];");

            // Cria a tabela
            DataTable dtTabela = new DataTable();
            dtTabela.AddColumns(new[] { "Nrs", "Titulo", "Descricao", "Autor", "Tamanho", "Estado", "Data", "Link", "Link Visualizar", "objectId", "versionId" });
            dtTabela.Columns["Data"].DataType = typeof(DateTime);
            
            string temp = linkpai.Replace("http://jundiai.construtivo.com/ssf/a/c/p_name/ss_forum/p_action/1/action/view_folder_entry/binderId/", "");

            // Carrega os links dos arquivos obsoletos
            string[] lista = construtivo.GetObsoleteLinks(temp.Split('/')[0], temp.Split('/')[2]);

            // Carrega os dados obsoletos
            foreach (string str in lista) dtTabela.Rows.AddNotNull(construtivo.GetDocumentData(str)?.ToList().AddIds(dtCodigos).ToArray());
           
            // Retorna
            return dtTabela;
        }
    }

    // Classe de Anaise das Listas
    public class AnaliseMaterial
    {
        // Copia as lista de material para tabela
        public static void ListaMaterial()
        {
            // Conexão
            string strConnection = "Server=tcp:ig1wvlolb4.database.windows.net,1433;Database=CMBM;User ID=alicio@ig1wvlolb4;Password=Sou-Aholyknight86;Trusted_Connection=False;"
                + "Encrypt=True;Max Pool Size=250;Connection Timeout=30; Pooling=true;";

            // Abre a conexão
            AzureData azure = new AzureData(strConnection);

            // carrega tabelas
            DataTable dtTabela = azure.CarregaLMs();

            // Atualiza as LMs
            AtualizaTabelaLMs(dtTabela, azure);
        }

        // Copia a lista de cabos do Word
        public static DataTable ListaMaterialFromWord(Stream arquivo, string nome)
        {
            // Cria uma tabela
            DataTable dtTabela = new DataTable();
            dtTabela.AddColumns(new [] { "Lista" , "Revisao" , "Item" , "CodigoCopem", "Quantidade" });

            try
            {
                // Carrega a tabela do word
                WordDocument doc = new WordDocument(arquivo);
                List<IWTable> filtro = doc.Sections.Cast<WSection>().Select(s => s.Tables.Cast<IWTable>()).SelectMany(x => x).Where(s => s[0, 0].Contains("item", false)).ToList();

                // Declara so indices
                int indiceCopem = 0;
                int indiceQtd = 0;
                
                if (filtro.Count > 0)
                {
                    int intTotalColunas = filtro[0].FirstRow.Cells.Count;

                    for (int indice = 1; indice < intTotalColunas; indice++)
                    {
                        if (filtro[0][0, indice].ContainsAny(new[] { "copem" , "projeto"}, true)||filtro[0][1, indice].ContainsAny(new[] { "copem", "projeto" }, true))
                            indiceCopem = indice;

                        if (filtro[0][0, indice].Contains("qtd", false) || filtro[0][1, indice].Contains("qtd", false)) indiceQtd = indice;
                    }
                }
                else
                {
                    filtro = doc.Sections.Cast<WSection>().Where(q => q.HeadersFooters.Header.Tables.Count > 0)
                        .Where(r => r.HeadersFooters.Header.Tables[0].LastCell.Contains("obser", false)).Select(s => s.Tables.Cast<IWTable>())                            .SelectMany(x => x).ToList();

                    IWTable ttemp = doc.Sections.Cast<WSection>().Where(q => q.HeadersFooters.Header.Tables.Count > 0)
                            .First(r => r.HeadersFooters.Header.Tables[0].LastCell.Contains("obser", false)).HeadersFooters.Header.Tables[0];

                    for (int indice = 1; indice < ttemp.LastRow.Cells.Count; indice++)
                    {
                        if (ttemp.LastRow.Cells[indice].ContainsAny(new[] { "copem", "projeto" }, true)) indiceCopem = indice;

                        if (ttemp.LastRow.Cells[indice].Contains("qtd", false)) indiceQtd = indice;
                    }
                }

                foreach (IWTable dt in filtro)
                {
                    // Carrega as linhas
                    IEnumerable<WTableRow> temp = dt.Rows.Cast<WTableRow>().Where(x => x.Cells[0].NotEquals("",false) && x.Cells[0].NotEquals("ITEM", false))
                        .Where(y => y.Cells[indiceQtd].NotEquals("",false));

                    foreach (WTableRow dr in temp)
                    {
                        object[] obj =
                        {
                            nome.Substring(0, 23), nome.Substring(24, 2), dr.Cells[0].Paragraphs[0].Text,
                            indiceCopem == 0 ? "" : dr.Cells[indiceCopem].Paragraphs[0].Text, dr.Cells[indiceQtd].Paragraphs[0].Text
                        };

                        dtTabela.Rows.Add(obj);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now + ": " + nome + " - " + ex.Message);
            }

            // Retorna a a tabela
            return dtTabela;
        }

        // Atualiza listas de material
        private static void AtualizaTabelaLMs(DataTable dtListas, AzureData azure)
        {
            // Carrega a tabela do Azure
            DataTable dtTabelaAzure = azure.GetTabelaAzure("SELECT [Lista], [Revisão] AS [Revisao], [Item], [Código Copem] AS [CodigoCopem], [Quantidade] FROM "
                +"[Listas Material];");

            // Verifica quais dados não estão na lista e estão na base de dados
            DataTable dtTemp = Utilidades.ComparaTabelas(dtTabelaAzure.Copy(), dtListas.Copy());

            // Remove Registros
            if (dtTemp.Rows.Count > 0) azure.RemoveTabelaListas("Listas Material", dtTemp);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Foram Removidos " + dtTemp.Rows.Count + " Registros");

            // Verifica quais dados estão nas listas e não estão na base de dados
            dtTemp = Utilidades.ComparaTabelas(dtListas.Copy(), dtTabelaAzure.Copy());

            // Adiciona os registros
            if (dtTemp.Rows.Count > 0) azure.InsereTabelaListas("Listas Material", dtTemp);

            // Informa o usuário
            Console.WriteLine(DateTime.Now + ": Foram Adicionados " + dtTemp.Rows.Count + " Registros");
        }
    }
    
    // Classe de Manipulação de Arquivos
    public class Arquivos
    {
        // Atualiza Todos os arquivos
        public static void AtualizaArquivo()
        {
            string strConnection = "Server=tcp:ig1wvlolb4.database.windows.net,1433;Database=CMBM;User ID=alicio@ig1wvlolb4;Password=Sou-Aholyknight86;Trusted_Connection=False;"
                + "Encrypt=True;Connection Timeout=30;";

            // Cria os acessos
            ConstrutivoData construtivo = new ConstrutivoData("Ricardo.gomes", "ricardo.gomes1528");
            AzureData azure = new AzureData(strConnection);

            // Carrega a tabela
            DataTable dtTabela = azure.GetTabelaAzure("SELECT DISTINCT Dados.Titulo, Dados.[Link Baixar] AS Links FROM (SELECT PDF AS Titulo, [PDF Link] AS [Link Baixar] "
                + " FROM [Dados PDF] UNION SELECT Titulo, [Link Baixar] FROM [Dados Projetos] UNION SELECT Titulo, [Link Baixar] FROM [Dados Obsoletos]) AS Dados "
                    + " LEFT JOIN [Arquivo Projetos] ON [Arquivo Projetos].Nome = Dados.Titulo WHERE [Arquivo Projetos].Nome IS NULL AND Dados.[Link Baixar] <> '';");

            // Atualiza base de arquivos
            UploadFromServer(azure, construtivo, dtTabela, "Arquivo Projetos");
        }

        // Atualiza arquivos via Servidor
        public static void UploadFromServer(AzureData azure, ConstrutivoData construtivo, DataTable dtTabela, string tabela)
        {
            foreach (DataRow dr in dtTabela.Rows)
                azure.InsereArquivo(construtivo.BaixaArquivo(construtivo.GetDocumentData(dr.Field<string>("Links"))[7] as string), dr.Field<string>("Titulo"),tabela);
        }
    }
    
    // classe de requisicoes
    public class Requisicoes
    {
        // Envia Solicitação
        public static void GetRequestStreamCallback(IAsyncResult callbackResult)
        {
            // Cria a requisição
            RequestState hwState = (RequestState)callbackResult.AsyncState;

            // Chama a resposta de modo assincrono
            try
            {
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
            
                hwState.HwrRequest.BeginGetResponse(GetResponseStreamCallback, hwState);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Erro: " + ex.Message);

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

                //Transforma o fluxo em String
                using (Stream responseStream = hwState.HwResponse.GetResponseStream())
                {
                    if (responseStream != null)
                    {
                        responseStream.CopyTo(hwState.OResposta);

                        // Requisicao completada OK
                        hwState.BErro = false;
                    }
                }

                // Seta o All Done
                hwState.AllDone.Set();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro: " + ex.Message);

                // Seta o All Done
                hwState.AllDone.Set();
            }
        }
    }

    // Utilidades
    public class Utilidades
    {
        // Compara duas Tabelas
        public static DataTable ComparaTabelas(DataTable first, DataTable second)
        {
            first.TableName = "FirstTable";
            second.TableName = "SecondTable";

            //Create Empty Table
            DataTable table = new DataTable("Difference");

            //Must use a Dataset to make use of a DataRelation object
            using (DataSet ds = new DataSet())
            {
                //Add tables
                ds.Tables.AddRange(new[] { first.Copy(), second.Copy() });

                //Get Columns for DataRelation
                DataColumn[] firstcolumns = new DataColumn[ds.Tables[0].Columns.Count];
                for (int i = 0; i < firstcolumns.Length; i++)  firstcolumns[i] = ds.Tables[0].Columns[i];

                DataColumn[] secondcolumns = new DataColumn[ds.Tables[1].Columns.Count];
                for (int i = 0; i < secondcolumns.Length; i++) secondcolumns[i] = ds.Tables[1].Columns[i];

                //Create DataRelation
                DataRelation r = new DataRelation(Empty, firstcolumns, secondcolumns, false);
                ds.Relations.Add(r);

                //Create columns for return table
                for (int i = 0; i < first.Columns.Count; i++)  table.Columns.Add(first.Columns[i].ColumnName, first.Columns[i].DataType);

                //If First Row not in Second, Add to return table.
                table.BeginLoadData();

                foreach (DataRow parentrow in ds.Tables[0].Rows)
                {
                    DataRow[] childrows = parentrow.GetChildRows(r);
                    if (childrows.Length == 0) table.LoadDataRow(parentrow.ItemArray, true);
                }

                table.EndLoadData();
            }

            return table;
        }
    }
    
    public static class UtilitiesExtensions
    {
        public static List<object> AddIds(this List<object> obj, DataTable dtCodigos)
        {
            DateTime offset = new DateTime(1900, 1, 1);
            const long divider = TimeSpan.TicksPerSecond;

            obj.Add(HtmlProcess.GetObjectId(dtCodigos, obj[1] as string));
            obj.Add(((((DateTime)obj[6]).Ticks - offset.Ticks) / divider).ToString("X2") + GetFileNameWithoutExtension(obj[1].ToString()).Split('-')[6]);

            return obj;
        }

        // Adiciona colunas em uma tabela
        public static void AddColumns(this DataTable dtTabela, string[] lista)
        {
            foreach (string str in lista) dtTabela.Columns.Add(str);
        }

        // Adiciona os parametros
        public static void AddicionaParametros(this SqlCommand entrada, string[] parametros, int[] tamanho)
        {
            int total = parametros.Length;
            if (parametros.Length == tamanho.Length )
                for (int indice = 0; indice < total; indice++) entrada.Parameters.Add("@" + parametros[indice], VarChar, tamanho[indice], parametros[indice]);
        }

        // Verifica se a celula contem o valor
        public static bool Contains(this WTableCell celula, string valor, bool todosparagrafos)
        {
            return todosparagrafos ? Join(" ", celula.Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray()).ToLower().Contains(valor) 
                : celula.Paragraphs[0].Text.ToLower().Contains(valor);
        }

        // Verifica se a celula é igual
        public static bool Equals(this WTableCell celula, string valor, bool todosparagrafos)
        {
            return todosparagrafos ? Join(" ", celula.Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray()) == (valor) : celula.Paragraphs[0].Text == valor;
        }

        // Verifica se a celula é diferente
        public static bool NotEquals(this WTableCell celula, string valor, bool todosparagrafos)
        {
            return todosparagrafos ? Join(" ", celula.Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray()) != (valor) : celula.Paragraphs[0].Text != valor;
        }

        // Verifica se contem algum
        public static bool ContainsAny(this WTableCell celula, string[] valor, bool todosparagrafos)
        {
            bool retorno = false;

            foreach (string str in valor)
            {
                if (todosparagrafos) if (Join(" ", celula.Paragraphs.Cast<WParagraph>().Select(s => s.Text).ToArray()).ToLower().Contains(str)) retorno = true;

                if(!todosparagrafos) if (celula.Paragraphs[0].Text.ToLower().Contains(str)) retorno = true;
            }

            return retorno;
        }

        // Adiciona linha se não for nulo
        public static bool AddNotNull(this DataRowCollection rows, object[] array)
        {
            if (array != null) rows.Add(array);
            return (array != null);
        }
    }
}