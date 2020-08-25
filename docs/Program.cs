using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace docs
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Processa(args);
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null)
                {
                    ex = ex.InnerException;
                }

                Console.WriteLine("Erro ao executar o comando");
                Console.WriteLine(ex.Message);
            }
            
            
        }

        private static void Processa(string[] args)
        {
            string basePath = null;
            string novoPath = null;
            string novoArquivo = null;

            if (args.Count() > 0)
            {
                novoArquivo = args[0];

                if (args[0].ToLower() == "--ajuda")
                {
                    MostrarAjuda();
                    return;
                }
            }

            if (args.Count() > 1)
            {
                basePath = args[1];
            }

            if (args.Count() > 2)
            {
                novoPath = args[2];
            }

            if (string.IsNullOrEmpty(novoArquivo))
            {
                throw new Exception("Arquivo a ser documentado precisa ser passado no primeiro parâmetro");
            }

            if (string.IsNullOrEmpty(basePath))
            {
                basePath = Environment.GetEnvironmentVariable("DOCS_BASE_PATH");
            }

            if (string.IsNullOrEmpty(novoPath))
            {
                novoPath = Environment.GetEnvironmentVariable("DOCS_NOVO_PATH");
            }

            if (!string.IsNullOrEmpty(novoPath) && Path.GetDirectoryName(novoArquivo).Length == 0)
            {
                novoArquivo = novoPath + "\\" + novoArquivo;
            }

            // Criar classe com o caminho para onde os xlms de origem estão
            var d = new MetaFinder(basePath);

            // Documentar fonte que tem com dados de dw e sem dados de delta
            d.Documentar(novoArquivo);

            Console.WriteLine("Documento documentado com sucesso");

            if (Debugger.IsAttached)
            {
                Console.ReadKey();
            }
        }

        private static void MostrarAjuda()
        {
            var ajudaStr =

@"    Este programa preenche os dados da parte do desenvolvedor na documentação baseando-se em
    planilhas de documentação que já foram preenchidas, a variável de ambiente DOCS_BASE_PATH 
    poder ser configurada para indicar onde as planilhas já preenchidas se encontram, senão, o 
    caminho para as planilhas preenchidas precisará se passado no segundo parâmetro do comando.
    
    Exemplo sem configurar DOCS_BASE_PATH: 
       docs <nome da planilha a ser preenchida> <path da planilhas preenchidas>
    
    Exemplo configurarando DOCS_BASE_PATH 
       docs <nome da planilha a ser preenchida>

    Opcionalmente a variável de ambiente DOCS_NOVO_PATH pode ser configurada para indicar o caminho 
    para <planilha a ser preenchida> evitando ter que especificar o caminho completo.
";


            Console.WriteLine(ajudaStr);
        }
    }
}
