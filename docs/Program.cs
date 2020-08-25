using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace docs
{
    class Program
    {
        static void Main(string[] args)
        {
            // Criar classe com o caminho para onde os xlms de origem estão
            var d = new MetaFinder(@"C:\docs\base");

            // Documentar fonte que tem com dados de dw e sem dados de delta
            d.Documentar(@"C:\docs\novos\Análise de crédito (DW) novo.xlsx");

            Console.WriteLine("Documento documentado com sucesso");
            Console.ReadKey();
        }
    }
}
