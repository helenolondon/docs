using DocumentFormat.OpenXml.Presentation;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace docs
{
    public class MetaFinder
    {
        private readonly string basePath = @"base";
        private List<DocInfo> _metaDb  = new List<DocInfo>();

        public MetaFinder(string xlsxBasePath)
        {
            this.basePath = xlsxBasePath;
        }

        private List<String> ListaArquivos()
        {
            DirectoryInfo d = new DirectoryInfo(basePath);
            FileInfo[] Files = d.GetFiles("*.xlsx");
            var ListaArquivos = new List<String>();

            foreach (FileInfo file in Files)
            {
                ListaArquivos.Add(file.FullName);
            }

            return ListaArquivos;
        }

        private void Extraia(string arquivo)
        {
            SLDocument sl = new SLDocument(arquivo, "Planilha1");
            string campoDW;
            string tabelaDW;
            DocInfo doc;
            string oF, oH, oG;

            for(int y = 7; y < 500; y++)
            {
                campoDW = sl.GetCellValueAsString($"D{y}");
                tabelaDW = sl.GetCellValueAsString($"E{y}");

                if((tabelaDW.ToUpper() == "DIM_DATA") || string.IsNullOrEmpty(campoDW) || string.IsNullOrEmpty(tabelaDW))
                {
                    continue;
                }

                oF = sl.GetCellValueAsString($"F{y}");
                oH = sl.GetCellValueAsString($"H{y}");
                oG = sl.GetCellValueAsString($"G{y}");

                if(string.IsNullOrEmpty(oF) && string.IsNullOrEmpty(oH) && string.IsNullOrEmpty(oG))
                {
                    continue;
                }

                doc = this._metaDb.Where(m => m.TabelaDW == tabelaDW && m.CampoDW == campoDW).FirstOrDefault();

                if(doc == null)
                {
                    doc = new DocInfo(campoDW, tabelaDW);
                    this._metaDb.Add(doc);
                }

                if (!string.IsNullOrEmpty(oF))
                {
                    doc.CampoOrigem = oF;
                }

                if (!string.IsNullOrEmpty(oH))
                {
                    doc.TabelaOrigem = oH;
                }

                if (!string.IsNullOrEmpty(oG))
                {
                    doc.BaseOrigem = oG;
                }
            }
        }

        private void PopMetaDb()
        {
            foreach (string f in this.ListaArquivos())
            {
                Extraia(f);
            }
        }

        public void Documentar(string novoDocumento)
        {
            if (this._metaDb.Count() == 0)
            {
                this.PopMetaDb();
            }

            SLDocument sl = new SLDocument(novoDocumento, "Planilha1");
            string campoDW;
            string tabelaDW;
            DocInfo doc;
            bool docChanged = false;

            for (int y = 7; y < 500; y++)
            {
                campoDW = sl.GetCellValueAsString($"D{y}");
                tabelaDW = sl.GetCellValueAsString($"E{y}");

                doc = this._metaDb.Where(m => m.TabelaDW == tabelaDW && m.CampoDW == campoDW).FirstOrDefault();

                if (doc != null)
                {
                    sl.SetCellValue($"F{y}", doc.CampoOrigem);
                    sl.SetCellValue($"H{y}", doc.TabelaOrigem);
                    sl.SetCellValue($"G{y}", doc.BaseOrigem);

                    docChanged = true;
                }
            }

            if (docChanged)
            {
                sl.Save();
            }
        }
    }
}
