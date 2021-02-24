using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace corruptfile
{
    public class Archivos
    {
        public string Nombre { get; set; }
        public string Extension { get; set; }
    }
    public class GenerateArchivos
    {
        readonly List<Archivos> ListaArchivos = new List<Archivos>();
        public GenerateArchivos()
        {
            ListaArchivos.Add(wordDoc);
            ListaArchivos.Add(powerDoc);
            ListaArchivos.Add(excelDoc);
            ListaArchivos.Add(pdfDoc);
        }

        public List<Archivos> GetArchivos()
        {
            return ListaArchivos;
        }
        readonly Archivos wordDoc = new Archivos()
        {
            Nombre = "Word",
            Extension = "docx"
        };
        readonly Archivos powerDoc = new Archivos()
        {
            Nombre = "Power Point",
            Extension = "pptx"
        };
        readonly Archivos excelDoc = new Archivos()
        {
            Nombre = "Excel",
            Extension = "xlsx"
        };
        readonly Archivos pdfDoc = new Archivos()
        {
            Nombre = "PDF",
            Extension = "pdf"
        };

    }
}
