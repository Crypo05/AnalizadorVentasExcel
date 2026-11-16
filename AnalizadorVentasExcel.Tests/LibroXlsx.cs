using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace AnalizadorVentasExcel.Tests
{
    /// <summary>
    /// Escribe un .xlsx mínimo para las pruebas: texto como inlineStr, números como valor.
    /// Un valor null deja la celda vacía. La estructura de partes es la del generador de referencia.
    /// </summary>
    internal static class LibroXlsx
    {
        /// <summary>Atajo para escribir una fila: Fila("a", 1, null).</summary>
        public static object?[] Fila(params object?[] celdas) => celdas;

        /// <summary>Un libro de una sola hoja ("Sheet"). Devuelve la ruta completa.</summary>
        public static string Escribir(string carpeta, string nombreArchivo, params object?[][] filas)
            => Escribir(carpeta, nombreArchivo, ("Sheet", filas));

        /// <summary>Un libro con varias hojas, en el orden dado.</summary>
        public static string Escribir(string carpeta, string nombreArchivo, params (string Nombre, object?[][] Filas)[] hojas)
        {
            Directory.CreateDirectory(carpeta);
            string ruta = Path.Combine(carpeta, nombreArchivo);

            using var archivo = new FileStream(ruta, FileMode.Create, FileAccess.Write, FileShare.None);
            using var zip = new ZipArchive(archivo, ZipArchiveMode.Create);

            var ct = new StringBuilder();
            ct.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            ct.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            ct.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            ct.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            ct.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (int i = 0; i < hojas.Length; i++)
                ct.Append($"<Override PartName=\"/xl/worksheets/sheet{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            ct.Append("</Types>");
            Parte(zip, "[Content_Types].xml", ct.ToString());

            Parte(zip, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>");

            var wb = new StringBuilder();
            wb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            wb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
            var rels = new StringBuilder();
            rels.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            rels.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < hojas.Length; i++)
            {
                wb.Append($"<sheet name=\"{Esc(hojas[i].Nombre)}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
                rels.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
                Parte(zip, $"xl/worksheets/sheet{i + 1}.xml", XmlHoja(hojas[i].Filas));
            }
            wb.Append("</sheets></workbook>");
            rels.Append("</Relationships>");
            Parte(zip, "xl/workbook.xml", wb.ToString());
            Parte(zip, "xl/_rels/workbook.xml.rels", rels.ToString());

            return ruta;
        }

        private static string XmlHoja(object?[][] filas)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            for (int r = 0; r < filas.Length; r++)
            {
                sb.Append($"<row r=\"{r + 1}\">");
                for (int c = 0; c < filas[r].Length; c++)
                {
                    object? v = filas[r][c];
                    if (v is null) continue;
                    string celda = Columna(c) + (r + 1).ToString(CultureInfo.InvariantCulture);
                    if (v is IFormattable f && v is not string)
                        sb.Append($"<c r=\"{celda}\"><v>{f.ToString(null, CultureInfo.InvariantCulture)}</v></c>");
                    else
                        sb.Append($"<c r=\"{celda}\" t=\"inlineStr\"><is><t>{Esc(Convert.ToString(v, CultureInfo.InvariantCulture) ?? "")}</t></is></c>");
                }
                sb.Append("</row>");
            }
            sb.Append("</sheetData></worksheet>");
            return sb.ToString();
        }

        private static string Columna(int i)
        {
            string s = "";
            i++;
            while (i > 0)
            {
                int r = (i - 1) % 26;
                s = (char)('A' + r) + s;
                i = (i - 1) / 26;
            }
            return s;
        }

        private static string Esc(string t) => System.Security.SecurityElement.Escape(t) ?? "";

        private static void Parte(ZipArchive zip, string nombre, string contenido)
        {
            var entrada = zip.CreateEntry(nombre);
            using var s = entrada.Open();
            var bytes = new UTF8Encoding(false).GetBytes(contenido);
            s.Write(bytes, 0, bytes.Length);
        }
    }

    /// <summary>Base de las clases de prueba: carpeta temporal propia, borrada al terminar.</summary>
    public abstract class PruebaConCarpeta : IDisposable
    {
        protected string Carpeta { get; } =
            Path.Combine(Path.GetTempPath(), "AnalizadorTests", Guid.NewGuid().ToString("N"));

        protected PruebaConCarpeta() => Directory.CreateDirectory(Carpeta);

        public void Dispose()
        {
            try { if (Directory.Exists(Carpeta)) Directory.Delete(Carpeta, true); } catch { }
            GC.SuppressFinalize(this);
        }
    }
}
