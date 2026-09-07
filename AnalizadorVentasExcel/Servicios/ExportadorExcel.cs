using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using AnalizadorVentasExcel.Modelos;

namespace AnalizadorVentasExcel.Servicios
{
    /// <summary>
    /// Escribe la tabla de la comparativa en un .xlsx.
    ///
    /// Un .xlsx es un zip con unas pocas partes XML, y lo que hay que exportar es una
    /// cuadrícula plana, así que se arma a mano en vez de traer una librería: el programa
    /// se distribuye como un único ejecutable autocontenido de ~150 MB que cada sucursal
    /// vuelve a descargar entera en cada actualización, y no vale la pena engordarlo para
    /// esto.
    ///
    /// Los importes se escriben como NÚMEROS con formato, no como el texto que se ve en
    /// pantalla: se ven igual en Excel, pero además se pueden ordenar, sumar y filtrar,
    /// que es para lo que uno se lleva los datos a una hoja. Los códigos, en cambio, van
    /// como texto a propósito, porque muchos empiezan con ceros ("00022") y Excel se los
    /// comería si los tomara como número.
    /// </summary>
    public static class ExportadorExcel
    {
        // Índices dentro de <cellXfs> de styles.xml.
        private const int EstiloNormal = 0;
        private const int EstiloEncabezado = 1;
        private const int EstiloMoneda = 2;
        private const int EstiloUtilidad = 3;
        private const int EstiloPorcentaje = 4;
        private const int EstiloTexto = 5;

        public static void ExportarComparativa(string ruta, IReadOnlyList<FilaComparativa> filas,
                                               IReadOnlyList<string> sucursales, MetricaPrecio metrica)
        {
            var encabezados = Encabezados(sucursales, metrica);
            string hoja = ConstruirHoja(filas, sucursales, metrica, encabezados);

            // Se escribe a un temporal y se mueve al final: si algo falla a mitad de camino,
            // el archivo que el usuario ya tenía no queda pisado ni a medias.
            string temporal = Path.Combine(Path.GetDirectoryName(ruta) ?? ".",
                                           Path.GetFileName(ruta) + ".tmp");
            try
            {
                using (var archivo = new FileStream(temporal, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var zip = new ZipArchive(archivo, ZipArchiveMode.Create))
                {
                    Escribir(zip, "[Content_Types].xml", ContentTypes);
                    Escribir(zip, "_rels/.rels", Rels);
                    Escribir(zip, "xl/workbook.xml", Workbook);
                    Escribir(zip, "xl/_rels/workbook.xml.rels", WorkbookRels);
                    Escribir(zip, "xl/styles.xml", Estilos);
                    Escribir(zip, "xl/worksheets/sheet1.xml", hoja);
                }

                File.Move(temporal, ruta, overwrite: true);
            }
            finally
            {
                if (File.Exists(temporal))
                    try { File.Delete(temporal); } catch { /* el temporal quedó suelto: no es crítico */ }
            }
        }

        /// <summary>Las mismas columnas de la tabla, incluidos los nombres que cambian con la métrica.</summary>
        private static List<string> Encabezados(IReadOnlyList<string> sucursales, MetricaPrecio metrica)
        {
            var lista = new List<string>(sucursales.Count * 2 + 8) { "Código", "Descripción" };

            if (ConjuntoPrecios.EsCombinada(metrica))
            {
                // Dos columnas por sucursal, y las de "Más barata"/"Más cara" se omiten
                // igual que en la tabla: con cuatro columnas de diferencia sobran.
                foreach (string s in sucursales) { lista.Add($"{s} costo"); lista.Add($"{s} venta"); }
                lista.Add("Dif. costo");
                lista.Add("Dif. costo %");
                lista.Add("Dif. venta");
                lista.Add("Dif. venta %");
            }
            else
            {
                bool utilidad = metrica == MetricaPrecio.Utilidad;
                lista.AddRange(sucursales);
                lista.Add(utilidad ? "Dif. (puntos)" : "Dif.");
                lista.Add("Dif. %");
                lista.Add(utilidad ? "Mayor utilidad" : "Más barata");
                lista.Add(utilidad ? "Menor utilidad" : "Más cara");
            }

            lista.Add("Suc.");
            lista.Add("Aviso");
            return lista;
        }

        private static string ConstruirHoja(IReadOnlyList<FilaComparativa> filas,
                                            IReadOnlyList<string> sucursales, MetricaPrecio metrica,
                                            List<string> encabezados)
        {
            bool combinada = ConjuntoPrecios.EsCombinada(metrica);
            int estiloValor = metrica == MetricaPrecio.Utilidad ? EstiloUtilidad : EstiloMoneda;
            int porSucursal = combinada ? 2 : 1;
            int columnas = encabezados.Count;
            int totalFilas = filas.Count + 1;

            var sb = new StringBuilder(filas.Count * 320 + 2048);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            sb.Append("<dimension ref=\"A1:").Append(Columna(columnas - 1)).Append(totalFilas).Append("\"/>");

            // La fila de encabezados queda fija al desplazarse, que con miles de productos
            // es la diferencia entre poder leer la hoja y no.
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\">");
            sb.Append("<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>");
            sb.Append("</sheetView></sheetViews>");
            sb.Append("<sheetFormatPr defaultRowHeight=\"15\"/>");

            sb.Append("<cols>");
            sb.Append("<col min=\"1\" max=\"1\" width=\"18\" customWidth=\"1\"/>");
            sb.Append("<col min=\"2\" max=\"2\" width=\"42\" customWidth=\"1\"/>");
            sb.Append("<col min=\"3\" max=\"").Append(columnas).Append("\" width=\"16\" customWidth=\"1\"/>");
            sb.Append("</cols>");

            sb.Append("<sheetData>");

            sb.Append("<row r=\"1\">");
            for (int c = 0; c < encabezados.Count; c++)
                CeldaTexto(sb, c, 1, encabezados[c], EstiloEncabezado);
            sb.Append("</row>");

            for (int i = 0; i < filas.Count; i++)
            {
                var f = filas[i];
                int fila = i + 2;
                int c = 0;

                sb.Append("<row r=\"").Append(fila).Append("\">");

                CeldaTexto(sb, c++, fila, f.Codigo, EstiloTexto);
                CeldaTexto(sb, c++, fila, f.Descripcion, EstiloNormal);

                for (int v = 0; v < sucursales.Count * porSucursal; v++, c++)
                {
                    // Sin valor no se escribe la celda: en la tabla se ve "—" y acá queda
                    // vacía, que es lo que Excel entiende como "este producto no está".
                    decimal? valor = v < f.Valores.Length ? f.Valores[v] : null;
                    if (valor.HasValue) CeldaNumero(sb, c, fila, valor.Value, estiloValor);
                }

                bool comparable = f.Presencia >= 2;
                if (comparable) CeldaNumero(sb, c, fila, f.Diferencia, estiloValor);
                c++;

                if (comparable && f.Minimo > 0m) CeldaNumero(sb, c, fila, f.DiferenciaPct, EstiloPorcentaje);
                c++;

                if (combinada)
                {
                    // La segunda diferencia sólo existe en la vista combinada, y se escribe
                    // con el mismo criterio: en blanco cuando no hay con qué comparar.
                    if (comparable) CeldaNumero(sb, c, fila, f.DiferenciaVenta, estiloValor);
                    c++;
                    if (comparable && f.DiferenciaVentaPctTexto != "-")
                        CeldaNumero(sb, c, fila, f.DiferenciaVentaPct, EstiloPorcentaje);
                    c++;
                }
                else
                {
                    CeldaTexto(sb, c++, fila, f.Mejor, EstiloNormal);
                    CeldaTexto(sb, c++, fila, f.Peor, EstiloNormal);
                }

                CeldaTexto(sb, c++, fila, f.PresenciaTexto, EstiloTexto);
                CeldaTexto(sb, c, fila, f.Aviso, EstiloNormal);

                sb.Append("</row>");
            }

            sb.Append("</sheetData>");
            sb.Append("<autoFilter ref=\"A1:").Append(Columna(columnas - 1)).Append(totalFilas).Append("\"/>");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        private static void CeldaTexto(StringBuilder sb, int columna, int fila, string valor, int estilo)
        {
            if (string.IsNullOrEmpty(valor)) return;
            sb.Append("<c r=\"").Append(Columna(columna)).Append(fila)
              .Append("\" s=\"").Append(estilo).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
            Escapar(sb, valor);
            sb.Append("</t></is></c>");
        }

        private static void CeldaNumero(StringBuilder sb, int columna, int fila, decimal valor, int estilo)
        {
            sb.Append("<c r=\"").Append(Columna(columna)).Append(fila)
              .Append("\" s=\"").Append(estilo).Append("\"><v>")
              .Append(valor.ToString(CultureInfo.InvariantCulture))
              .Append("</v></c>");
        }

        /// <summary>Índice 0 -> "A", 25 -> "Z", 26 -> "AA".</summary>
        private static string Columna(int indice)
        {
            Span<char> buffer = stackalloc char[4];
            int n = 0;
            for (int i = indice; ; i = i / 26 - 1)
            {
                buffer[n++] = (char)('A' + i % 26);
                if (i < 26) break;
            }

            Span<char> resultado = stackalloc char[n];
            for (int i = 0; i < n; i++) resultado[i] = buffer[n - 1 - i];
            return new string(resultado);
        }

        /// <summary>
        /// Escapa el texto para XML. Además descarta los caracteres de control, que son
        /// válidos en una celda de Excel pero rompen el XML: si un nombre de producto trae
        /// basura del sistema de la caja, el archivo entero dejaría de abrir.
        /// </summary>
        private static void Escapar(StringBuilder sb, string texto)
        {
            foreach (char ch in texto)
            {
                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    default:
                        if (ch >= ' ' || ch == '\t' || ch == '\n' || ch == '\r') sb.Append(ch);
                        break;
                }
            }
        }

        private static void Escribir(ZipArchive zip, string nombre, string contenido)
        {
            var entrada = zip.CreateEntry(nombre, CompressionLevel.Optimal);
            using var flujo = entrada.Open();
            // Sin BOM: Excel lo acepta, pero algunos lectores tropiezan con él.
            using var escritor = new StreamWriter(flujo, new UTF8Encoding(false));
            escritor.Write(contenido);
        }

        // ==========================================
        // Partes fijas del libro
        // ==========================================
        private const string ContentTypes =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
            "</Types>";

        private const string Rels =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>";

        private const string Workbook =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<sheets><sheet name=\"Comparativa\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
            "</workbook>";

        private const string WorkbookRels =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
            "</Relationships>";

        /// <summary>
        /// Los dos primeros rellenos (ninguno y gray125) son obligatorios y en ese orden:
        /// Excel da el archivo por corrupto si faltan.
        /// </summary>
        private const string Estilos =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<numFmts count=\"3\">" +
            "<numFmt numFmtId=\"164\" formatCode=\"&quot;₡&quot;#,##0.00\"/>" +
            "<numFmt numFmtId=\"165\" formatCode=\"0.00&quot; %&quot;\"/>" +
            "<numFmt numFmtId=\"166\" formatCode=\"0.0&quot; %&quot;\"/>" +
            "</numFmts>" +
            "<fonts count=\"2\">" +
            "<font><sz val=\"11\"/><color rgb=\"FF000000\"/><name val=\"Calibri\"/></font>" +
            "<font><b/><sz val=\"11\"/><color rgb=\"FFFFFFFF\"/><name val=\"Calibri\"/></font>" +
            "</fonts>" +
            "<fills count=\"3\">" +
            "<fill><patternFill patternType=\"none\"/></fill>" +
            "<fill><patternFill patternType=\"gray125\"/></fill>" +
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF8E44AD\"/><bgColor indexed=\"64\"/></patternFill></fill>" +
            "</fills>" +
            "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"6\">" +
            "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
            "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>" +
            "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
            "<xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
            "<xf numFmtId=\"166\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
            "<xf numFmtId=\"49\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
            "</cellXfs>" +
            "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
            "</styleSheet>";
    }
}
