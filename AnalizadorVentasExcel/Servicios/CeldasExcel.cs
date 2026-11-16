using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using ExcelDataReader;

namespace AnalizadorVentasExcel.Servicios
{
    /// <summary>
    /// Lectura de celdas compartida por los dos lectores del programa (ventas y precios).
    /// Vive aparte porque ambos leen los mismos libros de la misma caja registradora:
    /// mismas rarezas de formato, mismos números escritos como texto.
    /// </summary>
    internal static class CeldasExcel
    {
        /// <summary>Registra las páginas de códigos no-Unicode que necesitan los .xls antiguos.</summary>
        internal static void Preparar()
            => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        internal static bool EsFilaVacia(IExcelDataReader reader)
        {
            for (int c = 0; c < reader.FieldCount; c++)
                if (!reader.IsDBNull(c)) return false;
            return true;
        }

        /// <summary>Valor de la celda como texto, equivalente al GetString() anterior.</summary>
        internal static string Texto(IExcelDataReader reader, int col)
        {
            if (col < 0 || col >= reader.FieldCount || reader.IsDBNull(col)) return string.Empty;
            object v = reader.GetValue(col);
            return v switch
            {
                null => string.Empty,
                string s => s,
                double d => d.ToString(CultureInfo.InvariantCulture),
                DateTime dt => dt.ToString("yyyy-MM", CultureInfo.InvariantCulture),
                bool b => b ? "True" : "False",
                _ => Convert.ToString(v, CultureInfo.InvariantCulture) ?? string.Empty
            };
        }

        internal static bool LeerDecimal(IExcelDataReader reader, int col, out decimal valor)
        {
            valor = 0m;
            if (col < 0 || col >= reader.FieldCount || reader.IsDBNull(col)) return false;

            object v = reader.GetValue(col);
            switch (v)
            {
                case double d:
                    if (double.IsNaN(d) || double.IsInfinity(d)) return false;
                    try { valor = (decimal)d; } catch (OverflowException) { return false; }
                    return true;
                case decimal m:
                    valor = m; return true;
                case int i:
                    valor = i; return true;
                case string s:
                    return ParseDecimalFlexible(s, out valor);
                default:
                    return false;
            }
        }

        internal static bool ParseDecimalFlexible(string t, out decimal r)
        {
            r = 0m;
            if (string.IsNullOrWhiteSpace(t)) return false;
            string l = t.Replace("%", "").Replace("$", "").Replace("₡", "").Trim();
            if (decimal.TryParse(l, NumberStyles.Any, CultureInfo.InvariantCulture, out r)) return true;
            return decimal.TryParse(l, NumberStyles.Any, CultureInfo.GetCultureInfo("es-CR"), out r);
        }

        /// <summary>
        /// Textos normalizados de todas las celdas con contenido de la fila actual. Es lo que
        /// miran las firmas de formato y lo que se muestra al usuario cuando el encabezado
        /// no aparece ("la fila 3 se parece: tiene «Artículo», «Total»...").
        /// </summary>
        internal static List<string> TextosDeFila(IExcelDataReader reader)
        {
            var lista = new List<string>();
            for (int c = 0; c < reader.FieldCount; c++)
            {
                string t = Texto(reader, c).Trim();
                if (t.Length > 0) lista.Add(t);
            }
            return lista;
        }

        /// <summary>
        /// Textos de la fila actual por posición de columna (vacío donde no hay nada). A
        /// diferencia de <see cref="TextosDeFila"/>, conserva el índice: es lo que necesitan
        /// los detectores de encabezado para saber en qué columna está cada dato.
        /// </summary>
        internal static string[] CeldasDeFila(IExcelDataReader reader)
        {
            var celdas = new string[reader.FieldCount];
            for (int c = 0; c < celdas.Length; c++) celdas[c] = Texto(reader, c).Trim();
            return celdas;
        }

        /// <summary>
        /// Firma de una lista de precios, para que el lector de VENTAS avise que el archivo es
        /// del otro sistema, que es el error más común: los dos leen "cualquier Excel de la
        /// carpeta". Basta una columna que identifique el producto (código o descripción) y
        /// otra de importe (precio, costo o IVI), en celdas distintas. No cuenta si la fila
        /// trae «Año mes» o «Total»: eso es de ventas, aunque le falte otra columna.
        ///
        /// Antes pedía «Cód. Artículo» más «Precio IVI» o «Precio costo», y una lista exportada
        /// con otras columnas («Descripción artículo», «Precio venta I.V.A.», «Costo I.V.A.»)
        /// se tomaba por un reporte de ventas incompleto al que le "faltaba Total".
        /// </summary>
        internal static bool PareceEncabezadoPrecios(IReadOnlyList<string> celdas)
            => PareceEncabezadoPrecios(celdas, out _);

        internal static bool PareceEncabezadoPrecios(IReadOnlyList<string> celdas, out List<string> encontradas)
        {
            encontradas = new List<string>();
            bool identificador = false, importe = false;
            foreach (string celda in celdas)
            {
                string n = Normalizar(celda);
                if (n.Length == 0) continue;
                if (EsColumnaDeVentas(n)) { encontradas.Clear(); return false; }

                bool id = n.Contains("cod", StringComparison.Ordinal) || n.Contains("descrip", StringComparison.Ordinal);
                bool monto = n.Contains("ivi", StringComparison.Ordinal) || n.Contains("precio", StringComparison.Ordinal) ||
                             n.Contains("costo", StringComparison.Ordinal);
                if (!id && !monto) continue;

                encontradas.Add(celda.Trim());
                identificador |= id;
                importe |= monto;
            }

            bool parece = identificador && importe && encontradas.Count >= 2;
            if (!parece) encontradas.Clear();
            return parece;
        }

        /// <summary>«Año mes» o «Total»: columnas que sólo trae el reporte de ventas.</summary>
        private static bool EsColumnaDeVentas(string normalizado)
            => normalizado == "total" || normalizado == "total venta" || normalizado == "mes" || normalizado == "ano" ||
               (normalizado.Contains("ano", StringComparison.Ordinal) && normalizado.Contains("mes", StringComparison.Ordinal));

        /// <summary>Firma del reporte de ventas (Año mes + Total + Familia), para el lector de precios.</summary>
        internal static bool PareceEncabezadoVentas(IReadOnlyList<string> celdas)
            => PareceEncabezadoVentas(celdas, out _);

        internal static bool PareceEncabezadoVentas(IReadOnlyList<string> celdas, out List<string> encontradas)
        {
            encontradas = new List<string>();
            bool fecha = false, total = false, familia = false;
            foreach (string celda in celdas)
            {
                string n = Normalizar(celda);
                bool esFecha = n.Contains("ano", StringComparison.Ordinal) || n == "mes";
                bool esTotal = n == "total" || n == "total venta";
                bool esFamilia = n.Contains("familia", StringComparison.Ordinal);
                if (esFecha || esTotal || esFamilia) encontradas.Add(celda.Trim());
                fecha |= esFecha; total |= esTotal; familia |= esFamilia;
            }
            bool parece = total && familia && fecha;
            if (!parece) encontradas.Clear();
            return parece;
        }

        /// <summary>
        /// Si el encabezado dice que el importe trae el impuesto incluido: «Costo I.V.A.»,
        /// «Costo IVA», «Costo IVI». Se mira por palabras, sin los puntos de la sigla.
        /// </summary>
        internal static bool TieneMarcaDeImpuesto(string normalizado)
        {
            string sinPuntos = normalizado.Replace(".", string.Empty);
            foreach (string palabra in sinPuntos.Split(SeparadoresDePalabra, StringSplitOptions.RemoveEmptyEntries))
                if (palabra == "iva" || palabra == "ivi") return true;
            return false;
        }

        private static readonly char[] SeparadoresDePalabra = { ' ', '-', '_', '/', '(', ')', '+', ':', ',', ';' };

        /// <summary>
        /// Forma de comparar descripciones entre listas: sin tildes, sin mayúsculas y con los
        /// espacios de más colapsados. Es igualdad exacta después de eso; nada aproximado.
        /// </summary>
        internal static string Compactar(string texto)
        {
            string n = Normalizar(texto);
            var sb = new StringBuilder(n.Length);
            bool espacio = false;
            foreach (char c in n)
            {
                if (char.IsWhiteSpace(c)) { espacio = sb.Length > 0; continue; }
                if (espacio) { sb.Append(' '); espacio = false; }
                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Minúsculas y sin tildes. Los encabezados vienen escritos de forma inconsistente
        /// entre reportes ("Cód. Artículo", "Cod. articulo"), así que se comparan normalizados.
        /// </summary>
        internal static string Normalizar(string texto)
        {
            if (texto.Length == 0) return string.Empty;

            string descompuesto = texto.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(descompuesto.Length);
            foreach (char c in descompuesto)
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) sb.Append(c);
            return sb.ToString();
        }
    }
}
