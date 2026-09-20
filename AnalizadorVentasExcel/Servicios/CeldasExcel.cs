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
        /// Firma del reporte de precios (Cód. Artículo + Precio IVI o Precio costo). Se usa
        /// desde el lector de VENTAS para avisar que el archivo es del otro sistema, que es
        /// el error más común: los dos leen "cualquier Excel de la carpeta".
        /// </summary>
        internal static bool PareceEncabezadoPrecios(IReadOnlyList<string> celdas)
        {
            bool codigo = false, precio = false;
            foreach (string celda in celdas)
            {
                string n = Normalizar(celda);
                if (n.Contains("cod", StringComparison.Ordinal) && n.Contains("art", StringComparison.Ordinal)) codigo = true;
                if (n.Contains("ivi", StringComparison.Ordinal) || n.Contains("precio costo", StringComparison.Ordinal)) precio = true;
            }
            return codigo && precio;
        }

        /// <summary>Firma del reporte de ventas (Año mes + Total + Familia), para el lector de precios.</summary>
        internal static bool PareceEncabezadoVentas(IReadOnlyList<string> celdas)
        {
            bool fecha = false, total = false, familia = false;
            foreach (string celda in celdas)
            {
                string n = Normalizar(celda);
                if (n.Contains("ano", StringComparison.Ordinal) || n == "mes") fecha = true;
                if (n == "total" || n == "total venta") total = true;
                if (n.Contains("familia", StringComparison.Ordinal)) familia = true;
            }
            return total && familia && fecha;
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
