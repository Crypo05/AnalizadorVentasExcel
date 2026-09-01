using System.Globalization;

namespace AnalizadorVentasExcel.Modelos
{
    /// <summary>
    /// Fila de la tabla de resultados. Los textos se calculan una sola vez al construir
    /// el resumen: antes se formateaban dentro del getter, así que WPF rehacía el formato
    /// (incluido un clon de NumberFormatInfo) en cada repintado de celda.
    /// </summary>
    public sealed class ResumenDinamico
    {
        internal static readonly NumberFormatInfo FormatoCR = CrearFormato();

        private static NumberFormatInfo CrearFormato()
        {
            var n = (NumberFormatInfo)CultureInfo.GetCultureInfo("es-CR").NumberFormat.Clone();
            n.CurrencySymbol = "₡";
            return NumberFormatInfo.ReadOnly(n);
        }

        public string Etiqueta { get; set; } = string.Empty;
        public string DetalleSecundario { get; set; } = string.Empty;
        public double ValorNumerico { get; set; }
        public double MargenPromedio { get; set; }
        public string Participacion { get; set; } = "-";

        public string ValorFormateado { get; set; } = string.Empty;
        public string MargenFormateado { get; set; } = string.Empty;

        /// <summary>Formatea el valor según la operación y fija los textos de la fila.</summary>
        public void Formatear(string operacion)
        {
            ValorFormateado =
                operacion.Contains("Suma") ? ValorNumerico.ToString("C2", FormatoCR) :
                operacion.Contains("Promedio") ? ValorNumerico.ToString("P2", FormatoCR) :
                ValorNumerico.ToString("N0", FormatoCR);
            MargenFormateado = MargenPromedio.ToString("P2", FormatoCR);
        }
    }
}
