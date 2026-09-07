using System;
using System.Windows.Media;

namespace AnalizadorVentasExcel.Modelos
{
    /// <summary>Qué columna del reporte de precios se compara entre sucursales.</summary>
    public enum MetricaPrecio
    {
        PrecioVenta = 0,
        Costo = 1,
        Utilidad = 2
    }

    /// <summary>
    /// Una línea del reporte "Comparativa de precios": un producto en una sucursal.
    /// Igual que <see cref="VentaItem"/>, los textos ya vienen convertidos a ids de
    /// catálogo para que el conjunto entero viva en un array contiguo.
    /// </summary>
    public readonly struct PrecioItem
    {
        public readonly int SucursalId;
        public readonly int CodigoId;
        public readonly int DescripcionId;
        public readonly decimal Costo;
        public readonly decimal Impuesto;
        public readonly decimal Utilidad;
        public readonly decimal PrecioVenta;

        public PrecioItem(int sucursalId, int codigoId, int descripcionId,
                          decimal costo, decimal impuesto, decimal utilidad, decimal precioVenta)
        {
            SucursalId = sucursalId;
            CodigoId = codigoId;
            DescripcionId = descripcionId;
            Costo = costo;
            Impuesto = impuesto;
            Utilidad = utilidad;
            PrecioVenta = precioVenta;
        }
    }

    /// <summary>Las listas de precios cargadas más los catálogos de cada dimensión.</summary>
    public sealed class ConjuntoPrecios
    {
        public static readonly ConjuntoPrecios Vacio = new ConjuntoPrecios(
            Array.Empty<PrecioItem>(), new Catalogo(), new Catalogo(), new Catalogo());

        public PrecioItem[] Filas { get; }
        public Catalogo Sucursales { get; }
        public Catalogo Codigos { get; }
        public Catalogo Descripciones { get; }

        public int Count => Filas.Length;

        public ConjuntoPrecios(PrecioItem[] filas, Catalogo sucursales, Catalogo codigos, Catalogo descripciones)
        {
            Filas = filas;
            Sucursales = sucursales;
            Codigos = codigos;
            Descripciones = descripciones;
        }

        public static decimal ValorDe(in PrecioItem p, MetricaPrecio m) => m switch
        {
            MetricaPrecio.Costo => p.Costo,
            MetricaPrecio.Utilidad => p.Utilidad,
            _ => p.PrecioVenta
        };

        public static string NombreDe(MetricaPrecio m) => m switch
        {
            MetricaPrecio.Costo => "Precio de costo",
            MetricaPrecio.Utilidad => "% de utilidad",
            _ => "Precio de venta (IVI)"
        };

        /// <summary>
        /// En precio y costo la sucursal más barata es la mejor; en utilidad, la de mayor
        /// margen. De esto depende de qué color se pinta cada celda.
        /// </summary>
        public static bool MejorEsMayor(MetricaPrecio m) => m == MetricaPrecio.Utilidad;
    }

    /// <summary>
    /// Un producto con su valor en cada sucursal comparada. Los textos y los colores se
    /// calculan al construir la fila, no en el getter: la tabla los repinta constantemente.
    /// </summary>
    public sealed class FilaComparativa
    {
        internal static readonly Brush ColorMejor = Congelado(Color.FromRgb(0x1E, 0x8E, 0x3E));
        internal static readonly Brush ColorPeor = Congelado(Color.FromRgb(0xC0, 0x39, 0x2B));
        internal static readonly Brush ColorNormal = Congelado(Color.FromRgb(0x2C, 0x3E, 0x50));
        internal static readonly Brush ColorAusente = Congelado(Color.FromRgb(0xB2, 0xBA, 0xBB));

        private static Brush Congelado(Color c)
        {
            var b = new SolidColorBrush(c);
            b.Freeze();   // compartido entre miles de celdas y varios hilos
            return b;
        }

        public string Codigo { get; init; } = string.Empty;
        public string Descripcion { get; init; } = string.Empty;

        /// <summary>Nombre del producto en cada sucursal cuando no coinciden entre sí.</summary>
        public string DetalleDescripciones { get; init; } = string.Empty;
        public bool DescripcionesDistintas { get; init; }

        /// <summary>Valor en cada sucursal comparada, null si el producto no está en ella.</summary>
        public decimal?[] Valores { get; init; } = Array.Empty<decimal?>();
        public string[] Textos { get; init; } = Array.Empty<string>();
        public Brush[] Colores { get; init; } = Array.Empty<Brush>();

        /// <summary>En cuántas de las sucursales comparadas existe el producto.</summary>
        public int Presencia { get; init; }

        public decimal Minimo { get; init; }
        public decimal Maximo { get; init; }
        public decimal Diferencia { get; init; }
        public decimal DiferenciaPct { get; init; }

        public string DiferenciaTexto { get; init; } = "-";
        public string DiferenciaPctTexto { get; init; } = "-";

        /// <summary>Sucursal con el mejor valor y con el peor, según la métrica.</summary>
        public string Mejor { get; init; } = "-";
        public string Peor { get; init; } = "-";
        public string PresenciaTexto { get; init; } = string.Empty;
        public string Aviso { get; init; } = string.Empty;

        public static string Formatear(decimal valor, MetricaPrecio metrica)
            => metrica == MetricaPrecio.Utilidad
                ? valor.ToString("N2", ResumenDinamico.FormatoCR) + " %"
                : valor.ToString("C2", ResumenDinamico.FormatoCR);

        public static string FormatearPorcentaje(decimal valor)
            => valor.ToString("N1", ResumenDinamico.FormatoCR) + " %";
    }
}
