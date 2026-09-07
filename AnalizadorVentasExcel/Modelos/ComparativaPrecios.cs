using System;
using System.Windows.Media;

namespace AnalizadorVentasExcel.Modelos
{
    /// <summary>
    /// Qué columna del reporte de precios se compara entre sucursales.
    /// <see cref="CostoYVenta"/> es distinta a las demás: muestra dos valores por sucursal
    /// en vez de uno, para ver de un vistazo a cuánto se compra y a cuánto se vende.
    /// </summary>
    public enum MetricaPrecio
    {
        PrecioVenta = 0,
        Costo = 1,
        Utilidad = 2,
        CostoYVenta = 3
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
            MetricaPrecio.CostoYVenta => "Costo y precio de venta",
            _ => "Precio de venta (IVI)"
        };

        /// <summary>
        /// En precio y costo la sucursal más barata es la mejor; en utilidad, la de mayor
        /// margen. De esto depende de qué color se pinta cada celda.
        /// </summary>
        public static bool MejorEsMayor(MetricaPrecio m) => m == MetricaPrecio.Utilidad;

        /// <summary>
        /// La vista que trae costo y venta juntos: dos columnas por sucursal en vez de una,
        /// y dos juegos de columnas de diferencia.
        /// </summary>
        public static bool EsCombinada(MetricaPrecio m) => m == MetricaPrecio.CostoYVenta;

        /// <summary>Las dos métricas que se muestran juntas, en el orden de las columnas.</summary>
        public static readonly MetricaPrecio[] MetricasCombinadas =
            { MetricaPrecio.Costo, MetricaPrecio.PrecioVenta };
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

        /// <summary>
        /// Valor en cada sucursal comparada, null si el producto no está en ella.
        ///
        /// En la vista combinada hay DOS entradas por sucursal, intercaladas: el costo en
        /// las posiciones pares y el precio de venta en las impares (sucursal s ocupa 2s y
        /// 2s+1). Se guardan así, y no en dos arreglos, para que la tabla siga enlazando
        /// por posición (<c>Textos[i]</c>, <c>Colores[i]</c>) sin distinguir el modo.
        /// </summary>
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

        /// <summary>
        /// Segunda diferencia, sólo en la vista combinada: las de arriba son del costo y
        /// estas del precio de venta. Fuera de esa vista quedan en cero.
        /// </summary>
        public decimal DiferenciaVenta { get; init; }
        public decimal DiferenciaVentaPct { get; init; }
        public string DiferenciaVentaTexto { get; init; } = "-";
        public string DiferenciaVentaPctTexto { get; init; } = "-";

        /// <summary>
        /// La mayor de las dos diferencias porcentuales. Es la que ordena y filtra la vista
        /// combinada: lo que interesa es que el producto se separe entre sucursales, sin
        /// importar si se separa al comprarlo o al venderlo.
        /// </summary>
        public decimal DiferenciaMayorPct { get; init; }
        public decimal DiferenciaMayorAbs { get; init; }

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
