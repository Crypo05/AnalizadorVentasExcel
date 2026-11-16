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

    /// <summary>
    /// Qué columnas de importe trae la lista de una sucursal. Una lista sin «Porc. utilidad»
    /// (o con el costo con impuesto, que se ignora) no tiene ese dato: la comparativa lo
    /// muestra vacío en esa sucursal en vez de un 0 que la haría parecer la más barata.
    /// </summary>
    [Flags]
    public enum ColumnasLista
    {
        Ninguna = 0,
        Costo = 1,
        Utilidad = 2,
        PrecioVenta = 4,
        Todas = Costo | Utilidad | PrecioVenta
    }

    /// <summary>Las listas de precios cargadas más los catálogos de cada dimensión.</summary>
    public sealed class ConjuntoPrecios
    {
        public static readonly ConjuntoPrecios Vacio = new ConjuntoPrecios(
            Array.Empty<PrecioItem>(), new Catalogo(), new Catalogo(), new Catalogo());

        /// <summary>
        /// Prefijo de la clave de catálogo de un producto que vino de una lista sin código y
        /// no se pudo cruzar con ningún código. Ninguna celda de Excel puede traer el carácter
        /// nulo (el XML no lo admite), así que no choca con un código real.
        /// </summary>
        internal const char MarcaSinCodigo = '\0';

        public PrecioItem[] Filas { get; }
        public Catalogo Sucursales { get; }
        public Catalogo Codigos { get; }
        public Catalogo Descripciones { get; }

        /// <summary>Columnas de importe de cada sucursal, indexadas por id de sucursal.</summary>
        public ColumnasLista[] ColumnasPorSucursal { get; }

        /// <summary>Listas que no traían código y se cruzaron por descripción, con sus conteos.</summary>
        public IReadOnlyList<CrucePorDescripcion> CrucesPorDescripcion { get; }

        public bool HayCrucePorDescripcion => CrucesPorDescripcion.Count > 0;

        public int Count => Filas.Length;

        public ConjuntoPrecios(PrecioItem[] filas, Catalogo sucursales, Catalogo codigos, Catalogo descripciones,
                               ColumnasLista[]? columnasPorSucursal = null,
                               IReadOnlyList<CrucePorDescripcion>? crucesPorDescripcion = null)
        {
            Filas = filas;
            Sucursales = sucursales;
            Codigos = codigos;
            Descripciones = descripciones;
            ColumnasPorSucursal = columnasPorSucursal ?? Array.Empty<ColumnasLista>();
            CrucesPorDescripcion = crucesPorDescripcion ?? Array.Empty<CrucePorDescripcion>();
        }

        /// <summary>
        /// Si la lista de la sucursal trae el dato de la métrica. Sin información de columnas
        /// (conjuntos armados a mano) se asume que sí, que es como se comportaba antes.
        /// </summary>
        public bool Tiene(int sucursalId, MetricaPrecio m)
        {
            if ((uint)sucursalId >= (uint)ColumnasPorSucursal.Length) return true;
            var columnas = ColumnasPorSucursal[sucursalId];
            return m switch
            {
                MetricaPrecio.Costo => (columnas & ColumnasLista.Costo) != 0,
                MetricaPrecio.Utilidad => (columnas & ColumnasLista.Utilidad) != 0,
                _ => (columnas & ColumnasLista.PrecioVenta) != 0
            };
        }

        /// <summary>Lo que se muestra en la columna «Código»: los productos sin código no muestran la clave interna.</summary>
        public static string TextoCodigo(string clave)
            => clave.Length > 0 && clave[0] == MarcaSinCodigo ? "(sin código)" : clave;

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
    /// Cómo quedó el cruce de una lista que no traía «Cód. Artículo». Esa lista se cruza
    /// con las demás por la descripción (igualdad tras quitar tildes, mayúsculas y espacios
    /// de más; nada aproximado), y estos conteos son los que se le muestran al usuario.
    /// </summary>
    public sealed class CrucePorDescripcion
    {
        public string Sucursal { get; init; } = string.Empty;
        public string Archivo { get; init; } = string.Empty;

        /// <summary>Productos de la lista que entraron a la comparativa (sin los repetidos).</summary>
        public int Productos { get; init; }

        /// <summary>Productos que quedaron en la misma fila que el de alguna otra sucursal.</summary>
        public int Cruzados { get; init; }

        public int SinCruzar => Productos - Cruzados;

        /// <summary>
        /// De los que quedaron sin cruzar, los que no se cruzaron porque el nombre no
        /// identifica un único código: en las otras listas ese nombre tiene más de un
        /// código, o dos productos de esta lista apuntan al mismo código.
        /// </summary>
        public int NombresAmbiguos { get; init; }

        /// <summary>Filas que quedaron fuera porque su descripción se repite dentro de la misma lista.</summary>
        public int RepetidosFilas { get; init; }
        public int RepetidosNombres { get; init; }
        public IReadOnlyList<string> EjemplosRepetidos { get; init; } = Array.Empty<string>();
    }

    /// <summary>
    /// Textos del aviso de cruce por descripción. Una sola fuente para la ventana de la
    /// comparativa, el archivo exportado y el diagnóstico de carga, para que digan lo mismo.
    /// </summary>
    public static class AvisoCruce
    {
        public const string Riesgo =
            "Puede haber incongruencias: dos productos distintos con el mismo nombre quedan como uno solo, " +
            "y un mismo producto que en otra sucursal tiene otro nombre no se cruza.";

        /// <summary>El aviso completo, una línea por renglón. Vacío si no hay cruces por descripción.</summary>
        public static IReadOnlyList<string> Lineas(IReadOnlyList<CrucePorDescripcion> cruces)
        {
            if (cruces.Count == 0) return Array.Empty<string>();

            var lineas = new List<string>(cruces.Count + 2);
            string nombres = Enumerar(cruces.Select(c => c.Sucursal).ToList());
            lineas.Add(cruces.Count == 1
                ? $"Atención: la lista de {nombres} no trae «Cód. Artículo», así que se cruzó con las demás por descripción y no por código."
                : $"Atención: las listas de {nombres} no traen «Cód. Artículo», así que se cruzaron con las demás por descripción y no por código.");
            lineas.Add(Riesgo);
            foreach (var c in cruces) lineas.Add(c.Sucursal + ": " + Conteos(c));
            return lineas;
        }

        public static string Texto(IReadOnlyList<CrucePorDescripcion> cruces) => string.Join("\n", Lineas(cruces));

        /// <summary>Mensaje principal del diagnóstico de la lista cruzada por descripción.</summary>
        public static string ParaDiagnostico(CrucePorDescripcion c)
            => "No trae «Cód. Artículo»: se cruzó con las demás listas por descripción y no por código. " +
               Riesgo + " " + Conteos(c);

        /// <summary>"2.772 de 2.774 productos se cruzaron con otra sucursal; 2 quedaron sin cruzar; ..."</summary>
        public static string Conteos(CrucePorDescripcion c)
        {
            var f = ResumenDinamico.FormatoCR;
            var sb = new System.Text.StringBuilder();
            sb.Append(c.Cruzados.ToString("N0", f)).Append(" de ").Append(c.Productos.ToString("N0", f))
              .Append(c.Productos == 1 ? " producto se cruzó" : " productos se cruzaron").Append(" con otra sucursal; ")
              .Append(c.SinCruzar.ToString("N0", f)).Append(" quedaron sin cruzar");
            if (c.NombresAmbiguos > 0)
                sb.Append(" (").Append(c.NombresAmbiguos.ToString("N0", f))
                  .Append(" de ellos porque su nombre no identifica un único código)");
            if (c.RepetidosFilas > 0)
            {
                sb.Append("; ").Append(c.RepetidosFilas.ToString("N0", f))
                  .Append(" filas quedaron fuera porque su descripción se repite en la lista (")
                  .Append(c.RepetidosNombres.ToString("N0", f))
                  .Append(c.RepetidosNombres == 1 ? " descripción" : " descripciones");
                if (c.EjemplosRepetidos.Count > 0)
                    sb.Append(", por ejemplo ").Append(string.Join(", ", c.EjemplosRepetidos));
                sb.Append(')');
            }
            sb.Append('.');
            return sb.ToString();
        }

        /// <summary>"a", "a y b", "a, b y c".</summary>
        internal static string Enumerar(IReadOnlyList<string> partes) => partes.Count switch
        {
            0 => string.Empty,
            1 => partes[0],
            _ => string.Join(", ", partes.Take(partes.Count - 1)) + " y " + partes[^1]
        };
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

        /// <summary>Texto de la celda cuando el producto está en la sucursal pero su lista no trae esa columna.</summary>
        public const string TextoSinDato = "sin dato";

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

        /// <summary>
        /// Si hay al menos dos sucursales con dato para la métrica principal (o para el precio
        /// de venta, en la vista combinada). Sin eso la diferencia no existe: se muestra "-" y
        /// no se exporta un 0. Coincide con <c>Presencia &gt;= 2</c> cuando todas las listas
        /// traen todas las columnas.
        /// </summary>
        public bool ComparablePrincipal { get; init; }
        public bool ComparableVenta { get; init; }

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
