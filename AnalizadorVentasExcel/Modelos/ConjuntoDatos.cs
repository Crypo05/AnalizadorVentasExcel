using System.Collections.Generic;

namespace AnalizadorVentasExcel.Modelos
{
    /// <summary>
    /// Diccionario de una dimensión: convierte texto repetido en un identificador entero.
    /// Con esto las 250.000+ filas guardan un int por dimensión en lugar de una referencia
    /// a string distinta, y los filtros y agrupaciones trabajan sobre enteros.
    /// </summary>
    public sealed class Catalogo
    {
        private readonly Dictionary<string, int> _indice;
        private readonly List<string> _valores;

        public Catalogo(int capacidad = 64)
        {
            _indice = new Dictionary<string, int>(capacidad, System.StringComparer.Ordinal);
            _valores = new List<string>(capacidad);
        }

        public int Count => _valores.Count;

        public string this[int id] => (uint)id < (uint)_valores.Count ? _valores[id] : string.Empty;

        public int Id(string? valor)
        {
            valor ??= string.Empty;
            if (_indice.TryGetValue(valor, out int id)) return id;
            id = _valores.Count;
            _valores.Add(valor);
            _indice[valor] = id;
            return id;
        }

        /// <summary>Id de un valor ya conocido, o -1 si no existe. No inserta.</summary>
        public int IdExistente(string? valor)
            => valor != null && _indice.TryGetValue(valor, out int id) ? id : -1;

        public IReadOnlyList<string> Valores => _valores;
    }

    /// <summary>
    /// Una fila de venta. Es un struct de 6 enteros + 2 decimales para que el conjunto
    /// completo viva en un único array contiguo (mejor localidad de caché, sin cabecera
    /// de objeto ni presión sobre el GC por cada fila).
    /// </summary>
    public readonly struct VentaItem
    {
        public readonly int SucursalId;
        public readonly int PeriodoId;
        public readonly int CodigoId;
        public readonly int ArticuloId;
        public readonly int ProveedorId;
        public readonly int FamiliaId;
        public readonly decimal TotalVenta;
        public readonly decimal PorcentajeUtilidad;

        public VentaItem(int sucursalId, int periodoId, int codigoId, int articuloId,
                         int proveedorId, int familiaId, decimal totalVenta, decimal porcentajeUtilidad)
        {
            SucursalId = sucursalId;
            PeriodoId = periodoId;
            CodigoId = codigoId;
            ArticuloId = articuloId;
            ProveedorId = proveedorId;
            FamiliaId = familiaId;
            TotalVenta = totalVenta;
            PorcentajeUtilidad = porcentajeUtilidad;
        }
    }

    public enum Dimension
    {
        AnioMes = 0,
        Proveedor = 1,
        Familia = 2,
        Sucursal = 3,
        Articulo = 4
    }

    /// <summary>
    /// Todo el conjunto cargado: las filas más los catálogos de cada dimensión.
    /// </summary>
    public sealed class ConjuntoDatos
    {
        public static readonly ConjuntoDatos Vacio = new ConjuntoDatos(
            System.Array.Empty<VentaItem>(), new Catalogo(), new Catalogo(), new Catalogo(),
            new Catalogo(), new Catalogo(), new Catalogo(), new Catalogo(), System.Array.Empty<int>());

        public VentaItem[] Filas { get; }
        public Catalogo Sucursales { get; }
        public Catalogo Periodos { get; }
        public Catalogo Proveedores { get; }
        public Catalogo Familias { get; }
        public Catalogo Articulos { get; }
        public Catalogo Codigos { get; }

        /// <summary>Nombres de artículo recortados (Trim), usados por el explorador.</summary>
        public Catalogo ArticulosNormalizados { get; }

        /// <summary>Mapa ArticuloId -> id dentro de <see cref="ArticulosNormalizados"/>.</summary>
        public int[] ArticuloANormalizado { get; }

        public int Count => Filas.Length;

        public ConjuntoDatos(VentaItem[] filas, Catalogo sucursales, Catalogo periodos,
            Catalogo proveedores, Catalogo familias, Catalogo articulos, Catalogo codigos,
            Catalogo articulosNormalizados, int[] articuloANormalizado)
        {
            Filas = filas;
            Sucursales = sucursales;
            Periodos = periodos;
            Proveedores = proveedores;
            Familias = familias;
            Articulos = articulos;
            Codigos = codigos;
            ArticulosNormalizados = articulosNormalizados;
            ArticuloANormalizado = articuloANormalizado;
        }

        public Catalogo CatalogoDe(Dimension d) => d switch
        {
            Dimension.AnioMes => Periodos,
            Dimension.Proveedor => Proveedores,
            Dimension.Familia => Familias,
            Dimension.Sucursal => Sucursales,
            _ => Articulos
        };

        /// <summary>Id de la dimensión <paramref name="d"/> para la fila indicada.</summary>
        public static int IdDe(in VentaItem f, Dimension d) => d switch
        {
            Dimension.AnioMes => f.PeriodoId,
            Dimension.Proveedor => f.ProveedorId,
            Dimension.Familia => f.FamiliaId,
            Dimension.Sucursal => f.SucursalId,
            _ => f.ArticuloId
        };

        public static Dimension? DesdeTexto(string? texto) => texto switch
        {
            "Año Mes" => Dimension.AnioMes,
            "Proveedor" => Dimension.Proveedor,
            "Familia" => Dimension.Familia,
            "Sucursal" => Dimension.Sucursal,
            "Articulo" => Dimension.Articulo,
            _ => null
        };
    }
}
