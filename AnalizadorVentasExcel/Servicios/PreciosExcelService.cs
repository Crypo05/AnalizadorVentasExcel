using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using AnalizadorVentasExcel.Modelos;
using ExcelDataReader;
using static AnalizadorVentasExcel.Servicios.CeldasExcel;

namespace AnalizadorVentasExcel.Servicios
{
    /// <summary>Fila del reporte de precios tal como sale del Excel, antes de indexarse.</summary>
    internal readonly struct FilaPrecio
    {
        public readonly string Codigo, Descripcion;
        public readonly decimal Costo, Impuesto, Utilidad, PrecioVenta;

        public FilaPrecio(string codigo, string descripcion, decimal costo,
                          decimal impuesto, decimal utilidad, decimal precioVenta)
        {
            Codigo = codigo; Descripcion = descripcion; Costo = costo;
            Impuesto = impuesto; Utilidad = utilidad; PrecioVenta = precioVenta;
        }
    }

    public sealed class ResultadoCargaPrecios
    {
        public ConjuntoPrecios Datos { get; init; } = ConjuntoPrecios.Vacio;
        public int ArchivosLeidos { get; init; }
        public List<string> Errores { get; init; } = new();

        /// <summary>Archivos que se leyeron bien pero no traían el encabezado de precios.</summary>
        public List<string> Descartados { get; init; } = new();
    }

    /// <summary>
    /// Lectura de los reportes "Comparativa de precios": una lista de precios por sucursal,
    /// con las columnas Cód. Artículo, Descripción, Precio costo, Imp. ventas,
    /// Porc. utilidad y Precio IVI. Igual que el lector de ventas, cada archivo es una
    /// sucursal y los archivos se procesan en paralelo.
    /// </summary>
    public sealed class PreciosExcelService
    {
        static PreciosExcelService() => Preparar();

        public async Task<ResultadoCargaPrecios> CargarCarpetaAsync(
            IReadOnlyList<string> archivos,
            IProgress<string>? progreso = null, CancellationToken ct = default)
        {
            if (archivos.Count == 0) return new ResultadoCargaPrecios();

            var errores = new List<string>();
            var descartados = new List<string>();
            var porArchivo = new (string sucursal, List<FilaPrecio> filas)[archivos.Count];
            int completados = 0;

            await Task.Run(() =>
            {
                var opciones = new ParallelOptions
                {
                    MaxDegreeOfParallelism = Math.Min(archivos.Count, Environment.ProcessorCount),
                    CancellationToken = ct
                };

                Parallel.For(0, archivos.Count, opciones, i =>
                {
                    string ruta = archivos[i];
                    string sucursal = NombreSucursal(ruta);
                    try
                    {
                        var filas = LeerArchivo(ruta);
                        porArchivo[i] = (sucursal, filas);
                        if (filas.Count == 0)
                            lock (descartados) descartados.Add(Path.GetFileName(ruta));
                    }
                    catch (Exception ex)
                    {
                        porArchivo[i] = (sucursal, new List<FilaPrecio>());
                        lock (errores) errores.Add($"{Path.GetFileName(ruta)}: {ex.Message}");
                    }
                    finally
                    {
                        int hechos = Interlocked.Increment(ref completados);
                        progreso?.Report($"Procesando {hechos}/{archivos.Count} archivos...");
                    }
                });
            }, ct).ConfigureAwait(false);

            ct.ThrowIfCancellationRequested();
            progreso?.Report("Indexando precios...");

            var datos = await Task.Run(() => Consolidar(porArchivo), ct).ConfigureAwait(false);

            return new ResultadoCargaPrecios
            {
                Datos = datos,
                ArchivosLeidos = datos.Sucursales.Count,
                Errores = errores,
                Descartados = descartados
            };
        }

        /// <summary>
        /// Los archivos suelen llamarse "precios la bomba 07-09-26": la fecha y la palabra
        /// "precios" son ruido en la cabecera de la columna, así que se recortan.
        /// </summary>
        internal static string NombreSucursal(string ruta)
        {
            string nombre = Path.GetFileNameWithoutExtension(ruta).Trim();

            var partes = nombre.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                               .Where(p => !EsFecha(p))
                               .ToList();
            if (partes.Count > 1 && Normalizar(partes[0]).StartsWith("precio", StringComparison.Ordinal))
                partes.RemoveAt(0);

            return partes.Count == 0 ? nombre : string.Join(" ", partes);
        }

        /// <summary>Trozos como "07-09-26" o "2026-09-07": sólo dígitos y separadores.</summary>
        private static bool EsFecha(string parte)
        {
            bool digito = false;
            foreach (char c in parte)
            {
                if (char.IsDigit(c)) { digito = true; continue; }
                if (c == '-' || c == '/' || c == '.' || c == '_') continue;
                return false;
            }
            return digito;
        }

        /// <summary>
        /// Une lo leído de cada archivo asignando ids de catálogo, en una única pasada
        /// secuencial (los catálogos no son thread-safe). Un archivo sin filas no genera
        /// sucursal: no habría nada que comparar en su columna.
        /// </summary>
        private static ConjuntoPrecios Consolidar((string sucursal, List<FilaPrecio> filas)[] porArchivo)
        {
            int total = 0;
            foreach (var (_, filas) in porArchivo) total += filas?.Count ?? 0;

            var sucursales = new Catalogo(8);
            var codigos = new Catalogo(Math.Max(64, total / 2));
            var descripciones = new Catalogo(Math.Max(64, total / 2));

            var salida = new PrecioItem[total];
            int n = 0;

            for (int a = 0; a < porArchivo.Length; a++)
            {
                var (sucursal, filas) = porArchivo[a];
                porArchivo[a] = default;   // se suelta el buffer crudo en cuanto se consume

                if (filas == null || filas.Count == 0) continue;
                int sucId = sucursales.Id(sucursal);

                foreach (var f in filas)
                    salida[n++] = new PrecioItem(sucId, codigos.Id(f.Codigo), descripciones.Id(f.Descripcion),
                                                 f.Costo, f.Impuesto, f.Utilidad, f.PrecioVenta);
            }

            if (n != total) Array.Resize(ref salida, n);
            return new ConjuntoPrecios(salida, sucursales, codigos, descripciones);
        }

        // ==========================================
        // Lectura de un archivo
        // ==========================================
        private static List<FilaPrecio> LeerArchivo(string ruta)
        {
            var lista = new List<FilaPrecio>(4096);

            using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite,
                                              bufferSize: 1 << 16, FileOptions.SequentialScan);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            if (!reader.Read()) return lista;

            int colCodigo = -1, colDesc = -1, colCosto = -1, colImp = -1, colUtil = -1, colPrecio = -1;
            bool encabezadoEncontrado = false;
            int filasInspeccionadas = 0;

            // El Read() de arriba ya posicionó la primera fila: hay que evaluarla también.
            do
            {
                if (EsFilaVacia(reader)) continue;
                if (++filasInspeccionadas > 20) break;

                DetectarEncabezado(reader, ref colCodigo, ref colDesc, ref colCosto,
                                   ref colImp, ref colUtil, ref colPrecio);

                // Sin código no hay forma de cruzar el producto entre sucursales, y sin
                // ninguna columna numérica no hay nada que comparar.
                if (colCodigo != -1 && (colPrecio != -1 || colCosto != -1 || colUtil != -1))
                {
                    encabezadoEncontrado = true;
                    break;
                }
            } while (reader.Read());

            if (!encabezadoEncontrado) return lista;

            // Deduplicación local: los mismos textos se repiten miles de veces por archivo.
            var pool = new Dictionary<string, string>(1024, StringComparer.Ordinal);
            var vistos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            while (reader.Read())
            {
                if (EsFilaVacia(reader)) continue;

                string codigo = Texto(reader, colCodigo).Trim();
                if (codigo.Length == 0) continue;

                // Un total o subtotal al pie del reporte no es un producto.
                if (Normalizar(codigo).Contains("total", StringComparison.Ordinal)) continue;

                // El mismo código dos veces en una sucursal dejaría la comparación ambigua;
                // se conserva la primera aparición, que es la que ve el usuario en el Excel.
                if (!vistos.Add(codigo)) continue;

                string descripcion = colDesc != -1 ? Texto(reader, colDesc).Trim() : string.Empty;
                if (descripcion.Length == 0) descripcion = "Sin descripción";

                LeerDecimal(reader, colCosto, out decimal costo);
                LeerDecimal(reader, colImp, out decimal impuesto);
                LeerDecimal(reader, colUtil, out decimal utilidad);
                LeerDecimal(reader, colPrecio, out decimal precio);

                // Una fila sin ningún importe es un separador del reporte, no un producto.
                if (costo == 0m && precio == 0m && utilidad == 0m) continue;

                lista.Add(new FilaPrecio(Interno(pool, codigo), Interno(pool, descripcion),
                                         costo, impuesto, utilidad, precio));
            }

            return lista;
        }

        /// <summary>
        /// El orden de las comprobaciones importa: "Precio IVI - artículo" contiene tanto
        /// "precio" como "articulo", y "Precio costo" también contiene "precio".
        /// </summary>
        private static void DetectarEncabezado(IExcelDataReader reader,
            ref int colCodigo, ref int colDesc, ref int colCosto,
            ref int colImp, ref int colUtil, ref int colPrecio)
        {
            // Se reinician en cada fila candidata: el encabezado real debe traer todo junto.
            colCodigo = -1; colDesc = -1; colCosto = -1; colImp = -1; colUtil = -1; colPrecio = -1;

            for (int c = 0; c < reader.FieldCount; c++)
            {
                string val = Normalizar(Texto(reader, c));
                if (val.Length == 0) continue;

                if (val.Contains("ivi", StringComparison.Ordinal) ||
                    val.Contains("precio venta", StringComparison.Ordinal) ||
                    val.Contains("precio de venta", StringComparison.Ordinal)) colPrecio = c;
                else if (val.Contains("costo", StringComparison.Ordinal)) colCosto = c;
                else if (val.Contains("utilidad", StringComparison.Ordinal)) colUtil = c;
                else if (val.Contains("imp", StringComparison.Ordinal)) colImp = c;
                else if (val.Contains("descrip", StringComparison.Ordinal) ||
                         val.Contains("nombre", StringComparison.Ordinal)) colDesc = c;
                else if (val.Contains("cod", StringComparison.Ordinal) ||
                         val.Contains("barra", StringComparison.Ordinal) ||
                         val == "articulo") colCodigo = c;
            }
        }

        private static string Interno(Dictionary<string, string> pool, string s)
        {
            if (s.Length == 0) return string.Empty;
            if (pool.TryGetValue(s, out string? existente)) return existente;
            pool[s] = s;
            return s;
        }
    }
}
