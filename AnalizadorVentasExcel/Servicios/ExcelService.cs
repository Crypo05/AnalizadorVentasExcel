using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using AnalizadorVentasExcel.Modelos;
using ExcelDataReader;

namespace AnalizadorVentasExcel.Servicios
{
    /// <summary>Fila cruda tal como sale del Excel, antes de convertir textos a ids.</summary>
    internal readonly struct FilaCruda
    {
        public readonly string Periodo, Codigo, Articulo, Proveedor, Familia;
        public readonly decimal Total, Utilidad;

        public FilaCruda(string periodo, string codigo, string articulo, string proveedor,
                         string familia, decimal total, decimal utilidad)
        {
            Periodo = periodo; Codigo = codigo; Articulo = articulo;
            Proveedor = proveedor; Familia = familia; Total = total; Utilidad = utilidad;
        }
    }

    public sealed class ResultadoCarga
    {
        public ConjuntoDatos Datos { get; init; } = ConjuntoDatos.Vacio;
        public int ArchivosLeidos { get; init; }
        public List<string> Errores { get; init; } = new();
    }

    /// <summary>
    /// Lectura de los libros de ventas.
    /// Usa ExcelDataReader (lectura secuencial en streaming) en vez de cargar el libro
    /// completo en memoria, y procesa los archivos en paralelo.
    /// </summary>
    public sealed class ExcelService
    {
        static ExcelService()
        {
            // Necesario para los .xls antiguos (páginas de códigos no-Unicode).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public async Task<ResultadoCarga> CargarCarpetaAsync(
            IReadOnlyList<string> archivos, string? modo,
            IProgress<string>? progreso = null, CancellationToken ct = default)
        {
            if (archivos.Count == 0) return new ResultadoCarga();

            var errores = new List<string>();
            var porArchivo = new (string sucursal, List<FilaCruda> filas)[archivos.Count];
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
                    string sucursal = Path.GetFileNameWithoutExtension(ruta);
                    try
                    {
                        porArchivo[i] = (sucursal, LeerArchivo(ruta, modo));
                    }
                    catch (Exception ex)
                    {
                        porArchivo[i] = (sucursal, new List<FilaCruda>());
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
            progreso?.Report("Indexando datos...");

            var datos = await Task.Run(() => Consolidar(porArchivo), ct).ConfigureAwait(false);

            return new ResultadoCarga
            {
                Datos = datos,
                ArchivosLeidos = porArchivo.Count(p => p.filas is { Count: > 0 }),
                Errores = errores
            };
        }

        /// <summary>
        /// Une lo leído de cada archivo asignando ids de catálogo. Es una sola pasada
        /// secuencial: los catálogos no son thread-safe, y aquí es donde se elimina la
        /// duplicación de cadenas (cientos de miles de strings repetidos -> unos cientos).
        /// </summary>
        private static ConjuntoDatos Consolidar((string sucursal, List<FilaCruda> filas)[] porArchivo)
        {
            int total = 0;
            foreach (var (_, filas) in porArchivo) total += filas?.Count ?? 0;

            var sucursales = new Catalogo(8);
            var periodos = new Catalogo(64);
            var proveedores = new Catalogo(256);
            var familias = new Catalogo(128);
            var articulos = new Catalogo(Math.Max(64, total / 16));
            var codigos = new Catalogo(Math.Max(64, total / 16));
            var articulosNorm = new Catalogo(Math.Max(64, total / 16));
            var articuloANorm = new List<int>(Math.Max(64, total / 16));

            var salida = new VentaItem[total];
            int n = 0;

            for (int a = 0; a < porArchivo.Length; a++)
            {
                var (sucursal, filas) = porArchivo[a];
                // Se suelta la referencia al buffer crudo en cuanto se consume, para no
                // mantener vivas a la vez la representación cruda y la definitiva.
                porArchivo[a] = default;

                if (filas == null || filas.Count == 0) continue;
                int sucId = sucursales.Id(sucursal);

                foreach (var f in filas)
                {
                    int artId = articulos.Id(f.Articulo);
                    // El catálogo de artículos crece de a uno; mantenemos el mapa alineado.
                    while (articuloANorm.Count <= artId)
                        articuloANorm.Add(articulosNorm.Id(articulos[articuloANorm.Count].Trim()));

                    salida[n++] = new VentaItem(
                        sucId,
                        periodos.Id(f.Periodo),
                        codigos.Id(f.Codigo),
                        artId,
                        proveedores.Id(f.Proveedor),
                        familias.Id(f.Familia),
                        f.Total,
                        f.Utilidad);
                }
            }

            if (n != total) Array.Resize(ref salida, n);

            return new ConjuntoDatos(salida, sucursales, periodos, proveedores, familias,
                                     articulos, codigos, articulosNorm, articuloANorm.ToArray());
        }

        // ==========================================
        // Lectura de un archivo
        // ==========================================
        private static List<FilaCruda> LeerArchivo(string ruta, string? modo)
        {
            var lista = new List<FilaCruda>(4096);

            using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite,
                                              bufferSize: 1 << 16, FileOptions.SequentialScan);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            // Sólo la primera hoja, igual que antes.
            if (!reader.Read()) return lista;

            int colFecha = -1, colCodigo = -1, colDesc = -1, colProv = -1, colFam = -1, colTotal = -1, colUtil = -1;
            bool encabezadoEncontrado = false;
            int filasInspeccionadas = 0;

            // El Read() de arriba ya posicionó la primera fila: hay que evaluarla también.
            do
            {
                if (EsFilaVacia(reader)) continue;
                if (++filasInspeccionadas > 20) break;

                DetectarEncabezado(reader, ref colFecha, ref colCodigo, ref colDesc,
                                   ref colProv, ref colFam, ref colTotal, ref colUtil);

                if (colTotal != -1 && colFam != -1) { encabezadoEncontrado = true; break; }
            } while (reader.Read());

            if (!encabezadoEncontrado) return lista;

            bool esMinimarket = colCodigo != -1;
            if (modo != null && modo.Contains("Minimarket")) esMinimarket = true;
            else if (modo != null && modo.Contains("Souvenir")) esMinimarket = false;

            string ultPeriodo = string.Empty, ultProv = "General", ultFam = "General";

            // Deduplicación local: el mismo texto aparece miles de veces por archivo.
            var pool = new Dictionary<string, string>(1024, StringComparer.Ordinal);

            while (reader.Read())
            {
                if (EsFilaVacia(reader)) continue;

                string s;
                if (colFecha != -1 && (s = Texto(reader, colFecha)).Length != 0)
                    ultPeriodo = Interno(pool, s);

                if (colProv != -1 && (s = Texto(reader, colProv)).Length != 0)
                {
                    if (s.IndexOf("total", StringComparison.OrdinalIgnoreCase) < 0)
                        ultProv = Interno(pool, s);
                }

                if (colFam != -1 && (s = Texto(reader, colFam)).Length != 0)
                    ultFam = Interno(pool, s);

                if (esMinimarket && colCodigo != -1 && Texto(reader, colCodigo).Length == 0) continue;
                if (!esMinimarket && colFam != -1 && Texto(reader, colFam).Length == 0) continue;

                if (!LeerDecimal(reader, colTotal, out decimal total) || total == 0m) continue;

                decimal utilidad = 0m;
                if (colUtil != -1) LeerDecimal(reader, colUtil, out utilidad);

                string nombreReal = "Sin Nombre";
                if (esMinimarket)
                {
                    if (colDesc != -1 && (s = Texto(reader, colDesc)).Length != 0)
                        nombreReal = Interno(pool, s);
                }
                else nombreReal = ultFam;

                if (ultPeriodo.Length == 0) continue;
                if (ultPeriodo.IndexOf("año", StringComparison.OrdinalIgnoreCase) >= 0) continue;

                string codigo = colCodigo != -1 ? Interno(pool, Texto(reader, colCodigo)) : string.Empty;
                lista.Add(new FilaCruda(ultPeriodo, codigo, nombreReal, ultProv, ultFam, total, utilidad));
            }

            return lista;
        }

        private static void DetectarEncabezado(IExcelDataReader reader,
            ref int colFecha, ref int colCodigo, ref int colDesc,
            ref int colProv, ref int colFam, ref int colTotal, ref int colUtil)
        {
            // Se reinician en cada fila candidata: el encabezado real debe traer todo junto.
            colFecha = -1; colCodigo = -1; colDesc = -1; colProv = -1; colFam = -1; colTotal = -1; colUtil = -1;

            for (int c = 0; c < reader.FieldCount; c++)
            {
                string val = Texto(reader, c);
                if (val.Length == 0) continue;
                val = val.ToLowerInvariant().Trim();

                if (val.Contains("año") || val == "mes") colFecha = c;
                else if (val == "artículo" || val == "articulo") colCodigo = c;
                else if (val.Contains("desc") || val.Contains("nombre")) colDesc = c;
                else if (val.Contains("proveedor")) colProv = c;
                else if (val.Contains("familia")) colFam = c;
                else if (val == "total" || val == "total venta") colTotal = c;
                else if (val.Contains("utilidad") || val.Contains("%")) colUtil = c;
            }
        }

        private static bool EsFilaVacia(IExcelDataReader reader)
        {
            for (int c = 0; c < reader.FieldCount; c++)
                if (!reader.IsDBNull(c)) return false;
            return true;
        }

        /// <summary>Valor de la celda como texto, equivalente al GetString() anterior.</summary>
        private static string Texto(IExcelDataReader reader, int col)
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

        private static bool LeerDecimal(IExcelDataReader reader, int col, out decimal valor)
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

        private static bool ParseDecimalFlexible(string t, out decimal r)
        {
            r = 0m;
            if (string.IsNullOrWhiteSpace(t)) return false;
            string l = t.Replace("%", "").Replace("$", "").Replace("₡", "").Trim();
            if (decimal.TryParse(l, NumberStyles.Any, CultureInfo.InvariantCulture, out r)) return true;
            return decimal.TryParse(l, NumberStyles.Any, CultureInfo.GetCultureInfo("es-CR"), out r);
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
