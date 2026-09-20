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

        /// <summary>Un veredicto por archivo de la carpeta, cargado o no, en el orden de la carpeta.</summary>
        public List<DiagnosticoArchivo> Diagnosticos { get; init; } = new();

        public bool HayErrores => Diagnosticos.Exists(d => d.Estado == EstadoCarga.NoCargado);
        public bool HayAvisos => Diagnosticos.Exists(d => d.Estado == EstadoCarga.ConAvisos);
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

            var porArchivo = new (string sucursal, List<FilaPrecio> filas)[archivos.Count];
            var diagnosticos = new DiagnosticoArchivo[archivos.Count];
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
                    var diag = new DiagnosticoArchivo(ruta, sucursal);
                    try
                    {
                        porArchivo[i] = (sucursal, LeerArchivo(ruta, diag));
                    }
                    catch (Exception ex)
                    {
                        // La excepción cruda no le dice nada al usuario: el clasificador la
                        // traduce a qué pasó y qué hacer.
                        porArchivo[i] = (sucursal, new List<FilaPrecio>());
                        diag = DiagnosticoCarga.Clasificar(ex, ruta, sucursal);
                    }
                    finally
                    {
                        diagnosticos[i] = diag;
                        int hechos = Interlocked.Increment(ref completados);
                        progreso?.Report($"Procesando {hechos}/{archivos.Count} archivos...");
                    }
                });
            }, ct).ConfigureAwait(false);

            ct.ThrowIfCancellationRequested();
            progreso?.Report("Indexando precios...");

            var datos = await Task.Run(() => Consolidar(porArchivo), ct).ConfigureAwait(false);

            var lista = diagnosticos.ToList();
            DiagnosticoCarga.MarcarDuplicadas(lista);

            return new ResultadoCargaPrecios
            {
                Datos = datos,
                ArchivosLeidos = datos.Sucursales.Count,
                Diagnosticos = lista
            };
        }

        /// <summary>
        /// Palabras con las que la caja bautiza los reportes: "precios la bomba 07-09-26",
        /// "costos y precios catarata 07-09-26". Delante del nombre real de la sucursal son
        /// ruido, y con varias columnas abiertas a la vez hacen ilegible el encabezado.
        /// </summary>
        private static readonly string[] PalabrasDeReporte =
            { "precio", "precios", "costo", "costos", "lista", "listas", "utilidad", "y", "de", "del" };

        /// <summary>
        /// Nombre de la sucursal a partir del archivo: se le quitan la fecha y las palabras
        /// del tipo de reporte, hasta llegar a la primera que no lo sea. El recorte se frena
        /// siempre antes de dejarlo vacío, así que un archivo llamado sólo "precios.xlsx"
        /// conserva ese nombre en vez de quedarse sin ninguno.
        /// </summary>
        internal static string NombreSucursal(string ruta)
        {
            string nombre = Path.GetFileNameWithoutExtension(ruta).Trim();

            var partes = nombre.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                               .Where(p => !EsFecha(p))
                               .ToList();

            while (partes.Count > 1 && Array.IndexOf(PalabrasDeReporte, Normalizar(partes[0])) >= 0)
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

        /// <summary>Posición de cada columna conocida en una fila candidata a encabezado (-1 si no está).</summary>
        private struct Columnas
        {
            public int Codigo, Desc, Costo, Imp, Util, Precio;

            /// <summary>
            /// Sin código no hay forma de cruzar el producto entre sucursales, y sin ninguna
            /// columna numérica no hay nada que comparar.
            /// </summary>
            public bool Completo => Codigo != -1 && (Precio != -1 || Costo != -1 || Util != -1);

            public int Reconocidas =>
                (Codigo != -1 ? 1 : 0) + (Desc != -1 ? 1 : 0) + (Costo != -1 ? 1 : 0) +
                (Imp != -1 ? 1 : 0) + (Util != -1 ? 1 : 0) + (Precio != -1 ? 1 : 0);

            public static Columnas Vacias => new() { Codigo = -1, Desc = -1, Costo = -1, Imp = -1, Util = -1, Precio = -1 };
        }

        /// <summary>
        /// Lee una lista de precios y deja en <paramref name="diag"/> el veredicto: qué
        /// encabezado encontró (o a qué se parecía lo que había), en qué hoja, y cuántas
        /// filas descartó y por qué. Antes, un archivo sin encabezado devolvía la lista
        /// vacía y sólo quedaba el nombre en una lista de "descartados", sin la razón.
        /// </summary>
        private static List<FilaPrecio> LeerArchivo(string ruta, DiagnosticoArchivo diag)
        {
            var lista = new List<FilaPrecio>(4096);

            using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite,
                                              bufferSize: 1 << 16, FileOptions.SequentialScan);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            // --- Encabezado: en la primera hoja y, si no está, en las siguientes ---
            var col = Columnas.Vacias;
            bool encontrado = false;
            int hojaUsada = 0;
            string nombreHoja = string.Empty;

            // Lo que más se pareció a un encabezado, para poder decir qué le faltaba.
            int mejorPuntaje = 0, mejorFila = 0;
            string mejorHoja = string.Empty;
            List<string>? mejorReconocidas = null;
            Columnas mejorCol = Columnas.Vacias;
            bool pareceDeVentas = false;
            bool algunaFila = false;

            int indiceHoja = 0;
            do
            {
                int filasInspeccionadas = 0, numeroFila = 0;
                while (reader.Read())
                {
                    numeroFila++;
                    if (EsFilaVacia(reader)) continue;
                    algunaFila = true;
                    if (++filasInspeccionadas > 20) break;

                    // La firma del otro sistema se mira ANTES de aceptar el encabezado: el
                    // reporte de ventas trae «Artículo» y «% Utilidad», que este detector
                    // tomaría por código y utilidad, y cargaría la tabla dinámica entera como
                    // si fueran 4.000 productos.
                    if (PareceEncabezadoVentas(TextosDeFila(reader))) { pareceDeVentas = true; break; }

                    var candidata = DetectarEncabezado(reader, out var reconocidas);
                    if (candidata.Completo)
                    {
                        col = candidata; encontrado = true;
                        hojaUsada = indiceHoja; nombreHoja = reader.Name ?? string.Empty;
                        break;
                    }

                    if (candidata.Reconocidas > mejorPuntaje)
                    {
                        mejorPuntaje = candidata.Reconocidas; mejorCol = candidata;
                        mejorFila = numeroFila; mejorHoja = reader.Name ?? string.Empty;
                        mejorReconocidas = reconocidas;
                    }
                }
                if (encontrado) break;
                indiceHoja++;
            } while (reader.NextResult());

            if (!encontrado)
            {
                if (pareceDeVentas)
                    diag.Registrar(ProblemaCarga.ReporteDelOtroSistema,
                        "Es un reporte de ventas, no una lista de precios: tiene «Año mes», «Familia» y «Total».",
                        "Cargalo en la ventana principal (📂 Seleccionar Carpeta), no en la comparativa.");
                else if (!algunaFila)
                    diag.Registrar(ProblemaCarga.ArchivoVacio, "El archivo no tiene ninguna fila con datos.");
                else if (mejorReconocidas == null || mejorReconocidas.Count == 0)
                    diag.Registrar(ProblemaCarga.SinEncabezado,
                        "En las primeras 20 filas no hay nada que parezca un encabezado de precios " +
                        "(se buscan «Cód. Artículo», «Descripción», «Precio costo», «Porc. utilidad», «Precio IVI»).");
                else
                {
                    var faltan = new List<string>();
                    if (mejorCol.Codigo == -1) faltan.Add("«Cód. Artículo»");
                    if (mejorCol.Precio == -1 && mejorCol.Costo == -1 && mejorCol.Util == -1)
                        faltan.Add("alguna columna de importe («Precio IVI», «Precio costo» o «Porc. utilidad»)");
                    string donde = indiceHoja > 1 && mejorHoja.Length > 0 ? $" de la hoja «{mejorHoja}»" : string.Empty;
                    diag.Registrar(ProblemaCarga.SinEncabezado,
                        $"No hay un encabezado completo en las primeras 20 filas. La fila {mejorFila}{donde} se parece: " +
                        $"tiene {string.Join(", ", mejorReconocidas.Select(r => $"«{r}»"))}, pero falta {string.Join(" y ", faltan)}.");
                }
                return lista;
            }

            if (hojaUsada > 0)
                diag.Registrar(ProblemaCarga.HojaEquivocada,
                    $"Los datos estaban en la hoja «{nombreHoja}», no en la primera. Se usó esa.");

            // Columnas opcionales: se carga igual, pero la comparativa mostrará ceros ahí.
            if (col.Precio == -1)
                diag.Registrar(ProblemaCarga.EncabezadoIncompleto,
                    "Falta «Precio IVI»: la comparativa por precio de venta va a mostrar ₡0,00 en esta sucursal.");
            if (col.Costo == -1)
                diag.Registrar(ProblemaCarga.EncabezadoIncompleto,
                    "Falta «Precio costo»: la comparativa por costo va a mostrar ₡0,00 en esta sucursal.");
            if (col.Util == -1)
                diag.Registrar(ProblemaCarga.EncabezadoIncompleto,
                    "Falta «Porc. utilidad»: la comparativa por utilidad va a mostrar 0 % en esta sucursal.");
            if (col.Desc == -1)
                diag.Registrar(ProblemaCarga.EncabezadoIncompleto,
                    "Falta «Descripción»: los productos quedan como «Sin descripción».");

            // Deduplicación local: los mismos textos se repiten miles de veces por archivo.
            var pool = new Dictionary<string, string>(1024, StringComparer.Ordinal);
            var vistos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Por qué se descarta cada fila que se descarta.
            int sinCodigo = 0, filasDeTotal = 0, repetidos = 0, sinImporte = 0;

            while (reader.Read())
            {
                if (EsFilaVacia(reader)) continue;

                string codigo = Texto(reader, col.Codigo).Trim();
                if (codigo.Length == 0) { sinCodigo++; continue; }

                // Un total o subtotal al pie del reporte no es un producto.
                if (Normalizar(codigo).Contains("total", StringComparison.Ordinal)) { filasDeTotal++; continue; }

                // El mismo código dos veces en una sucursal dejaría la comparación ambigua;
                // se conserva la primera aparición, que es la que ve el usuario en el Excel.
                if (!vistos.Add(codigo)) { repetidos++; continue; }

                string descripcion = col.Desc != -1 ? Texto(reader, col.Desc).Trim() : string.Empty;
                if (descripcion.Length == 0) descripcion = "Sin descripción";

                LeerDecimal(reader, col.Costo, out decimal costo);
                LeerDecimal(reader, col.Imp, out decimal impuesto);
                LeerDecimal(reader, col.Util, out decimal utilidad);
                LeerDecimal(reader, col.Precio, out decimal precio);

                // Una fila sin ningún importe es un separador del reporte, no un producto.
                if (costo == 0m && precio == 0m && utilidad == 0m) { sinImporte++; continue; }

                lista.Add(new FilaPrecio(Interno(pool, codigo), Interno(pool, descripcion),
                                         costo, impuesto, utilidad, precio));
            }

            diag.FilasLeidas = lista.Count;
            AnotarDescartes(diag, lista.Count, sinCodigo, filasDeTotal, repetidos, sinImporte);
            return lista;
        }

        /// <summary>
        /// Convierte los contadores en una frase. Un código repetido siempre es aviso: pierde
        /// datos en silencio. Los productos sin ningún importe avisan sólo si pasan del 5 %
        /// (una docena en 2.900 es normal y no vale la pena marcar el archivo por eso). Las
        /// filas sin código y los totales son parte del formato y sólo van al detalle.
        /// </summary>
        private static void AnotarDescartes(DiagnosticoArchivo diag, int cargadas,
                                            int sinCodigo, int filasDeTotal, int repetidos, int sinImporte)
        {
            var f = ResumenDinamico.FormatoCR;
            var partes = new List<string>();
            if (sinImporte > 0) partes.Add($"{sinImporte.ToString("N0", f)} sin ningún importe (costo, precio y utilidad en 0)");
            if (repetidos > 0) partes.Add($"{repetidos.ToString("N0", f)} con código repetido (se tomó la primera aparición)");
            if (sinCodigo > 0) partes.Add($"{sinCodigo.ToString("N0", f)} sin código");
            if (filasDeTotal > 0) partes.Add($"{filasDeTotal.ToString("N0", f)} filas de total");

            if (cargadas == 0)
            {
                diag.Registrar(partes.Count == 0 ? ProblemaCarga.ArchivoVacio : ProblemaCarga.SinFilasUtiles,
                    partes.Count == 0
                        ? "El encabezado está, pero debajo no hay ninguna fila con datos."
                        : $"El encabezado está bien pero ninguna fila sirvió: {string.Join(", ", partes)}.");
                return;
            }

            if (partes.Count == 0) return;

            string resumen = $"{cargadas.ToString("N0", f)} productos cargados · {string.Join(" · ", partes)}.";
            if (repetidos > 0 || ExcelService.SuperaUmbral(sinImporte, cargadas))
                diag.Registrar(ProblemaCarga.FilasDescartadas, resumen);
            else diag.Detalles.Add(resumen);
        }

        /// <summary>
        /// Reconoce las columnas de una fila. El orden de las comprobaciones importa:
        /// "Precio IVI - artículo" contiene tanto "precio" como "articulo", y "Precio costo"
        /// también contiene "precio". Devuelve además los textos originales de las columnas
        /// reconocidas, para decirle al usuario qué encontró cuando falta el encabezado.
        /// </summary>
        private static Columnas DetectarEncabezado(IExcelDataReader reader, out List<string> reconocidas)
        {
            var col = Columnas.Vacias;
            reconocidas = new List<string>();

            for (int c = 0; c < reader.FieldCount; c++)
            {
                string original = Texto(reader, c).Trim();
                if (original.Length == 0) continue;
                string val = Normalizar(original);

                bool hit = true;
                if (val.Contains("ivi", StringComparison.Ordinal) ||
                    val.Contains("precio venta", StringComparison.Ordinal) ||
                    val.Contains("precio de venta", StringComparison.Ordinal)) col.Precio = c;
                else if (val.Contains("costo", StringComparison.Ordinal)) col.Costo = c;
                else if (val.Contains("utilidad", StringComparison.Ordinal)) col.Util = c;
                else if (val.Contains("imp", StringComparison.Ordinal)) col.Imp = c;
                else if (val.Contains("descrip", StringComparison.Ordinal) ||
                         val.Contains("nombre", StringComparison.Ordinal)) col.Desc = c;
                else if (val.Contains("cod", StringComparison.Ordinal) ||
                         val.Contains("barra", StringComparison.Ordinal) ||
                         val == "articulo") col.Codigo = c;
                else hit = false;

                if (hit) reconocidas.Add(original);
            }
            return col;
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
