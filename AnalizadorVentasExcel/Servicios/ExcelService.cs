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

        /// <summary>Un veredicto por archivo de la carpeta, cargado o no, en el orden de la carpeta.</summary>
        public List<DiagnosticoArchivo> Diagnosticos { get; init; } = new();

        public bool HayErrores => Diagnosticos.Exists(d => d.Estado == EstadoCarga.NoCargado);
        public bool HayAvisos => Diagnosticos.Exists(d => d.Estado == EstadoCarga.ConAvisos);
    }

    /// <summary>
    /// Lectura de los libros de ventas.
    /// Usa ExcelDataReader (lectura secuencial en streaming) en vez de cargar el libro
    /// completo en memoria, y procesa los archivos en paralelo.
    /// </summary>
    public sealed class ExcelService
    {
        static ExcelService() => Preparar();

        public async Task<ResultadoCarga> CargarCarpetaAsync(
            IReadOnlyList<string> archivos, string? modo,
            IProgress<string>? progreso = null, CancellationToken ct = default)
        {
            if (archivos.Count == 0) return new ResultadoCarga();

            var porArchivo = new (string sucursal, List<FilaCruda> filas)[archivos.Count];
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
                    string sucursal = Path.GetFileNameWithoutExtension(ruta);
                    var diag = new DiagnosticoArchivo(ruta, sucursal);
                    try
                    {
                        var filas = LeerArchivo(ruta, modo, diag);
                        porArchivo[i] = (sucursal, filas);
                    }
                    catch (Exception ex)
                    {
                        // La excepción cruda ("Invalid file signature.") no le dice nada al
                        // usuario: el clasificador la traduce a qué pasó y qué hacer.
                        porArchivo[i] = (sucursal, new List<FilaCruda>());
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
            progreso?.Report("Indexando datos...");

            var datos = await Task.Run(() => Consolidar(porArchivo), ct).ConfigureAwait(false);

            var lista = diagnosticos.ToList();
            DiagnosticoCarga.MarcarDuplicadas(lista);

            return new ResultadoCarga
            {
                Datos = datos,
                ArchivosLeidos = lista.Count(d => d.Cargado),
                Diagnosticos = lista
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
        /// <summary>Posición de cada columna conocida en una fila candidata a encabezado (-1 si no está).</summary>
        private struct Columnas
        {
            public int Fecha, Codigo, Desc, Prov, Fam, Total, Util;

            /// <summary>Las dos obligatorias: sin total no hay venta, sin familia no hay agrupación.</summary>
            public bool Completo => Total != -1 && Fam != -1;

            public int Reconocidas =>
                (Fecha != -1 ? 1 : 0) + (Codigo != -1 ? 1 : 0) + (Desc != -1 ? 1 : 0) + (Prov != -1 ? 1 : 0) +
                (Fam != -1 ? 1 : 0) + (Total != -1 ? 1 : 0) + (Util != -1 ? 1 : 0);

            public static Columnas Vacias => new() { Fecha = -1, Codigo = -1, Desc = -1, Prov = -1, Fam = -1, Total = -1, Util = -1 };
        }

        /// <summary>
        /// Lee un libro y deja en <paramref name="diag"/> el veredicto: qué encabezado
        /// encontró (o a qué se parecía lo que había), en qué hoja, y cuántas filas descartó
        /// y por qué. Antes, un archivo sin encabezado devolvía la lista vacía y nada más.
        /// </summary>
        private static List<FilaCruda> LeerArchivo(string ruta, string? modo, DiagnosticoArchivo diag)
        {
            var lista = new List<FilaCruda>(4096);

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
            bool pareceDePrecios = false;
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

                    // La firma del otro sistema se mira ANTES de aceptar el encabezado, por
                    // simetría con el lector de precios: ninguna lista de precios trae «Año
                    // mes» ni «Familia», así que la comprobación es segura.
                    if (PareceEncabezadoPrecios(TextosDeFila(reader))) { pareceDePrecios = true; break; }

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
                if (pareceDePrecios)
                    diag.Registrar(ProblemaCarga.ReporteDelOtroSistema,
                        "Es un reporte de precios, no de ventas: tiene «Cód. Artículo» y «Precio IVI».",
                        "Cargalo en ⚖️ Comparativa de Precios, no en esta ventana.");
                else if (!algunaFila)
                    diag.Registrar(ProblemaCarga.ArchivoVacio, "El archivo no tiene ninguna fila con datos.");
                else if (mejorReconocidas == null || mejorReconocidas.Count == 0)
                    diag.Registrar(ProblemaCarga.SinEncabezado,
                        "En las primeras 20 filas no hay nada que parezca un encabezado de ventas " +
                        "(se buscan «Año mes», «Artículo», «Familia», «Total», «% Utilidad»).");
                else
                {
                    var faltan = new List<string>();
                    if (mejorCol.Total == -1) faltan.Add("«Total»");
                    if (mejorCol.Fam == -1) faltan.Add("«Familia»");
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

            bool esMinimarket = col.Codigo != -1;
            if (modo != null && modo.Contains("Minimarket")) esMinimarket = true;
            else if (modo != null && modo.Contains("Souvenir")) esMinimarket = false;

            // Columnas opcionales: se carga igual, pero conviene saber qué se pierde.
            if (col.Fecha == -1)
                diag.Registrar(ProblemaCarga.EncabezadoIncompleto,
                    "Falta la columna «Año mes»: sin periodo no se puede cargar ninguna fila.");
            if (col.Util == -1)
                diag.Registrar(ProblemaCarga.EncabezadoIncompleto, "Falta «% Utilidad»: todo se carga con utilidad 0.");
            if (col.Prov == -1)
                diag.Registrar(ProblemaCarga.EncabezadoIncompleto, "Falta «Proveedor»: todo queda como «General».");
            if (esMinimarket && col.Desc == -1)
                diag.Registrar(ProblemaCarga.EncabezadoIncompleto,
                    "Falta «Artículo desc.»: los productos quedan como «Sin Nombre».");

            string ultPeriodo = string.Empty, ultProv = "General", ultFam = "General";

            // Deduplicación local: el mismo texto aparece miles de veces por archivo.
            var pool = new Dictionary<string, string>(1024, StringComparer.Ordinal);

            // Por qué se descarta cada fila que se descarta. Las filas sin clave y sin importe
            // son la estructura de la tabla dinámica (títulos de grupo) y no se cuentan.
            int sinClave = 0, sinImporte = 0, sinPeriodo = 0, totalesAnuales = 0;

            while (reader.Read())
            {
                if (EsFilaVacia(reader)) continue;

                string s;
                if (col.Fecha != -1 && (s = Texto(reader, col.Fecha)).Length != 0)
                    ultPeriodo = Interno(pool, s);

                if (col.Prov != -1 && (s = Texto(reader, col.Prov)).Length != 0)
                {
                    if (s.IndexOf("total", StringComparison.OrdinalIgnoreCase) < 0)
                        ultProv = Interno(pool, s);
                }

                if (col.Fam != -1 && (s = Texto(reader, col.Fam)).Length != 0)
                    ultFam = Interno(pool, s);

                bool tieneImporte = LeerDecimal(reader, col.Total, out decimal total) && total != 0m;
                bool tieneClave = esMinimarket
                    ? col.Codigo == -1 || Texto(reader, col.Codigo).Length != 0
                    : col.Fam == -1 || Texto(reader, col.Fam).Length != 0;

                if (!tieneClave) { if (tieneImporte) sinClave++; continue; }
                if (!tieneImporte) { sinImporte++; continue; }

                decimal utilidad = 0m;
                if (col.Util != -1) LeerDecimal(reader, col.Util, out utilidad);

                string nombreReal = "Sin Nombre";
                if (esMinimarket)
                {
                    if (col.Desc != -1 && (s = Texto(reader, col.Desc)).Length != 0)
                        nombreReal = Interno(pool, s);
                }
                else nombreReal = ultFam;

                if (ultPeriodo.Length == 0) { sinPeriodo++; continue; }
                if (ultPeriodo.IndexOf("año", StringComparison.OrdinalIgnoreCase) >= 0) { totalesAnuales++; continue; }

                string codigo = col.Codigo != -1 ? Interno(pool, Texto(reader, col.Codigo)) : string.Empty;
                lista.Add(new FilaCruda(ultPeriodo, codigo, nombreReal, ultProv, ultFam, total, utilidad));
            }

            diag.FilasLeidas = lista.Count;
            AnotarDescartes(diag, lista.Count, esMinimarket, sinClave, sinImporte, sinPeriodo, totalesAnuales);
            return lista;
        }

        /// <summary>
        /// Convierte los contadores en una frase. Los subtotales y totales anuales son parte
        /// del formato y sólo se mencionan en el detalle. Las filas que parecían datos pero
        /// se descartaron (sin importe, sin periodo) marcan el archivo con aviso únicamente
        /// si pasan del 5 %: en una tabla dinámica de 80.000 filas siempre hay unas decenas
        /// de productos con venta cero, y avisar por eso en cada carga enseñaría al usuario
        /// a ignorar el aviso.
        /// </summary>
        private static void AnotarDescartes(DiagnosticoArchivo diag, int cargadas, bool esMinimarket,
                                            int sinClave, int sinImporte, int sinPeriodo, int totalesAnuales)
        {
            var f = ResumenDinamico.FormatoCR;
            var partes = new List<string>();
            if (sinImporte > 0) partes.Add($"{sinImporte.ToString("N0", f)} sin importe (Total en 0)");
            if (sinPeriodo > 0) partes.Add($"{sinPeriodo.ToString("N0", f)} sin periodo");
            if (sinClave > 0) partes.Add($"{sinClave.ToString("N0", f)} {(esMinimarket ? "sin código" : "sin familia")} (subtotales)");
            if (totalesAnuales > 0) partes.Add($"{totalesAnuales.ToString("N0", f)} totales anuales");

            if (cargadas == 0)
            {
                diag.Registrar(partes.Count == 0 ? ProblemaCarga.ArchivoVacio : ProblemaCarga.SinFilasUtiles,
                    partes.Count == 0
                        ? "El encabezado está, pero debajo no hay ninguna fila con datos."
                        : $"El encabezado está bien pero ninguna fila sirvió: {string.Join(", ", partes)}.");
                return;
            }

            if (partes.Count == 0) return;

            string resumen = $"{cargadas.ToString("N0", f)} filas cargadas · {string.Join(" · ", partes)}.";
            if (SuperaUmbral(sinImporte + sinPeriodo, cargadas)) diag.Registrar(ProblemaCarga.FilasDescartadas, resumen);
            else diag.Detalles.Add(resumen);
        }

        /// <summary>Más del 5 % de las filas con pinta de datos se descartó.</summary>
        internal static bool SuperaUmbral(int descartadas, int cargadas)
            => descartadas > 0 && descartadas * 20L > (long)(descartadas + cargadas);

        /// <summary>
        /// Reconoce las columnas de una fila. Devuelve también los textos originales de las
        /// que reconoció, para poder decirle al usuario "tiene «Artículo» y «Total» pero
        /// falta «Familia»" cuando ninguna fila llega a encabezado completo.
        /// </summary>
        private static Columnas DetectarEncabezado(IExcelDataReader reader, out List<string> reconocidas)
        {
            var col = Columnas.Vacias;
            reconocidas = new List<string>();

            for (int c = 0; c < reader.FieldCount; c++)
            {
                string original = Texto(reader, c).Trim();
                if (original.Length == 0) continue;
                // Sin tildes ni mayúsculas, igual que el lector de precios: "AÑO MES" y
                // "Ano mes" tienen que reconocerse igual que "Año mes".
                string val = Normalizar(original);

                bool hit = true;
                if (val.Contains("ano", StringComparison.Ordinal) || val == "mes") col.Fecha = c;
                else if (val == "articulo") col.Codigo = c;
                else if (val.Contains("desc") || val.Contains("nombre")) col.Desc = c;
                else if (val.Contains("proveedor")) col.Prov = c;
                else if (val.Contains("familia")) col.Fam = c;
                else if (val == "total" || val == "total venta") col.Total = c;
                else if (val.Contains("utilidad") || val.Contains("%")) col.Util = c;
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
