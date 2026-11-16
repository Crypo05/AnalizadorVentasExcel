using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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

    /// <summary>Lo leído de un archivo: sus filas y lo que hace falta saber para consolidarlo.</summary>
    internal sealed class ListaLeida
    {
        public List<FilaPrecio> Filas { get; } = new(4096);

        /// <summary>La lista no trae «Cód. Artículo»: se cruza con las demás por descripción.</summary>
        public bool PorDescripcion { get; set; }

        /// <summary>Columnas de importe que trae la lista; las que faltan se muestran vacías, no en 0.</summary>
        public ColumnasLista Columnas { get; set; } = ColumnasLista.Todas;

        /// <summary>Filas excluidas por descripción repetida (sólo en listas sin código).</summary>
        public int RepetidosFilas { get; set; }
        public int RepetidosNombres { get; set; }
        public List<string> EjemplosRepetidos { get; } = new();
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
    ///
    /// Una lista sin «Cód. Artículo» se acepta igual y se cruza por descripción, sólo como
    /// respaldo y con aviso: las listas con código se siguen cruzando por código entre sí
    /// exactamente igual que antes.
    /// </summary>
    public sealed class PreciosExcelService
    {
        static PreciosExcelService() => Preparar();

        /// <summary>Texto que se pone a un producto sin descripción en una lista con código.</summary>
        private const string SinDescripcion = "Sin descripción";

        /// <summary>Lo mínimo para poder cargar una lista, dicho como se le dice al usuario.</summary>
        internal const string RequisitoMinimo =
            "hace falta «Cód. Artículo» o «Descripción», y al menos una de «Precio costo», «Porc. utilidad» o «Precio IVI»";

        internal const string SugerenciaListaEstandar =
            "Volvé a exportar la lista desde la caja con las columnas de la lista estándar: Cód. Artículo, " +
            "Descripción, Precio costo, Imp. ventas, Porc. utilidad y Precio IVI.";

        public async Task<ResultadoCargaPrecios> CargarCarpetaAsync(
            IReadOnlyList<string> archivos,
            IProgress<string>? progreso = null, CancellationToken ct = default)
        {
            if (archivos.Count == 0) return new ResultadoCargaPrecios();

            var sucursales = new string[archivos.Count];
            var listas = new ListaLeida?[archivos.Count];
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
                    sucursales[i] = sucursal;
                    var diag = new DiagnosticoArchivo(ruta, sucursal);
                    try
                    {
                        listas[i] = LeerArchivo(ruta, diag);
                    }
                    catch (Exception ex)
                    {
                        // La excepción cruda no le dice nada al usuario: el clasificador la
                        // traduce a qué pasó y qué hacer.
                        listas[i] = new ListaLeida();
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

            var (datos, cruces) = await Task.Run(() => Consolidar(sucursales, listas, archivos), ct).ConfigureAwait(false);

            // Cuántos productos de una lista sin código encontraron pareja sólo se sabe al
            // consolidar, así que el aviso de esa lista se agrega recién acá.
            foreach (var (indice, cruce) in cruces)
                diagnosticos[indice].Registrar(ProblemaCarga.CruceSinCodigo, AvisoCruce.ParaDiagnostico(cruce));

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

        // ==========================================
        // Consolidación y cruce
        // ==========================================

        /// <summary>
        /// Une lo leído de cada archivo asignando ids de catálogo, en una única pasada
        /// secuencial (los catálogos no son thread-safe). Un archivo sin filas no genera
        /// sucursal: no habría nada que comparar en su columna.
        ///
        /// Primero entran las listas con código, en el orden de la carpeta y exactamente como
        /// antes: el código es la llave. Después, cada lista sin código toma el código de la
        /// lista con código cuya descripción coincide con la suya (igualdad tras compactar,
        /// nada aproximado). No se cruza cuando el nombre no identifica un único código: si en
        /// las otras listas ese nombre tiene más de un código, o si dos productos de la lista
        /// apuntan al mismo código. Esos productos, y los que no encontraron pareja, quedan
        /// con una clave propia por descripción: dos listas sin código igual se cruzan entre sí.
        /// </summary>
        internal static (ConjuntoPrecios datos, List<(int indice, CrucePorDescripcion cruce)> cruces) Consolidar(
            IReadOnlyList<string> sucursalesArchivo, ListaLeida?[] listas, IReadOnlyList<string> rutas)
        {
            int total = 0;
            foreach (var l in listas) total += l?.Filas.Count ?? 0;

            var sucursales = new Catalogo(8);
            var codigos = new Catalogo(Math.Max(64, total / 2));
            var descripciones = new Catalogo(Math.Max(64, total / 2));
            var columnas = new List<ColumnasLista>();

            // La sucursal de cada archivo con filas, en el orden de la carpeta.
            var sucursalDeArchivo = new int[listas.Length];
            bool hayPorDescripcion = false;
            for (int a = 0; a < listas.Length; a++)
            {
                sucursalDeArchivo[a] = -1;
                var l = listas[a];
                if (l == null || l.Filas.Count == 0) continue;

                int id = sucursales.Id(sucursalesArchivo[a]);
                sucursalDeArchivo[a] = id;
                while (columnas.Count <= id) columnas.Add(ColumnasLista.Ninguna);
                columnas[id] |= l.Columnas;
                hayPorDescripcion |= l.PorDescripcion;
            }

            var salida = new PrecioItem[total];
            int n = 0;

            // --- 1. Listas con código: igual que siempre ---
            for (int a = 0; a < listas.Length; a++)
            {
                var l = listas[a];
                if (l == null || sucursalDeArchivo[a] < 0 || l.PorDescripcion) continue;
                int sucId = sucursalDeArchivo[a];

                foreach (var f in l.Filas)
                    salida[n++] = new PrecioItem(sucId, codigos.Id(f.Codigo), descripciones.Id(f.Descripcion),
                                                 f.Costo, f.Impuesto, f.Utilidad, f.PrecioVenta);
                listas[a] = null;   // se suelta el buffer crudo en cuanto se consume
            }

            var cruces = new List<(int indice, CrucePorDescripcion cruce)>();
            if (hayPorDescripcion)
            {
                var indice = IndicePorDescripcion(salida, n, descripciones);
                var tramos = new List<(int archivo, int desde, int hasta, bool[] ambiguo, ListaLeida lista)>();

                // --- 2. Listas sin código: el código sale de la descripción ---
                for (int a = 0; a < listas.Length; a++)
                {
                    var l = listas[a];
                    if (l == null || sucursalDeArchivo[a] < 0 || !l.PorDescripcion) continue;
                    int sucId = sucursalDeArchivo[a];
                    int cuantas = l.Filas.Count;

                    var claves = new string[cuantas];
                    var candidato = new int[cuantas];
                    var usos = new Dictionary<int, int>();
                    for (int i = 0; i < cuantas; i++)
                    {
                        claves[i] = Compactar(l.Filas[i].Descripcion);
                        candidato[i] = indice.TryGetValue(claves[i], out int c) ? c : SinPareja;
                        if (candidato[i] >= 0) usos[candidato[i]] = usos.GetValueOrDefault(candidato[i]) + 1;
                    }

                    int desde = n;
                    var ambiguo = new bool[cuantas];
                    for (int i = 0; i < cuantas; i++)
                    {
                        var f = l.Filas[i];
                        int c = candidato[i];
                        ambiguo[i] = c == NombreAmbiguo || (c >= 0 && usos[c] > 1);
                        int codId = c >= 0 && !ambiguo[i]
                            ? c
                            : codigos.Id(ConjuntoPrecios.MarcaSinCodigo + claves[i]);

                        salida[n++] = new PrecioItem(sucId, codId, descripciones.Id(f.Descripcion),
                                                     f.Costo, f.Impuesto, f.Utilidad, f.PrecioVenta);
                    }

                    tramos.Add((a, desde, n, ambiguo, l));
                    listas[a] = null;
                }

                // --- 3. Conteos: un producto quedó cruzado si su fila la comparte otra sucursal ---
                var primeraSucursal = new int[codigos.Count];
                var enVarias = new bool[codigos.Count];
                Array.Fill(primeraSucursal, -1);
                for (int i = 0; i < n; i++)
                {
                    int c = salida[i].CodigoId, s = salida[i].SucursalId;
                    if (primeraSucursal[c] < 0) primeraSucursal[c] = s;
                    else if (primeraSucursal[c] != s) enVarias[c] = true;
                }

                foreach (var (archivo, desde, hasta, ambiguo, l) in tramos)
                {
                    int cruzados = 0, ambiguosSinCruzar = 0;
                    for (int i = desde; i < hasta; i++)
                    {
                        if (enVarias[salida[i].CodigoId]) cruzados++;
                        else if (ambiguo[i - desde]) ambiguosSinCruzar++;
                    }

                    cruces.Add((archivo, new CrucePorDescripcion
                    {
                        Sucursal = sucursalesArchivo[archivo],
                        Archivo = Path.GetFileName(rutas[archivo]),
                        Productos = hasta - desde,
                        Cruzados = cruzados,
                        NombresAmbiguos = ambiguosSinCruzar,
                        RepetidosFilas = l.RepetidosFilas,
                        RepetidosNombres = l.RepetidosNombres,
                        EjemplosRepetidos = l.EjemplosRepetidos.ToArray()
                    }));
                }
            }

            if (n != total) Array.Resize(ref salida, n);
            return (new ConjuntoPrecios(salida, sucursales, codigos, descripciones, columnas.ToArray(),
                                        cruces.Select(c => c.cruce).ToList()),
                    cruces);
        }

        private const int SinPareja = -1;
        private const int NombreAmbiguo = -2;

        /// <summary>
        /// Descripción compactada -> código, armado con las filas de las listas con código.
        /// Si el mismo nombre aparece con dos códigos distintos (en la misma lista o en dos
        /// sucursales) queda marcado como ambiguo y no sirve para cruzar.
        /// </summary>
        private static Dictionary<string, int> IndicePorDescripcion(PrecioItem[] filas, int cuantas, Catalogo descripciones)
        {
            var compactas = new string?[descripciones.Count];
            var indice = new Dictionary<string, int>(StringComparer.Ordinal);
            string relleno = Compactar(SinDescripcion);

            for (int i = 0; i < cuantas; i++)
            {
                ref readonly var f = ref filas[i];
                string clave = compactas[f.DescripcionId] ??= Compactar(descripciones[f.DescripcionId]);
                if (clave.Length == 0 || clave == relleno) continue;

                if (!indice.TryGetValue(clave, out int codigo)) indice[clave] = f.CodigoId;
                else if (codigo != f.CodigoId) indice[clave] = NombreAmbiguo;
            }
            return indice;
        }

        // ==========================================
        // Lectura de un archivo
        // ==========================================

        /// <summary>Posición de cada columna conocida en una fila candidata a encabezado (-1 si no está).</summary>
        internal struct Columnas
        {
            public int Codigo, Desc, Costo, CostoConImpuesto, Imp, Util, Precio;

            /// <summary>Texto original de la columna de costo con impuesto, para nombrarla en el aviso.</summary>
            public string NombreCostoConImpuesto;

            /// <summary>
            /// Hace falta algo con qué identificar el producto (el código o, si no hay, la
            /// descripción) y al menos una columna numérica para comparar. El costo con
            /// impuesto no cuenta: se ignora.
            /// </summary>
            public bool Completo => (Codigo != -1 || Desc != -1) && (Precio != -1 || Costo != -1 || Util != -1);

            public int Reconocidas =>
                (Codigo != -1 ? 1 : 0) + (Desc != -1 ? 1 : 0) + (Costo != -1 ? 1 : 0) + (CostoConImpuesto != -1 ? 1 : 0) +
                (Imp != -1 ? 1 : 0) + (Util != -1 ? 1 : 0) + (Precio != -1 ? 1 : 0);

            public static Columnas Vacias => new()
            {
                Codigo = -1, Desc = -1, Costo = -1, CostoConImpuesto = -1, Imp = -1, Util = -1, Precio = -1,
                NombreCostoConImpuesto = string.Empty
            };
        }

        /// <summary>
        /// Lee una lista de precios y deja en <paramref name="diag"/> el veredicto: qué
        /// encabezado encontró (o a qué se parecía lo que había), en qué hoja, qué columnas
        /// de la lista estándar le faltan y cuántas filas descartó y por qué.
        /// </summary>
        private static ListaLeida LeerArchivo(string ruta, DiagnosticoArchivo diag)
        {
            var resultado = new ListaLeida();
            var lista = resultado.Filas;

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
            List<string> columnasDeVentas = new();
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

                    var celdas = CeldasDeFila(reader);

                    // Acá la firma del otro sistema se mira ANTES que el encabezado propio, al
                    // revés que en el lector de ventas: el reporte de ventas trae «Artículo» y
                    // «% Utilidad», que este detector tomaría por código y utilidad, y cargaría
                    // la tabla dinámica entera como si fueran 4.000 productos. La firma pide
                    // «Año mes», «Total» y «Familia» juntas, que ninguna lista de precios trae,
                    // así que mirarla primero no puede rechazar una lista de verdad.
                    if (PareceEncabezadoVentas(celdas, out var deVentas))
                    {
                        pareceDeVentas = true;
                        columnasDeVentas = deVentas;
                        break;
                    }

                    var candidata = Detectar(celdas, out var reconocidas);
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
                        $"Es un reporte de ventas, no una lista de precios: tiene {Citar(columnasDeVentas)}.",
                        "Cargalo en la ventana principal (📂 Seleccionar Carpeta), no en la comparativa.");
                else if (!algunaFila)
                    diag.Registrar(ProblemaCarga.ArchivoVacio, "El archivo no tiene ninguna fila con datos.");
                else if (mejorReconocidas == null || mejorReconocidas.Count == 0)
                    diag.Registrar(ProblemaCarga.SinEncabezado,
                        "En las primeras 20 filas no hay nada que parezca un encabezado de precios " +
                        "(se buscan «Cód. Artículo», «Descripción», «Precio costo», «Porc. utilidad», «Precio IVI»).");
                else
                {
                    string donde = indiceHoja > 1 && mejorHoja.Length > 0 ? $" de la hoja «{mejorHoja}»" : string.Empty;
                    diag.Registrar(ProblemaCarga.SinEncabezado,
                        $"La fila {mejorFila}{donde} parece el encabezado de una lista de precios (tiene " +
                        $"{Citar(mejorReconocidas)}), pero así no se puede cargar: {RequisitoMinimo}. " +
                        $"Le faltan: {string.Join("; ", Faltantes(mejorCol, out _))}.",
                        SugerenciaListaEstandar);
                }
                return resultado;
            }

            if (hojaUsada > 0)
                diag.Registrar(ProblemaCarga.HojaEquivocada,
                    $"Los datos estaban en la hoja «{nombreHoja}», no en la primera. Se usó esa.");

            bool porDescripcion = col.Codigo == -1;
            resultado.PorDescripcion = porDescripcion;
            resultado.Columnas = (col.Costo != -1 ? ColumnasLista.Costo : ColumnasLista.Ninguna) |
                                 (col.Util != -1 ? ColumnasLista.Utilidad : ColumnasLista.Ninguna) |
                                 (col.Precio != -1 ? ColumnasLista.PrecioVenta : ColumnasLista.Ninguna);

            // Todas las columnas que faltan, juntas y con lo que implica cada una. Si lo único
            // que falta es «Imp. ventas», que no se usa al comparar, queda sólo en el detalle.
            var faltan = Faltantes(col, out bool afectan);
            if (faltan.Count > 0)
            {
                string texto = $"Le faltan columnas de la lista estándar: {string.Join("; ", faltan)}.";
                if (afectan) diag.Registrar(ProblemaCarga.EncabezadoIncompleto, texto);
                else diag.Detalles.Add(texto);
            }

            // Deduplicación local: los mismos textos se repiten miles de veces por archivo.
            var pool = new Dictionary<string, string>(1024, StringComparer.Ordinal);
            var vistos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Por qué se descarta cada fila que se descarta.
            int sinCodigo = 0, sinDescripcion = 0, filasDeTotal = 0, repetidos = 0, sinImporte = 0;

            while (reader.Read())
            {
                if (EsFilaVacia(reader)) continue;

                string codigo = string.Empty;
                string descripcion = col.Desc != -1 ? Texto(reader, col.Desc).Trim() : string.Empty;

                if (porDescripcion)
                {
                    // Sin código, la descripción es la llave: la fila de totales al pie no trae
                    // descripción, y un renglón que dice "Total" tampoco es un producto.
                    if (descripcion.Length == 0) { sinDescripcion++; continue; }
                    if (EsFilaDeTotal(descripcion)) { filasDeTotal++; continue; }
                }
                else
                {
                    codigo = Texto(reader, col.Codigo).Trim();
                    if (codigo.Length == 0) { sinCodigo++; continue; }

                    // Un total o subtotal al pie del reporte no es un producto.
                    if (Normalizar(codigo).Contains("total", StringComparison.Ordinal)) { filasDeTotal++; continue; }

                    // El mismo código dos veces en una sucursal dejaría la comparación ambigua;
                    // se conserva la primera aparición, que es la que ve el usuario en el Excel.
                    if (!vistos.Add(codigo)) { repetidos++; continue; }
                }

                if (descripcion.Length == 0) descripcion = SinDescripcion;

                // El costo con impuesto no se lee: no es comparable con el «Precio costo» sin
                // impuesto de las otras listas, y esta lista queda sin costo.
                LeerDecimal(reader, col.Costo, out decimal costo);
                LeerDecimal(reader, col.Imp, out decimal impuesto);
                LeerDecimal(reader, col.Util, out decimal utilidad);
                LeerDecimal(reader, col.Precio, out decimal precio);

                // Una fila sin ningún importe es un separador del reporte, no un producto.
                if (costo == 0m && precio == 0m && utilidad == 0m) { sinImporte++; continue; }

                lista.Add(new FilaPrecio(Interno(pool, codigo), Interno(pool, descripcion),
                                         costo, impuesto, utilidad, precio));
            }

            if (porDescripcion) QuitarDescripcionesRepetidas(resultado);

            diag.FilasLeidas = lista.Count;
            AnotarDescartes(diag, lista.Count, sinCodigo, sinDescripcion, filasDeTotal, repetidos,
                            resultado.RepetidosFilas, sinImporte);
            return resultado;
        }

        /// <summary>
        /// En una lista sin código, dos filas con la misma descripción (tras compactarla) no
        /// se pueden distinguir: "DURACELL BATERIA AAA UND" a ₡550 y a ₡3.500 son dos
        /// productos con el mismo nombre. Ninguna de las dos se cruza; quedan todas fuera.
        /// </summary>
        private static void QuitarDescripcionesRepetidas(ListaLeida resultado)
        {
            var lista = resultado.Filas;
            var grupos = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            var orden = new List<string>();
            for (int i = 0; i < lista.Count; i++)
            {
                string clave = Compactar(lista[i].Descripcion);
                if (!grupos.TryGetValue(clave, out var indices))
                {
                    grupos[clave] = indices = new List<int>(1);
                    orden.Add(clave);
                }
                indices.Add(i);
            }

            var excluir = new bool[lista.Count];
            foreach (string clave in orden)
            {
                var indices = grupos[clave];
                if (indices.Count < 2) continue;

                resultado.RepetidosNombres++;
                resultado.RepetidosFilas += indices.Count;
                if (resultado.EjemplosRepetidos.Count < 3)
                    resultado.EjemplosRepetidos.Add($"«{lista[indices[0]].Descripcion}» ({indices.Count} veces)");
                foreach (int i in indices) excluir[i] = true;
            }

            if (resultado.RepetidosFilas == 0) return;

            int destino = 0;
            for (int i = 0; i < lista.Count; i++)
                if (!excluir[i]) lista[destino++] = lista[i];
            lista.RemoveRange(destino, lista.Count - destino);
        }

        /// <summary>Renglones de total que la caja pone en la columna de descripción.</summary>
        private static bool EsFilaDeTotal(string descripcion)
        {
            string n = Compactar(descripcion);
            return n is "total" or "totales" or "total general" or "gran total" or "subtotal";
        }

        /// <summary>
        /// Columnas de la lista estándar que no están, en el orden en que la caja las exporta,
        /// cada una con lo que implica su ausencia. <paramref name="afectan"/> es falso si lo
        /// único que falta es «Imp. ventas», que la comparativa no usa.
        /// </summary>
        internal static List<string> Faltantes(Columnas col, out bool afectan)
        {
            var faltan = new List<string>();
            afectan = false;

            if (col.Codigo == -1 && col.Desc == -1)
            {
                faltan.Add("«Cód. Artículo» y «Descripción» (sin ninguna de las dos no hay con qué identificar los productos)");
                afectan = true;
            }
            else if (col.Codigo == -1)
            {
                faltan.Add("«Cód. Artículo» (se cruza por descripción y puede haber incongruencias)");
                afectan = true;
            }
            else if (col.Desc == -1)
            {
                faltan.Add("«Descripción» (los productos quedan como «Sin descripción»)");
                afectan = true;
            }

            if (col.Costo == -1)
            {
                faltan.Add(col.CostoConImpuesto != -1
                    ? $"«Precio costo» (trae «{col.NombreCostoConImpuesto}», que es el costo con impuesto y se ignora: " +
                      "el costo queda vacío para esta lista)"
                    : "«Precio costo» (el costo queda vacío para esta lista)");
                afectan = true;
            }

            if (col.Imp == -1)
                faltan.Add("«Imp. ventas» (no afecta la comparación)");

            if (col.Util == -1)
            {
                faltan.Add("«Porc. utilidad» (la utilidad queda vacía para esta lista)");
                afectan = true;
            }

            if (col.Precio == -1)
            {
                faltan.Add("«Precio IVI» (el precio de venta queda vacío para esta lista)");
                afectan = true;
            }

            return faltan;
        }

        /// <summary>
        /// Mensaje del lector de VENTAS cuando el archivo resulta ser una lista de precios: qué
        /// columnas de la lista tiene y cuáles de la lista estándar le faltan. No menciona
        /// columnas de ventas: decirle a quien cargó una lista de precios que "falta Total"
        /// lo manda a buscar una columna que su archivo nunca va a tener.
        /// </summary>
        internal static string DescribirListaParaVentas(IReadOnlyList<string> celdas)
        {
            var col = Detectar(celdas, out var reconocidas);
            var faltan = Faltantes(col, out _);

            var sb = new StringBuilder();
            sb.Append("Es una lista de precios, no un reporte de ventas: tiene ").Append(Citar(reconocidas)).Append('.');
            if (faltan.Count > 0)
                sb.Append(" Para la comparativa le faltan: ").Append(string.Join("; ", faltan)).Append('.');
            if (!col.Completo)
                sb.Append(" Así como está tampoco se puede cargar en la comparativa: ").Append(RequisitoMinimo).Append('.');
            return sb.ToString();
        }

        /// <summary>"«a», «b» y «c»".</summary>
        internal static string Citar(IReadOnlyList<string> textos)
            => AvisoCruce.Enumerar(textos.Select(t => $"«{t}»").ToList());

        /// <summary>
        /// Convierte los contadores en una frase. Un código o una descripción repetidos siempre
        /// son aviso: pierden datos en silencio. Los productos sin ningún importe avisan sólo si
        /// pasan del 5 % (una docena en 2.900 es normal y no vale la pena marcar el archivo por
        /// eso). Las filas sin código o sin descripción y los totales son parte del formato y
        /// sólo van al detalle.
        /// </summary>
        private static void AnotarDescartes(DiagnosticoArchivo diag, int cargadas, int sinCodigo, int sinDescripcion,
                                            int filasDeTotal, int repetidos, int descripcionRepetida, int sinImporte)
        {
            var f = ResumenDinamico.FormatoCR;
            var partes = new List<string>();
            if (sinImporte > 0) partes.Add($"{sinImporte.ToString("N0", f)} sin ningún importe (costo, precio y utilidad en 0)");
            if (repetidos > 0) partes.Add($"{repetidos.ToString("N0", f)} con código repetido (se tomó la primera aparición)");
            if (descripcionRepetida > 0)
                partes.Add($"{descripcionRepetida.ToString("N0", f)} con descripción repetida (quedaron fuera del cruce)");
            if (sinCodigo > 0) partes.Add($"{sinCodigo.ToString("N0", f)} sin código");
            if (sinDescripcion > 0) partes.Add($"{sinDescripcion.ToString("N0", f)} sin descripción");
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
            if (repetidos > 0 || descripcionRepetida > 0 || ExcelService.SuperaUmbral(sinImporte, cargadas))
                diag.Registrar(ProblemaCarga.FilasDescartadas, resumen);
            else diag.Detalles.Add(resumen);
        }

        /// <summary>
        /// Reconoce las columnas de una fila. El orden de las comprobaciones importa: el costo
        /// con impuesto («Costo I.V.A.») va primero porque si no "costo" lo tomaría por el
        /// costo sin impuesto; "Precio IVI - artículo" contiene tanto "precio" como
        /// "articulo", y "Precio costo" también contiene "precio". Devuelve además los textos
        /// originales de las columnas reconocidas, para decirle al usuario qué encontró.
        /// </summary>
        internal static Columnas Detectar(IReadOnlyList<string> celdas, out List<string> reconocidas)
        {
            var col = Columnas.Vacias;
            reconocidas = new List<string>();

            for (int c = 0; c < celdas.Count; c++)
            {
                string original = celdas[c].Trim();
                if (original.Length == 0) continue;
                string val = Normalizar(original);

                bool hit = true;
                if (val.Contains("costo", StringComparison.Ordinal) && TieneMarcaDeImpuesto(val))
                {
                    col.CostoConImpuesto = c;
                    col.NombreCostoConImpuesto = original;
                }
                else if (val.Contains("ivi", StringComparison.Ordinal) ||
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
