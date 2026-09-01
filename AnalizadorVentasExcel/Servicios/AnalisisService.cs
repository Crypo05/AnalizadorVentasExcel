using System;
using System.Collections.Generic;
using System.Linq;
using AnalizadorVentasExcel.Modelos;

namespace AnalizadorVentasExcel.Servicios
{
    public enum Operacion { Suma, Promedio, Conteo }

    public sealed class SerieGrafico
    {
        public string Nombre { get; init; } = string.Empty;
        public double[] Valores { get; init; } = Array.Empty<double>();
    }

    public sealed class ResultadoAnalisis
    {
        public List<ResumenDinamico> Tabla { get; init; } = new();
        public List<string> EtiquetasX { get; init; } = new();
        public List<SerieGrafico> Series { get; init; } = new();
        public int FilasFiltradas { get; init; }
    }

    public sealed class PeticionAnalisis
    {
        public Dimension EjeX { get; init; }
        public IReadOnlyList<Dimension> Desglose { get; init; } = Array.Empty<Dimension>();
        public Operacion Operacion { get; init; }
        public IReadOnlyList<string> Sucursales { get; init; } = Array.Empty<string>();
        public IReadOnlyList<string> Periodos { get; init; } = Array.Empty<string>();
        public IReadOnlyList<string> Proveedores { get; init; } = Array.Empty<string>();
        public IReadOnlyList<string> Familias { get; init; } = Array.Empty<string>();
        public string TextoOperacion { get; init; } = "Suma Total (Colones)";
    }

    /// <summary>
    /// Motor de filtrado y agregación.
    ///
    /// La versión anterior recorría el conjunto filtrado una vez por cada combinación de
    /// etiqueta del eje X y serie (hasta 20 x 10 = 200 recorridos completos de 250.000 filas)
    /// y volvía a recorrerlo para ordenar las etiquetas. Aquí todo se resuelve con unas pocas
    /// pasadas lineales sobre índices enteros, y la tabla y el gráfico comparten la misma
    /// agregación en lugar de calcularla dos veces.
    /// </summary>
    public sealed class AnalisisService
    {
        private const int MaxEtiquetasX = 20;
        private const int MaxSeries = 10;

        private readonly ConjuntoDatos _datos;

        // Buffers reutilizados entre llamadas para no reasignar en cada cambio de filtro.
        private int[] _indicesFiltrados = Array.Empty<int>();

        public AnalisisService(ConjuntoDatos datos) => _datos = datos;

        public ConjuntoDatos Datos => _datos;

        // ==========================================
        // Filtrado
        // ==========================================

        /// <summary>
        /// Convierte una lista de textos seleccionados en una máscara indexada por id.
        /// Sustituye al List&lt;string&gt;.Contains() por fila, que era una búsqueda lineal.
        /// </summary>
        private static bool[] Mascara(Catalogo catalogo, IReadOnlyList<string> seleccionados)
        {
            var mascara = new bool[catalogo.Count];
            for (int i = 0; i < seleccionados.Count; i++)
            {
                int id = catalogo.IdExistente(seleccionados[i]);
                if (id >= 0) mascara[id] = true;
            }
            return mascara;
        }

        private int Filtrar(PeticionAnalisis p, out int[] indices)
        {
            var mSuc = Mascara(_datos.Sucursales, p.Sucursales);
            var mPer = Mascara(_datos.Periodos, p.Periodos);
            var mPro = Mascara(_datos.Proveedores, p.Proveedores);
            var mFam = Mascara(_datos.Familias, p.Familias);

            var filas = _datos.Filas;
            if (_indicesFiltrados.Length < filas.Length)
                _indicesFiltrados = new int[filas.Length];

            var buf = _indicesFiltrados;
            int n = 0;
            for (int i = 0; i < filas.Length; i++)
            {
                ref readonly var f = ref filas[i];
                if (mSuc[f.SucursalId] && mPer[f.PeriodoId] && mPro[f.ProveedorId] && mFam[f.FamiliaId])
                    buf[n++] = i;
            }
            indices = buf;
            return n;
        }

        // ==========================================
        // Agregación
        // ==========================================

        private struct Acumulador
        {
            public decimal Suma;
            public decimal SumaUtilidad;
            public int Conteo;
            public int XSlot;
            public int SerieSlot;
        }

        /// <summary>Clave compuesta del desglose (hasta 4 dimensiones), sin asignar memoria.</summary>
        private readonly struct ClaveSerie : IEquatable<ClaveSerie>
        {
            private readonly int _a, _b, _c, _d;
            public ClaveSerie(int a, int b, int c, int d) { _a = a; _b = b; _c = c; _d = d; }
            public bool Equals(ClaveSerie o) => _a == o._a && _b == o._b && _c == o._c && _d == o._d;
            public override bool Equals(object? o) => o is ClaveSerie k && Equals(k);
            public override int GetHashCode() => HashCode.Combine(_a, _b, _c, _d);
            public int At(int i) => i switch { 0 => _a, 1 => _b, 2 => _c, _ => _d };
        }

        private static ClaveSerie ClaveDe(in VentaItem f, IReadOnlyList<Dimension> dims)
        {
            int a = dims.Count > 0 ? ConjuntoDatos.IdDe(f, dims[0]) : -1;
            int b = dims.Count > 1 ? ConjuntoDatos.IdDe(f, dims[1]) : -1;
            int c = dims.Count > 2 ? ConjuntoDatos.IdDe(f, dims[2]) : -1;
            int d = dims.Count > 3 ? ConjuntoDatos.IdDe(f, dims[3]) : -1;
            return new ClaveSerie(a, b, c, d);
        }

        public ResultadoAnalisis Analizar(PeticionAnalisis p)
        {
            int n = Filtrar(p, out int[] indices);
            var filas = _datos.Filas;
            var dims = p.Desglose;
            bool hayDesglose = dims.Count > 0;

            // --- Pasada 1: numerar en orden de aparición las etiquetas X y las series ---
            var xSlots = new Dictionary<int, int>();            // idDimensionX -> slot denso
            var xIds = new List<int>();
            var serieSlots = new Dictionary<ClaveSerie, int>(); // clave desglose -> slot denso
            var serieClaves = new List<ClaveSerie>();

            // --- y acumular por celda (x, serie) en una sola pasada ---
            var celdaSlots = new Dictionary<long, int>();
            var celdas = new List<Acumulador>();

            double sumaGlobalD = 0d;
            var ejeX = p.EjeX;

            for (int k = 0; k < n; k++)
            {
                ref readonly var f = ref filas[indices[k]];
                sumaGlobalD += (double)f.TotalVenta;

                int xId = ConjuntoDatos.IdDe(f, ejeX);
                if (!xSlots.TryGetValue(xId, out int xSlot))
                {
                    xSlot = xIds.Count;
                    xSlots[xId] = xSlot;
                    xIds.Add(xId);
                }

                int serieSlot = 0;
                if (hayDesglose)
                {
                    var clave = ClaveDe(f, dims);
                    if (!serieSlots.TryGetValue(clave, out serieSlot))
                    {
                        serieSlot = serieClaves.Count;
                        serieSlots[clave] = serieSlot;
                        serieClaves.Add(clave);
                    }
                }

                long claveCelda = ((long)xSlot << 32) | (uint)serieSlot;
                if (!celdaSlots.TryGetValue(claveCelda, out int slot))
                {
                    slot = celdas.Count;
                    celdaSlots[claveCelda] = slot;
                    celdas.Add(new Acumulador { XSlot = xSlot, SerieSlot = serieSlot });
                }

                var arr = System.Runtime.InteropServices.CollectionsMarshal.AsSpan(celdas);
                ref var acc = ref arr[slot];
                acc.Suma += f.TotalVenta;
                acc.SumaUtilidad += f.PorcentajeUtilidad;
                acc.Conteo++;
            }

            // --- Tabla ---
            var catX = _datos.CatalogoDe(ejeX);
            var tabla = new List<ResumenDinamico>(celdas.Count);
            var celdasSpan = System.Runtime.InteropServices.CollectionsMarshal.AsSpan(celdas);

            for (int i = 0; i < celdasSpan.Length; i++)
            {
                ref var acc = ref celdasSpan[i];
                double valor = Valor(in acc, p.Operacion);
                var fila = new ResumenDinamico
                {
                    Etiqueta = catX[xIds[acc.XSlot]],
                    DetalleSecundario = hayDesglose ? NombreSerie(serieClaves[acc.SerieSlot], dims) : "Total General",
                    ValorNumerico = valor,
                    MargenPromedio = acc.Conteo > 0 ? (double)(acc.SumaUtilidad / acc.Conteo) : 0d,
                    Participacion = (p.Operacion == Operacion.Suma && sumaGlobalD > 0)
                        ? (valor / sumaGlobalD).ToString("P1", ResumenDinamico.FormatoCR)
                        : "-"
                };
                fila.Formatear(p.TextoOperacion);
                tabla.Add(fila);
            }

            tabla = ejeX == Dimension.AnioMes
                ? tabla.OrderBy(x => x.Etiqueta, StringComparer.CurrentCulture).ThenByDescending(x => x.ValorNumerico).ToList()
                : tabla.OrderByDescending(x => x.ValorNumerico).ToList();

            // --- Gráfico: se reutiliza la agregación anterior ---
            var (etiquetasX, series) = ConstruirGrafico(
                celdas, xIds, catX, serieClaves, dims, ejeX, p.Operacion, hayDesglose);

            return new ResultadoAnalisis
            {
                Tabla = tabla,
                EtiquetasX = etiquetasX,
                Series = series,
                FilasFiltradas = n
            };
        }

        private (List<string>, List<SerieGrafico>) ConstruirGrafico(
            List<Acumulador> celdas, List<int> xIds, Catalogo catX,
            List<ClaveSerie> serieClaves, IReadOnlyList<Dimension> dims,
            Dimension ejeX, Operacion op, bool hayDesglose)
        {
            int totalX = xIds.Count;
            int totalSeries = hayDesglose ? serieClaves.Count : 1;
            var celdasSpan = System.Runtime.InteropServices.CollectionsMarshal.AsSpan(celdas);

            // Totales por etiqueta X y por serie, para elegir los "top N".
            var accX = new Acumulador[totalX];
            var accSerie = new Acumulador[totalSeries];
            for (int i = 0; i < celdasSpan.Length; i++)
            {
                ref var c = ref celdasSpan[i];
                ref var ax = ref accX[c.XSlot];
                ax.Suma += c.Suma; ax.SumaUtilidad += c.SumaUtilidad; ax.Conteo += c.Conteo;
                ref var asr = ref accSerie[c.SerieSlot];
                asr.Suma += c.Suma; asr.SumaUtilidad += c.SumaUtilidad; asr.Conteo += c.Conteo;
            }

            // Selección de etiquetas del eje X.
            int[] xElegidos;
            bool esTiempo = ejeX == Dimension.AnioMes;
            if (esTiempo)
            {
                xElegidos = Enumerable.Range(0, totalX)
                    .OrderBy(s => catX[xIds[s]], StringComparer.CurrentCulture)
                    .ToArray();
            }
            else
            {
                xElegidos = Enumerable.Range(0, totalX)
                    .OrderByDescending(s => Valor(in accX[s], op))
                    .Take(MaxEtiquetasX)
                    .ToArray();
            }

            var etiquetas = new List<string>(xElegidos.Length);
            var posicionX = new int[totalX];
            for (int i = 0; i < posicionX.Length; i++) posicionX[i] = -1;
            for (int i = 0; i < xElegidos.Length; i++)
            {
                posicionX[xElegidos[i]] = i;
                etiquetas.Add(catX[xIds[xElegidos[i]]]);
            }

            // Selección de series.
            int[] seriesElegidas = hayDesglose
                ? Enumerable.Range(0, totalSeries)
                    .OrderByDescending(s => Valor(in accSerie[s], op))
                    .Take(MaxSeries)
                    .ToArray()
                : new[] { 0 };

            var posicionSerie = new int[totalSeries];
            for (int i = 0; i < posicionSerie.Length; i++) posicionSerie[i] = -1;
            for (int i = 0; i < seriesElegidas.Length; i++) posicionSerie[seriesElegidas[i]] = i;

            // Matriz densa serie x etiqueta, rellenada en una pasada sobre las celdas.
            int filas = seriesElegidas.Length, cols = etiquetas.Count;
            var matriz = new Acumulador[filas * cols];
            for (int i = 0; i < celdasSpan.Length; i++)
            {
                ref var c = ref celdasSpan[i];
                int px = posicionX[c.XSlot];
                if (px < 0) continue;
                int ps = posicionSerie[c.SerieSlot];
                if (ps < 0) continue;
                ref var m = ref matriz[ps * cols + px];
                m.Suma += c.Suma; m.SumaUtilidad += c.SumaUtilidad; m.Conteo += c.Conteo;
            }

            var series = new List<SerieGrafico>(filas);
            for (int s = 0; s < filas; s++)
            {
                var valores = new double[cols];
                for (int x = 0; x < cols; x++) valores[x] = Valor(in matriz[s * cols + x], op);
                series.Add(new SerieGrafico
                {
                    Nombre = hayDesglose ? NombreSerie(serieClaves[seriesElegidas[s]], dims) : "Total",
                    Valores = valores
                });
            }

            return (etiquetas, series);
        }

        private static double Valor(in Acumulador a, Operacion op) => op switch
        {
            Operacion.Suma => (double)a.Suma,
            Operacion.Promedio => a.Conteo > 0 ? (double)(a.SumaUtilidad / a.Conteo) : 0d,
            _ => a.Conteo
        };

        private string NombreSerie(ClaveSerie clave, IReadOnlyList<Dimension> dims)
        {
            if (dims.Count == 0) return string.Empty;
            if (dims.Count == 1) return _datos.CatalogoDe(dims[0])[clave.At(0)];

            var partes = new string[dims.Count];
            for (int i = 0; i < dims.Count; i++) partes[i] = _datos.CatalogoDe(dims[i])[clave.At(i)];
            return string.Join(" - ", partes);
        }

        // ==========================================
        // Explorador de productos (auditoría)
        // ==========================================

        public sealed class ProductoExplorado
        {
            public int NormalizadoId { get; init; }
            public ResumenDinamico Fila { get; init; } = new();
        }

        /// <summary>
        /// Consolida por nombre de artículo (recortado) mostrando en cuántas sucursales
        /// está disponible. Una sola pasada, contra los N recorridos anidados anteriores.
        /// </summary>
        public List<ProductoExplorado> Explorar(IReadOnlyList<string> periodos, IReadOnlyList<string> sucursales)
        {
            var mPer = Mascara(_datos.Periodos, periodos);
            var mSuc = Mascara(_datos.Sucursales, sucursales);
            var filas = _datos.Filas;
            var mapNorm = _datos.ArticuloANormalizado;

            int totalNorm = _datos.ArticulosNormalizados.Count;
            int totalSuc = _datos.Sucursales.Count;

            var acc = new Acumulador[totalNorm];
            var visto = new bool[totalNorm];
            // Bitset sucursales por producto: normalmente son pocas sucursales.
            var sucursalesPorProducto = new ulong[totalNorm * ((totalSuc + 63) / 64)];
            int palabras = (totalSuc + 63) / 64;

            for (int i = 0; i < filas.Length; i++)
            {
                ref readonly var f = ref filas[i];
                if (!mPer[f.PeriodoId] || !mSuc[f.SucursalId]) continue;

                int id = mapNorm[f.ArticuloId];
                visto[id] = true;
                ref var a = ref acc[id];
                a.Suma += f.TotalVenta;
                a.SumaUtilidad += f.PorcentajeUtilidad;
                a.Conteo++;
                sucursalesPorProducto[id * palabras + (f.SucursalId >> 6)] |= 1UL << (f.SucursalId & 63);
            }

            var resultado = new List<ProductoExplorado>();
            var nombresSuc = new List<string>(totalSuc);

            for (int id = 0; id < totalNorm; id++)
            {
                if (!visto[id]) continue;

                nombresSuc.Clear();
                for (int s = 0; s < totalSuc; s++)
                    if ((sucursalesPorProducto[id * palabras + (s >> 6)] & (1UL << (s & 63))) != 0)
                        nombresSuc.Add(_datos.Sucursales[s]);
                nombresSuc.Sort(StringComparer.CurrentCulture);

                ref var a = ref acc[id];
                var fila = new ResumenDinamico
                {
                    Etiqueta = _datos.ArticulosNormalizados[id],
                    DetalleSecundario = string.Join(", ", nombresSuc),
                    ValorNumerico = (double)a.Suma,
                    MargenPromedio = a.Conteo > 0 ? (double)(a.SumaUtilidad / a.Conteo) : 0d,
                    Participacion = $"{nombresSuc.Count} Tiendas"
                };
                fila.Formatear("Suma");
                resultado.Add(new ProductoExplorado { NormalizadoId = id, Fila = fila });
            }

            resultado.Sort((a, b) => string.Compare(a.Fila.Etiqueta, b.Fila.Etiqueta, StringComparison.CurrentCulture));
            return resultado;
        }

        /// <summary>
        /// Serie de margen por sucursal y periodo para un producto del explorador.
        /// Devuelve null en los meses sin venta para que la línea se corte.
        /// </summary>
        public (List<string> sucursales, List<double?[]> valores) MargenPorSucursal(
            int normalizadoId, IReadOnlyList<string> periodosOrdenados)
        {
            int totalSuc = _datos.Sucursales.Count;
            var posPeriodo = new Dictionary<int, int>(periodosOrdenados.Count);
            for (int i = 0; i < periodosOrdenados.Count; i++)
            {
                int id = _datos.Periodos.IdExistente(periodosOrdenados[i]);
                if (id >= 0) posPeriodo[id] = i;
            }

            int cols = periodosOrdenados.Count;
            var sumas = new decimal[totalSuc * cols];
            var conteos = new int[totalSuc * cols];
            var usada = new bool[totalSuc];

            var filas = _datos.Filas;
            var mapNorm = _datos.ArticuloANormalizado;
            for (int i = 0; i < filas.Length; i++)
            {
                ref readonly var f = ref filas[i];
                if (mapNorm[f.ArticuloId] != normalizadoId) continue;
                if (!posPeriodo.TryGetValue(f.PeriodoId, out int col)) continue;
                int idx = f.SucursalId * cols + col;
                sumas[idx] += f.PorcentajeUtilidad;
                conteos[idx]++;
                usada[f.SucursalId] = true;
            }

            var nombres = new List<string>();
            var valores = new List<double?[]>();
            var orden = Enumerable.Range(0, totalSuc).Where(s => usada[s])
                                  .OrderBy(s => _datos.Sucursales[s], StringComparer.CurrentCulture);

            foreach (int s in orden)
            {
                var v = new double?[cols];
                for (int c = 0; c < cols; c++)
                {
                    int idx = s * cols + c;
                    v[c] = conteos[idx] > 0 ? (double)(sumas[idx] / conteos[idx]) * 100d : null;
                }
                nombres.Add(_datos.Sucursales[s]);
                valores.Add(v);
            }

            return (nombres, valores);
        }
    }
}
