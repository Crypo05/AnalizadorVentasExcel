using System;
using System.Collections.Generic;
using System.Linq;
using AnalizadorVentasExcel.Modelos;

namespace AnalizadorVentasExcel.Servicios
{
    /// <summary>Qué productos entran a la tabla según en cuántas sucursales aparecen.</summary>
    public enum PresenciaMinima
    {
        EnDosOMas = 0,
        EnTodas = 1,
        Todos = 2
    }

    public enum OrdenComparativa
    {
        DiferenciaPct = 0,
        DiferenciaAbs = 1,
        Codigo = 2,
        Descripcion = 3
    }

    public sealed class PeticionComparativa
    {
        public MetricaPrecio Metrica { get; init; }
        public IReadOnlyList<string> Sucursales { get; init; } = Array.Empty<string>();
        public PresenciaMinima Presencia { get; init; }
        public bool SoloConDiferencia { get; init; }

        /// <summary>Diferencia porcentual mínima para que el producto se liste (0 = sin umbral).</summary>
        public decimal UmbralPct { get; init; }
        public string? Busqueda { get; init; }
        public OrdenComparativa Orden { get; init; }
    }

    public sealed class ResultadoComparativa
    {
        public List<FilaComparativa> Filas { get; init; } = new();

        /// <summary>Sucursales comparadas, en el mismo orden que las columnas de la tabla.</summary>
        public List<string> Sucursales { get; init; } = new();

        /// <summary>Productos que existen en al menos dos de las sucursales comparadas.</summary>
        public int Comparables { get; init; }
        public int ConDiferencia { get; init; }
        public int Iguales { get; init; }

        /// <summary>Productos que sólo están en una sucursal: no se pueden comparar.</summary>
        public int Exclusivos { get; init; }
        public decimal DiferenciaPromedioPct { get; init; }
    }

    /// <summary>
    /// Cruce de las listas de precios por código de artículo.
    ///
    /// El código es la única llave fiable entre sucursales: las descripciones del mismo
    /// producto difieren de una tienda a otra ("PASTILLAS ALEVE GELS UND" contra
    /// "ALEVE GEL UND"), así que cruzar por nombre perdería la mayoría de los productos.
    /// Todo el cruce se resuelve con dos pasadas lineales sobre matrices densas
    /// código x sucursal indexadas por id, sin diccionarios por producto.
    /// </summary>
    public sealed class ComparativaService
    {
        private readonly ConjuntoPrecios _datos;

        public ComparativaService(ConjuntoPrecios datos) => _datos = datos;

        public ConjuntoPrecios Datos => _datos;

        public ResultadoComparativa Comparar(PeticionComparativa p)
        {
            // --- Columnas: las sucursales elegidas, en orden alfabético estable ---
            var idsSucursal = new List<int>(p.Sucursales.Count);
            foreach (string nombre in p.Sucursales)
            {
                int id = _datos.Sucursales.IdExistente(nombre);
                if (id >= 0) idsSucursal.Add(id);
            }
            idsSucursal.Sort((a, b) => string.Compare(_datos.Sucursales[a], _datos.Sucursales[b],
                                                      StringComparison.CurrentCulture));

            int cols = idsSucursal.Count;
            var nombresSucursal = idsSucursal.Select(id => _datos.Sucursales[id]).ToList();
            if (cols == 0) return new ResultadoComparativa();

            var posSucursal = new int[_datos.Sucursales.Count];
            for (int i = 0; i < posSucursal.Length; i++) posSucursal[i] = -1;
            for (int c = 0; c < cols; c++) posSucursal[idsSucursal[c]] = c;

            // La vista combinada trae dos métricas por sucursal (costo y venta); el resto,
            // una sola. Todo lo demás del cálculo es igual, así que se recorre un arreglo
            // de métricas en vez de duplicar el motor.
            var metricas = ConjuntoPrecios.EsCombinada(p.Metrica)
                ? ConjuntoPrecios.MetricasCombinadas
                : new[] { p.Metrica };
            int m = metricas.Length;

            // --- Pasada 1: matriz densa código x métrica x sucursal ---
            int totalCodigos = _datos.Codigos.Count;
            var valores = new decimal[totalCodigos * m * cols];
            var presente = new bool[totalCodigos * cols];
            var descripcionId = new int[totalCodigos * cols];
            var presencia = new int[totalCodigos];

            foreach (ref readonly var f in _datos.Filas.AsSpan())
            {
                int col = posSucursal[f.SucursalId];
                if (col < 0) continue;

                int idx = f.CodigoId * cols + col;
                if (!presente[idx]) { presente[idx] = true; presencia[f.CodigoId]++; }
                descripcionId[idx] = f.DescripcionId;

                for (int k = 0; k < m; k++)
                    valores[(f.CodigoId * m + k) * cols + col] = ConjuntoPrecios.ValorDe(in f, metricas[k]);
            }

            // --- Pasada 2: una fila por producto ---
            string? busqueda = string.IsNullOrWhiteSpace(p.Busqueda) ? null : CeldasExcel.Normalizar(p.Busqueda);

            var filas = new List<FilaComparativa>();
            int comparables = 0, conDiferencia = 0, exclusivos = 0;
            decimal sumaPct = 0m;
            int conPct = 0;

            var resumen = new ResumenMetrica[m];

            for (int codigo = 0; codigo < totalCodigos; codigo++)
            {
                int presentes = presencia[codigo];
                if (presentes == 0) continue;

                int basePresencia = codigo * cols;
                decimal mayorPct = 0m, mayorAbs = 0m;
                bool algunaDifiere = false, algunPctValido = false;

                for (int k = 0; k < m; k++)
                {
                    int baseValores = (codigo * m + k) * cols;
                    decimal minimo = decimal.MaxValue, maximo = decimal.MinValue;
                    for (int c = 0; c < cols; c++)
                    {
                        if (!presente[basePresencia + c]) continue;
                        decimal v = valores[baseValores + c];
                        if (v < minimo) minimo = v;
                        if (v > maximo) maximo = v;
                    }

                    decimal dif = maximo - minimo;
                    decimal pct = minimo > 0m ? dif / minimo * 100m : 0m;
                    resumen[k] = new ResumenMetrica(baseValores, minimo, maximo, dif, pct);

                    if (dif != 0m) algunaDifiere = true;
                    if (dif > mayorAbs) mayorAbs = dif;
                    if (minimo > 0m)
                    {
                        algunPctValido = true;
                        if (pct > mayorPct) mayorPct = pct;
                    }
                }

                if (presentes >= 2)
                {
                    comparables++;
                    if (algunaDifiere) conDiferencia++;
                    if (algunPctValido) { sumaPct += mayorPct; conPct++; }
                }
                else exclusivos++;

                // --- Filtros de la vista ---
                bool pasaPresencia = p.Presencia switch
                {
                    PresenciaMinima.EnTodas => presentes == cols,
                    PresenciaMinima.Todos => true,
                    _ => presentes >= 2
                };
                if (!pasaPresencia) continue;
                if (p.SoloConDiferencia && (presentes < 2 || !algunaDifiere)) continue;
                if (p.UmbralPct > 0m && mayorPct < p.UmbralPct) continue;

                var fila = ConstruirFila(codigo, basePresencia, cols, presentes, valores, presente,
                                         descripcionId, nombresSucursal, metricas, resumen,
                                         mayorPct, mayorAbs, busqueda);
                if (fila != null) filas.Add(fila);
            }

            Ordenar(filas, p.Orden);

            return new ResultadoComparativa
            {
                Filas = filas,
                Sucursales = nombresSucursal,
                Comparables = comparables,
                ConDiferencia = conDiferencia,
                Iguales = comparables - conDiferencia,
                Exclusivos = exclusivos,
                DiferenciaPromedioPct = conPct > 0 ? sumaPct / conPct : 0m
            };
        }

        /// <summary>
        /// Arma la fila visible. Devuelve null cuando el producto no coincide con la
        /// búsqueda: el texto buscado puede estar en el código o en cualquiera de las
        /// descripciones, y esas sólo se conocen aquí.
        /// </summary>
        private FilaComparativa? ConstruirFila(
            int codigo, int basePresencia, int cols, int presentes,
            decimal[] valores, bool[] presente, int[] descripcionId, List<string> nombresSucursal,
            MetricaPrecio[] metricas, ResumenMetrica[] resumen,
            decimal mayorPct, decimal mayorAbs, string? busqueda)
        {
            int m = metricas.Length;

            // Una entrada por (sucursal, métrica), intercaladas: con una sola métrica queda
            // igual que antes, y con dos la sucursal s ocupa las posiciones 2s y 2s+1.
            var textos = new string[cols * m];
            var colores = new System.Windows.Media.Brush[cols * m];
            var listaValores = new decimal?[cols * m];

            var mejores = new List<string>();
            var peores = new List<string>();

            for (int k = 0; k < m; k++)
            {
                ref readonly var r = ref resumen[k];
                bool mejorEsMayor = ConjuntoPrecios.MejorEsMayor(metricas[k]);
                bool hayDiferencia = presentes >= 2 && r.Diferencia != 0m;
                decimal valorMejor = mejorEsMayor ? r.Maximo : r.Minimo;
                decimal valorPeor = mejorEsMayor ? r.Minimo : r.Maximo;

                for (int c = 0; c < cols; c++)
                {
                    int destino = c * m + k;

                    if (!presente[basePresencia + c])
                    {
                        textos[destino] = "—";
                        colores[destino] = FilaComparativa.ColorAusente;
                        listaValores[destino] = null;
                        continue;
                    }

                    decimal v = valores[r.BaseValores + c];
                    listaValores[destino] = v;
                    textos[destino] = FilaComparativa.Formatear(v, metricas[k]);

                    if (!hayDiferencia) colores[destino] = FilaComparativa.ColorNormal;
                    else if (v == valorMejor)
                    {
                        colores[destino] = FilaComparativa.ColorMejor;
                        if (k == 0) mejores.Add(nombresSucursal[c]);
                    }
                    else if (v == valorPeor)
                    {
                        colores[destino] = FilaComparativa.ColorPeor;
                        if (k == 0) peores.Add(nombresSucursal[c]);
                    }
                    else colores[destino] = FilaComparativa.ColorNormal;
                }
            }

            // Las descripciones no dependen de la métrica: se recorren una sola vez.
            string descripcion = string.Empty;
            bool distintas = false;
            var detalle = new List<string>();

            for (int c = 0; c < cols; c++)
            {
                if (!presente[basePresencia + c]) continue;

                string nombre = _datos.Descripciones[descripcionId[basePresencia + c]];
                detalle.Add($"{nombresSucursal[c]}: {nombre}");
                if (descripcion.Length == 0) descripcion = nombre;
                else if (!distintas && !CoincidenNombres(descripcion, nombre)) distintas = true;
            }

            if (busqueda != null)
            {
                string codigoTexto = _datos.Codigos[codigo];
                bool coincide = CeldasExcel.Normalizar(codigoTexto).Contains(busqueda, StringComparison.Ordinal);
                if (!coincide)
                    foreach (string d in detalle)
                        if (CeldasExcel.Normalizar(d).Contains(busqueda, StringComparison.Ordinal)) { coincide = true; break; }
                if (!coincide) return null;
            }

            string aviso = presentes < 2
                ? "Sólo en " + PrimeraSucursal(basePresencia, cols, presente, nombresSucursal)
                : distintas ? "⚠ Nombres distintos" : string.Empty;

            // La primera métrica manda en las columnas de siempre (con la vista combinada,
            // el costo); la segunda, cuando existe, va a las columnas de venta.
            var principal = resumen[0];
            bool difPrincipal = presentes >= 2 && principal.Diferencia != 0m;
            var secundaria = m > 1 ? resumen[1] : default;
            bool difSecundaria = m > 1 && presentes >= 2 && secundaria.Diferencia != 0m;

            return new FilaComparativa
            {
                Codigo = _datos.Codigos[codigo],
                Descripcion = descripcion,
                DetalleDescripciones = string.Join("  •  ", detalle),
                DescripcionesDistintas = distintas,
                Valores = listaValores,
                Textos = textos,
                Colores = colores,
                Presencia = presentes,
                PresenciaTexto = $"{presentes}/{cols}",
                Aviso = aviso,
                Minimo = principal.Minimo,
                Maximo = principal.Maximo,
                Diferencia = difPrincipal ? principal.Diferencia : 0m,
                DiferenciaPct = difPrincipal ? principal.DiferenciaPct : 0m,
                DiferenciaTexto = presentes < 2 ? "-" : FilaComparativa.Formatear(principal.Diferencia, metricas[0]),
                DiferenciaPctTexto = presentes < 2 ? "-"
                                   : principal.Minimo > 0m ? FilaComparativa.FormatearPorcentaje(principal.DiferenciaPct)
                                   : "-",
                DiferenciaVenta = difSecundaria ? secundaria.Diferencia : 0m,
                DiferenciaVentaPct = difSecundaria ? secundaria.DiferenciaPct : 0m,
                DiferenciaVentaTexto = m < 2 || presentes < 2 ? "-" : FilaComparativa.Formatear(secundaria.Diferencia, metricas[1]),
                DiferenciaVentaPctTexto = m < 2 || presentes < 2 ? "-"
                                        : secundaria.Minimo > 0m ? FilaComparativa.FormatearPorcentaje(secundaria.DiferenciaPct)
                                        : "-",
                DiferenciaMayorPct = presentes < 2 ? 0m : mayorPct,
                DiferenciaMayorAbs = presentes < 2 ? 0m : mayorAbs,
                Mejor = presentes < 2 ? "-" : !difPrincipal ? "Todas iguales" : string.Join(", ", mejores),
                Peor = presentes < 2 || !difPrincipal ? "-" : string.Join(", ", peores)
            };
        }

        /// <summary>Lo que se resuelve por métrica antes de armar la fila.</summary>
        private readonly struct ResumenMetrica
        {
            public readonly int BaseValores;
            public readonly decimal Minimo, Maximo, Diferencia, DiferenciaPct;

            public ResumenMetrica(int baseValores, decimal minimo, decimal maximo,
                                  decimal diferencia, decimal diferenciaPct)
            {
                BaseValores = baseValores; Minimo = minimo; Maximo = maximo;
                Diferencia = diferencia; DiferenciaPct = diferenciaPct;
            }
        }

        private static string PrimeraSucursal(int baseIdx, int cols, bool[] presente, List<string> nombres)
        {
            for (int c = 0; c < cols; c++) if (presente[baseIdx + c]) return nombres[c];
            return string.Empty;
        }

        /// <summary>
        /// Dos nombres del mismo producto se consideran iguales si sólo cambian tildes,
        /// mayúsculas o espacios de más. Todo lo demás se marca para que el usuario lo revise.
        /// </summary>
        private static bool CoincidenNombres(string a, string b)
            => string.Equals(Compactar(a), Compactar(b), StringComparison.Ordinal);

        private static string Compactar(string texto)
        {
            string n = CeldasExcel.Normalizar(texto);
            var sb = new System.Text.StringBuilder(n.Length);
            bool espacio = false;
            foreach (char c in n)
            {
                if (char.IsWhiteSpace(c)) { espacio = sb.Length > 0; continue; }
                if (espacio) { sb.Append(' '); espacio = false; }
                sb.Append(c);
            }
            return sb.ToString();
        }

        private static void Ordenar(List<FilaComparativa> filas, OrdenComparativa orden)
        {
            // Se ordena por la MAYOR de las diferencias del producto. Con una sola métrica
            // es exactamente su diferencia; con la vista combinada, el producto sube si se
            // separa entre sucursales, sin importar si se separa al comprarlo o al venderlo.
            Comparison<FilaComparativa> comparar = orden switch
            {
                OrdenComparativa.DiferenciaAbs => (a, b) => b.DiferenciaMayorAbs.CompareTo(a.DiferenciaMayorAbs),
                OrdenComparativa.Codigo => (a, b) => string.Compare(a.Codigo, b.Codigo, StringComparison.CurrentCulture),
                OrdenComparativa.Descripcion => (a, b) => string.Compare(a.Descripcion, b.Descripcion, StringComparison.CurrentCulture),
                _ => (a, b) =>
                {
                    int c = b.DiferenciaMayorPct.CompareTo(a.DiferenciaMayorPct);
                    return c != 0 ? c : b.DiferenciaMayorAbs.CompareTo(a.DiferenciaMayorAbs);
                }
            };
            filas.Sort(comparar);
        }
    }
}
