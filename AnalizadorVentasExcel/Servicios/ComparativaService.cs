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

            // --- Pasada 1: matriz densa código x sucursal ---
            int totalCodigos = _datos.Codigos.Count;
            var valores = new decimal[totalCodigos * cols];
            var presente = new bool[totalCodigos * cols];
            var descripcionId = new int[totalCodigos * cols];
            var presencia = new int[totalCodigos];

            foreach (ref readonly var f in _datos.Filas.AsSpan())
            {
                int col = posSucursal[f.SucursalId];
                if (col < 0) continue;

                int idx = f.CodigoId * cols + col;
                if (!presente[idx]) { presente[idx] = true; presencia[f.CodigoId]++; }
                valores[idx] = ConjuntoPrecios.ValorDe(in f, p.Metrica);
                descripcionId[idx] = f.DescripcionId;
            }

            // --- Pasada 2: una fila por producto ---
            bool mejorEsMayor = ConjuntoPrecios.MejorEsMayor(p.Metrica);
            string? busqueda = string.IsNullOrWhiteSpace(p.Busqueda) ? null : CeldasExcel.Normalizar(p.Busqueda);

            var filas = new List<FilaComparativa>();
            int comparables = 0, conDiferencia = 0, exclusivos = 0;
            decimal sumaPct = 0m;
            int conPct = 0;

            for (int codigo = 0; codigo < totalCodigos; codigo++)
            {
                int presentes = presencia[codigo];
                if (presentes == 0) continue;

                int baseIdx = codigo * cols;
                decimal minimo = decimal.MaxValue, maximo = decimal.MinValue;
                for (int c = 0; c < cols; c++)
                {
                    if (!presente[baseIdx + c]) continue;
                    decimal v = valores[baseIdx + c];
                    if (v < minimo) minimo = v;
                    if (v > maximo) maximo = v;
                }

                decimal diferencia = maximo - minimo;
                decimal diferenciaPct = minimo > 0m ? diferencia / minimo * 100m : 0m;

                if (presentes >= 2)
                {
                    comparables++;
                    if (diferencia != 0m) conDiferencia++;
                    if (minimo > 0m) { sumaPct += diferenciaPct; conPct++; }
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
                if (p.SoloConDiferencia && (presentes < 2 || diferencia == 0m)) continue;
                if (p.UmbralPct > 0m && diferenciaPct < p.UmbralPct) continue;

                var fila = ConstruirFila(codigo, baseIdx, cols, presentes, valores, presente,
                                         descripcionId, nombresSucursal, minimo, maximo,
                                         diferencia, diferenciaPct, p.Metrica, mejorEsMayor, busqueda);
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
            int codigo, int baseIdx, int cols, int presentes,
            decimal[] valores, bool[] presente, int[] descripcionId, List<string> nombresSucursal,
            decimal minimo, decimal maximo, decimal diferencia, decimal diferenciaPct,
            MetricaPrecio metrica, bool mejorEsMayor, string? busqueda)
        {
            decimal valorMejor = mejorEsMayor ? maximo : minimo;
            decimal valorPeor = mejorEsMayor ? minimo : maximo;
            bool hayDiferencia = presentes >= 2 && diferencia != 0m;

            var textos = new string[cols];
            var colores = new System.Windows.Media.Brush[cols];
            var listaValores = new decimal?[cols];
            var mejores = new List<string>();
            var peores = new List<string>();

            string descripcion = string.Empty;
            bool distintas = false;
            var detalle = new List<string>();

            for (int c = 0; c < cols; c++)
            {
                if (!presente[baseIdx + c])
                {
                    textos[c] = "—";
                    colores[c] = FilaComparativa.ColorAusente;
                    listaValores[c] = null;
                    continue;
                }

                decimal v = valores[baseIdx + c];
                listaValores[c] = v;
                textos[c] = FilaComparativa.Formatear(v, metrica);

                if (!hayDiferencia) colores[c] = FilaComparativa.ColorNormal;
                else if (v == valorMejor) { colores[c] = FilaComparativa.ColorMejor; mejores.Add(nombresSucursal[c]); }
                else if (v == valorPeor) { colores[c] = FilaComparativa.ColorPeor; peores.Add(nombresSucursal[c]); }
                else colores[c] = FilaComparativa.ColorNormal;

                string nombre = _datos.Descripciones[descripcionId[baseIdx + c]];
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
                ? "Sólo en " + PrimeraSucursal(baseIdx, cols, presente, nombresSucursal)
                : distintas ? "⚠ Nombres distintos" : string.Empty;

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
                Minimo = minimo,
                Maximo = maximo,
                Diferencia = hayDiferencia ? diferencia : 0m,
                DiferenciaPct = hayDiferencia ? diferenciaPct : 0m,
                DiferenciaTexto = presentes < 2 ? "-" : FilaComparativa.Formatear(diferencia, metrica),
                DiferenciaPctTexto = presentes < 2 ? "-"
                                   : minimo > 0m ? FilaComparativa.FormatearPorcentaje(diferenciaPct)
                                   : "-",
                Mejor = presentes < 2 ? "-" : !hayDiferencia ? "Todas iguales" : string.Join(", ", mejores),
                Peor = presentes < 2 || !hayDiferencia ? "-" : string.Join(", ", peores)
            };
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
            Comparison<FilaComparativa> comparar = orden switch
            {
                OrdenComparativa.DiferenciaAbs => (a, b) => b.Diferencia.CompareTo(a.Diferencia),
                OrdenComparativa.Codigo => (a, b) => string.Compare(a.Codigo, b.Codigo, StringComparison.CurrentCulture),
                OrdenComparativa.Descripcion => (a, b) => string.Compare(a.Descripcion, b.Descripcion, StringComparison.CurrentCulture),
                _ => (a, b) =>
                {
                    int c = b.DiferenciaPct.CompareTo(a.DiferenciaPct);
                    return c != 0 ? c : b.Diferencia.CompareTo(a.Diferencia);
                }
            };
            filas.Sort(comparar);
        }
    }
}
