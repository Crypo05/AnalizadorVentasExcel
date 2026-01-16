using System;
using System.Collections.Concurrent; // Para listas seguras en paralelo
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks; // Para async/await
using System.Windows;
using System.Windows.Controls;
using System.Diagnostics;
using System.Net.Http;
using Microsoft.Win32;
using ExcelDataReader; // NUEVO MOTOR DE ALTO RENDIMIENTO
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using SkiaSharp;
using LiveChartsCore.Measure;
using LiveChartsCore.Kernel;

namespace AnalizadorVentasExcel
{
    public partial class MainWindow : Window
    {
        // ==========================================
        // CONFIGURACIÓN GENERAL
        // ==========================================
        private const string VersionActual = "2.0.0"; // Versión High Performance
        private const string UrlVersionRemota = "https://raw.githubusercontent.com/TU_USUARIO/TU_REPO/main/version.txt";
        private const string UrlDescarga = "https://github.com/TU_USUARIO/TU_REPO/raw/main/AnalizadorVentasExcel.exe";

        private List<VentaItem> _datosGlobales = new List<VentaItem>();
        private bool _cargandoFiltros = false;
        private bool _modoExploracion = false;
        private CultureInfo _culturaCR;

        public MainWindow()
        {
            InitializeComponent();

            // CRÍTICO: Registra la codificación para leer Excel viejos y nuevos
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            LimpiarVersionesAntiguas();
            ConfigurarCulturaManual();
            CargarOpcionesDesglose();
            this.Title = $"Analizador Corporativo v{VersionActual} | Desarrollado por Mateo Sanabria";

            // Evento para el gráfico interactivo
            GridResultados.SelectionChanged += GridResultados_SelectionChanged;
        }

        // ==========================================
        // 1. CARGA MASIVA OPTIMIZADA (PARALLEL + READER)
        // ==========================================
        private async void BtnCargarCarpeta_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Title = "Seleccione cualquier archivo Excel de la carpeta", Filter = "Excel|*.xlsx;*.xls", CheckFileExists = true };

            if (dialog.ShowDialog() == true)
            {
                TxtEstadoArchivo.Text = "Iniciando motor de alto rendimiento...";
                BtnCargarCarpeta.IsEnabled = false;

                string carpeta = Path.GetDirectoryName(dialog.FileName);
                string[] archivos = Directory.GetFiles(carpeta, "*.xls*");

                if (archivos.Length == 0) return;

                string modo = (CmbTipoNegocio.SelectedItem as ComboBoxItem)?.Content.ToString();

                // ConcurrentBag es seguro para agregar datos desde múltiples hilos a la vez
                var datosTemporales = new ConcurrentBag<VentaItem>();
                var errores = new ConcurrentBag<string>();

                Stopwatch cronometro = Stopwatch.StartNew();

                try
                {
                    // Ejecutamos en segundo plano para no congelar la ventana
                    await Task.Run(() =>
                    {
                        var servicio = new ExcelServiceOptimizado();

                        // PARALELISMO: Procesa múltiples archivos simultáneamente usando todos los núcleos del CPU
                        Parallel.ForEach(archivos, archivo =>
                        {
                            try
                            {
                                string sucursal = Path.GetFileNameWithoutExtension(archivo);
                                var resultados = servicio.CargarDatosRapido(archivo, modo, sucursal);

                                foreach (var item in resultados) datosTemporales.Add(item);
                            }
                            catch (Exception ex)
                            {
                                errores.Add($"\n{Path.GetFileName(archivo)}: {ex.Message}");
                            }
                        });
                    });

                    cronometro.Stop();
                    _datosGlobales = datosTemporales.ToList();

                    if (_datosGlobales.Any())
                    {
                        TxtEstadoArchivo.Text = $"Carga Rápida: {_datosGlobales.Count:N0} registros en {cronometro.Elapsed.TotalSeconds:N1} seg.";
                        TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Green;

                        InicializarFiltros();
                        AplicarFiltros();

                        if (!errores.IsEmpty) MessageBox.Show($"Advertencia: Algunos archivos no se pudieron leer.\n{string.Join("", errores.Take(5))}...");
                    }
                    else
                    {
                        TxtEstadoArchivo.Text = "No se encontraron datos válidos.";
                        TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Red;
                    }
                }
                catch (Exception ex) { MessageBox.Show($"Error Crítico: {ex.Message}"); }
                finally { BtnCargarCarpeta.IsEnabled = true; }
            }
        }

        // ==========================================
        // 2. MODO EXPLORADOR DE PRODUCTOS
        // ==========================================
        private void BtnAuditar_Click(object sender, RoutedEventArgs e)
        {
            if (!_datosGlobales.Any()) { MessageBox.Show("Primero cargue datos.", "Sin Datos"); return; }

            _modoExploracion = true;

            // Obtener filtros actuales
            var fechasSeleccionadas = ObtenerSeleccionados(LstFiltroFecha);
            var sucursalesSeleccionadas = ObtenerSeleccionados(LstFiltroSucursal);

            if (!fechasSeleccionadas.Any()) fechasSeleccionadas = _datosGlobales.Select(x => x.Periodo).Distinct().ToList();
            if (!sucursalesSeleccionadas.Any()) sucursalesSeleccionadas = _datosGlobales.Select(x => x.Sucursal).Distinct().ToList();

            // Filtrar datos base
            var datosFiltrados = _datosGlobales.Where(x =>
                fechasSeleccionadas.Contains(x.Periodo) &&
                sucursalesSeleccionadas.Contains(x.Sucursal)
            ).ToList();

            if (!datosFiltrados.Any()) { MessageBox.Show("No hay datos para los filtros seleccionados."); return; }

            // Generar Tabla Consolidada
            var consolidado = datosFiltrados
                .GroupBy(x => x.ArticuloNombre.Trim())
                .Select(g => new ResumenDinamico
                {
                    Etiqueta = g.Key,
                    // Detalle: Lista de sucursales donde existe
                    DetalleSecundario = string.Join(", ", g.Select(x => x.Sucursal).Distinct().OrderBy(s => s)),
                    ValorNumerico = (double)g.Sum(x => x.TotalVenta),
                    MargenPromedio = g.Any() ? (double)g.Average(x => x.PorcentajeUtilidad) : 0,
                    TipoFormato = "Suma",
                    // Disponibilidad: Conteo de tiendas
                    Participacion = $"{g.Select(x => x.Sucursal).Distinct().Count()} Tiendas"
                })
                .OrderBy(x => x.Etiqueta)
                .ToList();

            GridResultados.ItemsSource = consolidado;

            TxtTituloReporte.Text = "📦 Explorador de Productos";
            TxtSubtitulo.Text = $"Viendo {consolidado.Count} productos únicos. Seleccione uno para comparar sucursales.";

            if (ColumnaValor != null) ColumnaValor.Header = "Venta Total";
            if (ColumnaParticipacion != null) ColumnaParticipacion.Header = "Disponibilidad";

            if (GraficoVentas != null) GraficoVentas.Series = new ISeries[] { };

            MessageBox.Show("Explorador Listo.\nSeleccione un producto en la tabla para ver la COMPARATIVA DE SUCURSALES mes a mes.", "Modo Comparativo");
        }

        // Evento: Al seleccionar un producto, mostrar GRÁFICO MULTI-SUCURSAL
        private void GridResultados_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_modoExploracion || GridResultados.SelectedItem == null) return;

            var itemSeleccionado = GridResultados.SelectedItem as ResumenDinamico;
            if (itemSeleccionado == null) return;

            var nombreProducto = itemSeleccionado.Etiqueta;

            // Filtros actuales
            var fechasSeleccionadas = ObtenerSeleccionados(LstFiltroFecha);
            if (!fechasSeleccionadas.Any()) fechasSeleccionadas = _datosGlobales.Select(x => x.Periodo).Distinct().ToList();

            // Ordenar fechas cronológicamente para el Eje X
            var fechasOrdenadas = fechasSeleccionadas.OrderBy(x => x).ToList();

            // Filtrar datos crudos del producto seleccionado
            var datosProducto = _datosGlobales
                .Where(x => x.ArticuloNombre.Trim() == nombreProducto && fechasSeleccionadas.Contains(x.Periodo))
                .ToList();

            ActualizarGraficoComparativo(datosProducto, fechasOrdenadas, nombreProducto);
        }

        // Gráfico: Comparativa Sucursales Mes a Mes
        private void ActualizarGraficoComparativo(List<VentaItem> datos, List<string> mesesEjeX, string nombreProducto)
        {
            if (GraficoVentas == null) return;
            GraficoVentas.TooltipFindingStrategy = TooltipFindingStrategy.CompareOnlyX;

            var listaSeries = new List<ISeries>();
            var datosPorSucursal = datos.GroupBy(x => x.Sucursal).OrderBy(g => g.Key).ToList();

            foreach (var grupoSucursal in datosPorSucursal)
            {
                var nombreSucursal = grupoSucursal.Key;
                var valores = new List<double?>();

                foreach (var mes in mesesEjeX)
                {
                    var ventaMes = grupoSucursal.Where(x => x.Periodo == mes).ToList();
                    if (ventaMes.Any()) valores.Add((double)ventaMes.Average(x => x.PorcentajeUtilidad) * 100);
                    else valores.Add(null); // Hueco en el gráfico si no hubo venta
                }

                listaSeries.Add(new LineSeries<double?>
                {
                    Name = nombreSucursal,
                    Values = valores,
                    LineSmoothness = 0,
                    GeometrySize = 8,
                    Stroke = new SolidColorPaint { StrokeThickness = 3 },
                    Fill = null,
                    TooltipLabelFormatter = p => $"{p.Context.Series.Name}: {p.Model:N2}%"
                });
            }

            GraficoVentas.Series = listaSeries.ToArray();
            GraficoVentas.XAxes = new Axis[] { new Axis { Labels = mesesEjeX, LabelsRotation = 0, TextSize = 12, Name = "Comparativa Mensual" } };
            GraficoVentas.YAxes = new Axis[] { new Axis { Labeler = v => $"{v:N0}%", Name = $"Margen Utilidad: {nombreProducto}" } };
        }

        // ==========================================
        // 3. MÉTODOS ESTÁNDAR (Filtros y Pivot)
        // ==========================================
        private void AplicarFiltros()
        {
            _modoExploracion = false;
            if (GridResultados == null || CmbAgrupacion == null) return;
            if (ColumnaParticipacion != null) ColumnaParticipacion.Header = "% Part.";

            var itemEjeX = CmbAgrupacion.SelectedItem as ComboBoxItem;
            var itemOp = CmbOperacion.SelectedItem as ComboBoxItem;
            if (itemEjeX == null || itemOp == null) return;

            string ejeX = itemEjeX.Content.ToString();
            string operacion = itemOp.Content.ToString();
            var dimensionesSerie = ObtenerSeleccionados(LstDesglose);
            bool hayDesglose = dimensionesSerie.Any();

            var sSucs = ObtenerSeleccionados(LstFiltroSucursal);
            var sFechas = ObtenerSeleccionados(LstFiltroFecha);
            var sProvs = ObtenerSeleccionados(LstFiltroProveedor);
            var sFams = ObtenerSeleccionados(LstFiltroFamilia);

            if (!sSucs.Any()) { GridResultados.ItemsSource = null; return; }

            var datos = _datosGlobales.Where(x =>
                sSucs.Contains(x.Sucursal) && sFechas.Contains(x.Periodo) &&
                sProvs.Contains(x.Proveedor) && sFams.Contains(x.Familia)).ToList();

            double sumaGlobal = datos.Sum(x => (double)x.TotalVenta);
            Func<IEnumerable<VentaItem>, double> calcMargen = (g) => g.Any() ? (double)g.Average(x => x.PorcentajeUtilidad) : 0;

            List<ResumenDinamico> resumenTabla;

            if (hayDesglose)
            {
                resumenTabla = datos.GroupBy(x => new { KeyX = ObtenerLlaveSimple(x, ejeX), KeySerie = ObtenerLlaveCompuesta(x, dimensionesSerie) })
                    .Select(g => new ResumenDinamico
                    {
                        Etiqueta = g.Key.KeyX,
                        DetalleSecundario = g.Key.KeySerie,
                        ValorNumerico = CalcularValor(g, operacion),
                        MargenPromedio = calcMargen(g),
                        TipoFormato = operacion,
                        Participacion = (operacion.Contains("Suma") && sumaGlobal > 0) ? (CalcularValor(g, operacion) / sumaGlobal).ToString("P1", _culturaCR) : "-"
                    }).OrderByDescending(x => x.ValorNumerico).ToList();
            }
            else
            {
                resumenTabla = datos.GroupBy(x => ObtenerLlaveSimple(x, ejeX))
                    .Select(g => new ResumenDinamico
                    {
                        Etiqueta = g.Key,
                        DetalleSecundario = "Total General",
                        ValorNumerico = CalcularValor(g, operacion),
                        MargenPromedio = calcMargen(g),
                        TipoFormato = operacion,
                        Participacion = (operacion.Contains("Suma") && sumaGlobal > 0) ? (CalcularValor(g, operacion) / sumaGlobal).ToString("P1", _culturaCR) : "-"
                    }).OrderByDescending(x => x.ValorNumerico).ToList();
            }
            if (ejeX == "Año Mes") resumenTabla = resumenTabla.OrderBy(x => x.Etiqueta).ThenByDescending(x => x.ValorNumerico).ToList();

            GridResultados.ItemsSource = resumenTabla;
            if (ColumnaValor != null) ColumnaValor.Header = operacion;
            TxtTituloReporte.Text = hayDesglose ? $"Análisis: {ejeX} vs Series" : $"Total por {ejeX}";
            TxtSubtitulo.Text = $"{datos.Count} registros filtrados.";
            ActualizarGraficoMultiNivel(datos, ejeX, dimensionesSerie, operacion);
        }

        // ==========================================
        // 4. INFRAESTRUCTURA Y HELPERS
        // ==========================================
        private double CalcularValor(IEnumerable<VentaItem> datos, string operacion) { if (operacion.Contains("Suma")) return (double)datos.Sum(x => x.TotalVenta); if (operacion.Contains("Promedio")) return (double)(datos.Any() ? datos.Average(x => x.PorcentajeUtilidad) : 0); return datos.Count(); }
        private string ObtenerLlaveSimple(VentaItem item, string criterio) { switch (criterio) { case "Año Mes": return item.Periodo; case "Proveedor": return item.Proveedor; case "Familia": return item.Familia; case "Sucursal": return item.Sucursal; case "Articulo": return item.ArticuloNombre; default: return "General"; } }
        private string ObtenerLlaveCompuesta(VentaItem item, List<string> dimensiones) { if (!dimensiones.Any()) return ""; var partes = new List<string>(); foreach (var dim in dimensiones) partes.Add(ObtenerLlaveSimple(item, dim)); return string.Join(" - ", partes); }

        private void ActualizarGraficoMultiNivel(List<VentaItem> datos, string ejeX, List<string> dimensionesSerie, string operacion)
        {
            if (GraficoVentas == null) return;
            GraficoVentas.Series = new ISeries[] { };
            GraficoVentas.TooltipFindingStrategy = TooltipFindingStrategy.CompareOnlyX;
            var etiquetasX = datos.Select(x => ObtenerLlaveSimple(x, ejeX)).Distinct().ToList();
            bool esTiempo = (ejeX == "Año Mes");
            if (esTiempo) etiquetasX = etiquetasX.OrderBy(x => x).ToList();
            else etiquetasX = etiquetasX.OrderByDescending(lbl => CalcularValor(datos.Where(d => ObtenerLlaveSimple(d, ejeX) == lbl), operacion)).Take(20).ToList();
            var listaSeries = new List<ISeries>();
            Func<ChartPoint, string> tp = point => { double val = point.PrimaryValue; if (Math.Abs(val) < 0.01) return null; return $"{point.Context.Series.Name}: {val.ToString("N0", _culturaCR)}"; };
            if (dimensionesSerie.Any())
            {
                var top = datos.GroupBy(x => ObtenerLlaveCompuesta(x, dimensionesSerie)).OrderByDescending(g => CalcularValor(g, operacion)).Take(10).Select(g => g.Key).ToList();
                foreach (var s in top)
                {
                    var v = new List<double>(); foreach (var x in etiquetasX) v.Add(CalcularValor(datos.Where(d => ObtenerLlaveSimple(d, ejeX) == x && ObtenerLlaveCompuesta(d, dimensionesSerie) == s), operacion));
                    if (esTiempo) listaSeries.Add(new LineSeries<double> { Name = s, Values = v, LineSmoothness = 0, GeometrySize = 8, Stroke = new SolidColorPaint { StrokeThickness = 3 }, Fill = null, TooltipLabelFormatter = tp });
                    else listaSeries.Add(new ColumnSeries<double> { Name = s, Values = v, TooltipLabelFormatter = tp });
                }
            }
            else
            {
                var v = new List<double>(); foreach (var x in etiquetasX) v.Add(CalcularValor(datos.Where(d => ObtenerLlaveSimple(d, ejeX) == x), operacion));
                listaSeries.Add(new ColumnSeries<double> { Name = "Total", Values = v, Fill = new SolidColorPaint(SKColors.DarkCyan), TooltipLabelFormatter = tp });
            }
            GraficoVentas.Series = listaSeries.ToArray();
            GraficoVentas.XAxes = new Axis[] { new Axis { Labels = etiquetasX, LabelsRotation = 25, TextSize = 11 } };
            GraficoVentas.YAxes = new Axis[] { new Axis { Labeler = val => val.ToString("N0", _culturaCR) } };
        }

        // Métodos de Soporte UI
        private void LimpiarVersionesAntiguas() { try { string p = Process.GetCurrentProcess().MainModule.FileName + ".old"; if (File.Exists(p)) File.Delete(p); } catch { } }
        private void CargarOpcionesDesglose() { LstDesglose.ItemsSource = new List<string> { "Proveedor", "Familia", "Sucursal", "Año Mes" }; }
        private void ConfigurarCulturaManual() { _culturaCR = (CultureInfo)CultureInfo.CreateSpecificCulture("es-CR").Clone(); _culturaCR.NumberFormat.CurrencySymbol = "₡"; CultureInfo.DefaultThreadCurrentCulture = _culturaCR; CultureInfo.DefaultThreadCurrentUICulture = _culturaCR; }
        private async void BtnActualizar_Click(object sender, RoutedEventArgs e) { MessageBox.Show("Sistema Actualizado."); }

        // Inicialización de Filtros
        private void InicializarFiltros() { _cargandoFiltros = true; LstFiltroSucursal.ItemsSource = _datosGlobales.Select(x => x.Sucursal).Distinct().OrderBy(x => x).ToList(); LstFiltroSucursal.SelectAll(); LstFiltroFecha.ItemsSource = _datosGlobales.Select(x => x.Periodo).Distinct().OrderByDescending(x => x).ToList(); LstFiltroFecha.SelectAll(); LstFiltroProveedor.ItemsSource = _datosGlobales.Select(x => x.Proveedor).Distinct().OrderBy(x => x).ToList(); LstFiltroProveedor.SelectAll(); ActualizarChecklistFamilias(); _cargandoFiltros = false; }
        private void ActualizarChecklistFamilias() { if (LstFiltroProveedor == null) return; var p = ObtenerSeleccionados(LstFiltroProveedor); var s = ObtenerSeleccionados(LstFiltroSucursal); var q = _datosGlobales.Where(x => p.Contains(x.Proveedor) && s.Contains(x.Sucursal)); LstFiltroFamilia.ItemsSource = q.Select(x => x.Familia).Distinct().OrderBy(x => x).ToList(); LstFiltroFamilia.SelectAll(); }
        private List<string> ObtenerSeleccionados(ListBox lb) { var l = new List<string>(); if (lb.SelectedItems == null) return l; foreach (var i in lb.SelectedItems) l.Add(i is ListBoxItem bi ? bi.Content.ToString() : i.ToString()); return l; }
        private void BtnSelectAllSucursal_Click(object s, RoutedEventArgs e) => LstFiltroSucursal.SelectAll();
        private void BtnSelectAllFecha_Click(object s, RoutedEventArgs e) => LstFiltroFecha.SelectAll();
        private void BtnSelectAllProv_Click(object s, RoutedEventArgs e) => LstFiltroProveedor.SelectAll();
        private void BtnSelectAllFam_Click(object s, RoutedEventArgs e) => LstFiltroFamilia.SelectAll();
        private void AplicarFiltros_Event(object s, RoutedEventArgs e) { if (!_cargandoFiltros) AplicarFiltros(); }
        private void AplicarFiltros_Event(object s, SelectionChangedEventArgs e) { if (!_cargandoFiltros) AplicarFiltros(); }
        private void LstFiltroProveedor_SelectionChanged(object s, SelectionChangedEventArgs e) { if (!_cargandoFiltros) { ActualizarChecklistFamilias(); AplicarFiltros(); } }
    }

    // ==========================================
    // CLASES DE MODELO
    // ==========================================
    public class VentaItem { public string Sucursal { get; set; } public string Periodo { get; set; } public string ArticuloCodigo { get; set; } public string ArticuloNombre { get; set; } public string Proveedor { get; set; } public string Familia { get; set; } public decimal TotalVenta { get; set; } public decimal PorcentajeUtilidad { get; set; } }

    public class ResumenDinamico
    {
        public string Etiqueta { get; set; }
        public string DetalleSecundario { get; set; }
        public double ValorNumerico { get; set; }
        public double MargenPromedio { get; set; }
        public string TipoFormato { get; set; }
        public string Participacion { get; set; }
        public string ValorFormateado { get { var cr = CultureInfo.GetCultureInfo("es-CR"); var n = (NumberFormatInfo)cr.NumberFormat.Clone(); n.CurrencySymbol = "₡"; if (TipoFormato.Contains("Suma")) return ValorNumerico.ToString("C2", n); if (TipoFormato.Contains("Promedio")) return ValorNumerico.ToString("P2", n); return ValorNumerico.ToString("N0", n); } }
        public string MargenFormateado { get { return MargenPromedio.ToString("P2", CultureInfo.GetCultureInfo("es-CR")); } }
    }

    // ==========================================
    // SERVICIO EXCEL OPTIMIZADO (EXCEL DATA READER)
    // ==========================================
    public class ExcelServiceOptimizado
    {
        public List<VentaItem> CargarDatosRapido(string ruta, string modo, string nombreSucursal)
        {
            var lista = new List<VentaItem>();

            using (var stream = File.Open(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int colFecha = -1, colCodigo = -1, colDesc = -1, colProv = -1, colFam = -1, colTotal = -1, colUtil = -1;
                    bool encabezadoEncontrado = false;

                    // 1. Detección de encabezados (Primeras 20 filas)
                    int filasLeidas = 0;
                    while (reader.Read() && filasLeidas < 20)
                    {
                        filasLeidas++;
                        if (encabezadoEncontrado) break;

                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            var valor = reader.GetValue(i)?.ToString()?.ToLower().Trim();
                            if (string.IsNullOrEmpty(valor)) continue;

                            if (valor.Contains("año") || valor == "mes") colFecha = i;
                            else if (valor == "artículo" || valor == "articulo") colCodigo = i;
                            else if (valor.Contains("desc") || valor.Contains("nombre")) colDesc = i;
                            else if (valor.Contains("proveedor")) colProv = i;
                            else if (valor.Contains("familia")) colFam = i;
                            else if (valor == "total" || valor == "total venta") colTotal = i;
                            else if (valor.Contains("utilidad") || valor.Contains("%")) colUtil = i;
                        }

                        if (colTotal != -1 && colFam != -1) { encabezadoEncontrado = true; break; }
                    }

                    if (!encabezadoEncontrado) return lista;

                    bool esMinimarket = (colCodigo != -1);
                    if (modo != null && modo.Contains("Minimarket")) esMinimarket = true;
                    else if (modo != null && modo.Contains("Souvenir")) esMinimarket = false;

                    string ultPeriodo = "", ultProv = "General", ultFam = "General";

                    // 2. Lectura rápida de datos
                    while (reader.Read())
                    {
                        try
                        {
                            // Fill Down (Fecha)
                            if (colFecha != -1)
                            {
                                var val = reader.GetValue(colFecha);
                                if (val != null)
                                {
                                    if (val is DateTime dt) ultPeriodo = dt.ToString("yyyy-MM");
                                    else ultPeriodo = val.ToString();
                                }
                            }

                            // Fill Down (Proveedor)
                            if (colProv != -1)
                            {
                                var val = reader.GetValue(colProv)?.ToString();
                                if (!string.IsNullOrEmpty(val) && !val.ToLower().Contains("total")) ultProv = val;
                            }

                            // Fill Down (Familia)
                            if (colFam != -1)
                            {
                                var val = reader.GetValue(colFam)?.ToString();
                                if (!string.IsNullOrEmpty(val)) ultFam = val;
                            }

                            // Filtros Vacíos
                            if (esMinimarket) { if (colCodigo != -1 && (reader.GetValue(colCodigo) == null)) continue; }
                            else
                            {
                                if (colFam != -1 && string.IsNullOrEmpty(reader.GetValue(colFam)?.ToString())) continue;
                                if (colProv != -1) { var p = reader.GetValue(colProv)?.ToString(); if (p != null && p.ToLower().Contains("total")) continue; }
                            }

                            // Venta
                            if (colTotal == -1) continue;
                            decimal total = 0;
                            var objTotal = reader.GetValue(colTotal);
                            if (!ParseObjetoDecimal(objTotal, out total)) continue;
                            if (total == 0) continue;

                            // Utilidad
                            decimal utilidad = 0;
                            if (colUtil != -1)
                            {
                                var objUtil = reader.GetValue(colUtil);
                                ParseObjetoDecimal(objUtil, out utilidad);
                            }

                            // Nombre / Descripción
                            string nombreReal = "Sin Nombre";
                            if (esMinimarket && colDesc != -1)
                            {
                                var n = reader.GetValue(colDesc)?.ToString();
                                if (!string.IsNullOrEmpty(n)) nombreReal = n;
                            }
                            else if (!esMinimarket)
                            {
                                nombreReal = ultFam;
                            }

                            if (!string.IsNullOrEmpty(ultPeriodo) && !ultPeriodo.ToLower().Contains("año"))
                            {
                                lista.Add(new VentaItem
                                {
                                    Sucursal = nombreSucursal,
                                    Periodo = ultPeriodo,
                                    ArticuloCodigo = (colCodigo != -1) ? reader.GetValue(colCodigo)?.ToString() : "",
                                    ArticuloNombre = nombreReal,
                                    Proveedor = ultProv,
                                    Familia = ultFam,
                                    TotalVenta = total,
                                    PorcentajeUtilidad = utilidad
                                });
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            return lista;
        }

        private bool ParseObjetoDecimal(object valor, out decimal resultado)
        {
            resultado = 0;
            if (valor == null) return false;
            if (valor is double d) { resultado = (decimal)d; return true; }
            if (valor is decimal dec) { resultado = dec; return true; }
            if (valor is int i) { resultado = i; return true; }

            string s = valor.ToString();
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Replace("%", "").Replace("$", "").Replace("₡", "").Trim();

            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out resultado)) return true;
            var cr = CultureInfo.GetCultureInfo("es-CR");
            if (decimal.TryParse(s, NumberStyles.Any, cr, out resultado)) return true;

            return false;
        }
    }
}