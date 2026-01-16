using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Diagnostics;
using System.Net.Http;
using Microsoft.Win32;
using ClosedXML.Excel;
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
        // CONFIGURACIÓN
        // ==========================================
        private const string VersionActual = "1.6.0"; // Versión Gráfico Multi-Sucursal
        private const string UrlVersionRemota = "https://raw.githubusercontent.com/TU_USUARIO/TU_REPO/main/version.txt";
        private const string UrlDescarga = "https://github.com/TU_USUARIO/TU_REPO/raw/main/AnalizadorVentasExcel.exe";

        private List<VentaItem> _datosGlobales = new List<VentaItem>();
        private bool _cargandoFiltros = false;
        private bool _modoExploracion = false;
        private CultureInfo _culturaCR;

        public MainWindow()
        {
            InitializeComponent();
            LimpiarVersionesAntiguas();
            ConfigurarCulturaManual();
            CargarOpcionesDesglose();
            this.Title = $"Analizador Corporativo v{VersionActual} | Desarrollado por Mateo Sanabria";

            // Evento para ver detalle al seleccionar un producto
            GridResultados.SelectionChanged += GridResultados_SelectionChanged;
        }

        // ==========================================
        // MODO EXPLORADOR
        // ==========================================
        private void BtnAuditar_Click(object sender, RoutedEventArgs e)
        {
            if (!_datosGlobales.Any())
            {
                MessageBox.Show("Primero cargue datos.", "Sin Datos");
                return;
            }

            _modoExploracion = true;

            // 1. Obtener filtros actuales
            var fechasSeleccionadas = ObtenerSeleccionados(LstFiltroFecha);
            var sucursalesSeleccionadas = ObtenerSeleccionados(LstFiltroSucursal);

            if (!fechasSeleccionadas.Any()) fechasSeleccionadas = _datosGlobales.Select(x => x.Periodo).Distinct().ToList();
            if (!sucursalesSeleccionadas.Any()) sucursalesSeleccionadas = _datosGlobales.Select(x => x.Sucursal).Distinct().ToList();

            // 2. Filtrar
            var datosFiltrados = _datosGlobales.Where(x =>
                fechasSeleccionadas.Contains(x.Periodo) &&
                sucursalesSeleccionadas.Contains(x.Sucursal)
            ).ToList();

            if (!datosFiltrados.Any()) { MessageBox.Show("No hay datos para los filtros seleccionados."); return; }

            // 3. Tabla Consolidada (Muestra disponibilidad de tiendas)
            var consolidado = datosFiltrados
                .GroupBy(x => x.ArticuloNombre.Trim())
                .Select(g => new ResumenDinamico
                {
                    Etiqueta = g.Key,
                    // Detalle: Lista de sucursales
                    DetalleSecundario = string.Join(", ", g.Select(x => x.Sucursal).Distinct().OrderBy(s => s)),
                    ValorNumerico = (double)g.Sum(x => x.TotalVenta),
                    MargenPromedio = g.Any() ? (double)g.Average(x => x.PorcentajeUtilidad) : 0,
                    TipoFormato = "Suma",
                    // Disponibilidad: Cantidad de tiendas
                    Participacion = $"{g.Select(x => x.Sucursal).Distinct().Count()} Tiendas"
                })
                .OrderBy(x => x.Etiqueta)
                .ToList();

            GridResultados.ItemsSource = consolidado;

            // Títulos
            TxtTituloReporte.Text = "📦 Explorador de Productos";
            TxtSubtitulo.Text = $"Viendo {consolidado.Count} productos únicos. Seleccione uno para comparar sucursales.";

            if (ColumnaValor != null) ColumnaValor.Header = "Venta Total";
            if (ColumnaParticipacion != null) ColumnaParticipacion.Header = "Disponibilidad";

            // Limpiar gráfico
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

        // ==========================================
        // NUEVO GRÁFICO: Comparativa Sucursales Mes a Mes
        // ==========================================
        private void ActualizarGraficoComparativo(List<VentaItem> datos, List<string> mesesEjeX, string nombreProducto)
        {
            if (GraficoVentas == null) return;

            // Esto permite ver los tooltips de todas las sucursales al mismo tiempo al pasar el mouse por el mes
            GraficoVentas.TooltipFindingStrategy = TooltipFindingStrategy.CompareOnlyX;

            var listaSeries = new List<ISeries>();

            // 1. Agrupar datos por Sucursal
            var datosPorSucursal = datos.GroupBy(x => x.Sucursal).OrderBy(g => g.Key).ToList();

            foreach (var grupoSucursal in datosPorSucursal)
            {
                var nombreSucursal = grupoSucursal.Key;
                var valores = new List<double?>(); // Usamos nullable para huecos

                // 2. Alinear datos con el Eje X (Meses)
                foreach (var mes in mesesEjeX)
                {
                    // Buscamos si hubo venta en ese mes para esta sucursal
                    var ventaMes = grupoSucursal.Where(x => x.Periodo == mes).ToList();

                    if (ventaMes.Any())
                    {
                        // Promedio de utilidad de ese mes
                        double utilidad = (double)ventaMes.Average(x => x.PorcentajeUtilidad);
                        valores.Add(utilidad * 100); // Convertir a escala 0-100
                    }
                    else
                    {
                        // Si no hubo venta, agregamos null para que la línea se corte o no dibuje punto
                        valores.Add(null);
                    }
                }

                // 3. Crear Serie para la Sucursal
                listaSeries.Add(new LineSeries<double?>
                {
                    Name = nombreSucursal,
                    Values = valores,
                    LineSmoothness = 0, // Líneas rectas para mayor precisión
                    GeometrySize = 8,
                    Stroke = new SolidColorPaint { StrokeThickness = 3 }, // Grosor de línea
                    Fill = null, // Sin relleno debajo de la línea para no ensuciar
                    TooltipLabelFormatter = p => $"{p.Context.Series.Name}: {p.Model:N2}%"
                });
            }

            GraficoVentas.Series = listaSeries.ToArray();
            GraficoVentas.XAxes = new Axis[] {
                new Axis {
                    Labels = mesesEjeX,
                    LabelsRotation = 0,
                    TextSize = 12,
                    Name = "Comparativa Mensual"
                }
            };
            GraficoVentas.YAxes = new Axis[] {
                new Axis {
                    Labeler = v => $"{v:N0}%",
                    Name = $"Margen Utilidad: {nombreProducto}"
                }
            };
        }

        // ==========================================
        // MÉTODOS ESTÁNDAR (Lógica base intacta)
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
        // INFRAESTRUCTURA (Helpers y Excel)
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
                    if (esTiempo) listaSeries.Add(new LineSeries<double> { Name = s, Values = v, LineSmoothness = 0, GeometrySize = 10, Stroke = new SolidColorPaint { StrokeThickness = 3 }, Fill = null, TooltipLabelFormatter = tp });
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

        private void LimpiarVersionesAntiguas() { try { string p = Process.GetCurrentProcess().MainModule.FileName + ".old"; if (File.Exists(p)) File.Delete(p); } catch { } }
        private void CargarOpcionesDesglose() { LstDesglose.ItemsSource = new List<string> { "Proveedor", "Familia", "Sucursal", "Año Mes" }; }
        private void ConfigurarCulturaManual() { _culturaCR = (CultureInfo)CultureInfo.CreateSpecificCulture("es-CR").Clone(); _culturaCR.NumberFormat.CurrencySymbol = "₡"; CultureInfo.DefaultThreadCurrentCulture = _culturaCR; CultureInfo.DefaultThreadCurrentUICulture = _culturaCR; }
        private async void BtnActualizar_Click(object sender, RoutedEventArgs e) { MessageBox.Show("Sistema Actualizado."); }

        private void BtnCargarCarpeta_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Title = "Seleccione archivo Excel", Filter = "Excel|*.xlsx;*.xls", CheckFileExists = true };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    string carpeta = Path.GetDirectoryName(dialog.FileName);
                    string[] archivos = Directory.GetFiles(carpeta, "*.xls*");
                    if (archivos.Length == 0) return;
                    TxtEstadoArchivo.Text = $"Procesando {archivos.Length} archivos...";
                    var servicio = new ExcelService();
                    string modo = (CmbTipoNegocio.SelectedItem as ComboBoxItem)?.Content.ToString();
                    _datosGlobales.Clear();
                    int contador = 0; string errores = "";
                    foreach (string archivo in archivos) { try { string sucursal = Path.GetFileNameWithoutExtension(archivo); _datosGlobales.AddRange(servicio.CargarDatos(archivo, modo, sucursal)); contador++; } catch (Exception ex) { errores += $"\n{Path.GetFileName(archivo)}: {ex.Message}"; } }
                    if (_datosGlobales.Any()) { TxtEstadoArchivo.Text = $"Carga OK: {contador} archivos."; TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Green; InicializarFiltros(); AplicarFiltros(); if (!string.IsNullOrEmpty(errores)) MessageBox.Show($"Errores:{errores}"); }
                    else { TxtEstadoArchivo.Text = "No se encontraron datos."; TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Red; }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }

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

    public class ExcelService
    {
        public List<VentaItem> CargarDatos(string ruta, string modo, string nombreSucursal)
        {
            var lista = new List<VentaItem>();
            using (var workbook = new XLWorkbook(ruta))
            {
                var ws = workbook.Worksheet(1);
                IXLRow encabezadoRow = null;
                int colFecha = -1, colCodigo = -1, colDesc = -1, colProv = -1, colFam = -1, colTotal = -1, colUtil = -1;
                foreach (var row in ws.RowsUsed().Take(20))
                {
                    colFecha = -1; colCodigo = -1; colDesc = -1; colProv = -1; colFam = -1; colTotal = -1; colUtil = -1;
                    foreach (var cell in row.CellsUsed())
                    {
                        string val = cell.GetString().ToLower().Trim();
                        if (val.Contains("año") || val == "mes") colFecha = cell.Address.ColumnNumber;
                        else if (val == "artículo" || val == "articulo") colCodigo = cell.Address.ColumnNumber;
                        else if (val.Contains("desc") || val.Contains("nombre") || val.Contains("descripcion")) colDesc = cell.Address.ColumnNumber;
                        else if (val.Contains("proveedor")) colProv = cell.Address.ColumnNumber;
                        else if (val.Contains("familia")) colFam = cell.Address.ColumnNumber;
                        else if (val == "total" || val == "total venta") colTotal = cell.Address.ColumnNumber;
                        else if (val.Contains("utilidad") || val.Contains("%")) colUtil = cell.Address.ColumnNumber;
                    }
                    if (colTotal != -1 && colFam != -1) { encabezadoRow = row; break; }
                }
                if (encabezadoRow == null) return lista;
                bool esMinimarket = (colCodigo != -1); if (modo != null && modo.Contains("Minimarket")) esMinimarket = true; else if (modo != null && modo.Contains("Souvenir")) esMinimarket = false;
                string ultPeriodo = "", ultProv = "General", ultFam = "General";
                foreach (var row in ws.RowsUsed().Where(r => r.RowNumber() > encabezadoRow.RowNumber()))
                {
                    try
                    {
                        if (colFecha != -1 && !row.Cell(colFecha).IsEmpty()) { var c = row.Cell(colFecha); ultPeriodo = c.DataType == XLDataType.DateTime ? c.GetDateTime().ToString("yyyy-MM") : c.GetString(); }
                        if (colProv != -1 && !row.Cell(colProv).IsEmpty()) { string p = row.Cell(colProv).GetString(); if (!p.ToLower().Contains("total")) ultProv = p; }
                        if (colFam != -1 && !row.Cell(colFam).IsEmpty()) ultFam = row.Cell(colFam).GetString();
                        if (esMinimarket && colCodigo != -1 && row.Cell(colCodigo).IsEmpty()) continue;
                        if (!esMinimarket && colFam != -1 && row.Cell(colFam).IsEmpty()) continue;
                        if (colTotal == -1) continue; decimal total = 0; var cTotal = row.Cell(colTotal); if (cTotal.DataType == XLDataType.Number) total = (decimal)cTotal.GetDouble(); else ParseDecimalFlexible(cTotal.GetString(), out total); if (total == 0) continue;
                        decimal utilidad = 0; if (colUtil != -1) { var cUtil = row.Cell(colUtil); if (!cUtil.IsEmpty()) { if (cUtil.DataType == XLDataType.Number) utilidad = (decimal)cUtil.GetDouble(); else ParseDecimalFlexible(cUtil.GetString(), out utilidad); } }
                        string nombreReal = "Sin Nombre"; if (esMinimarket && colDesc != -1 && !row.Cell(colDesc).IsEmpty()) nombreReal = row.Cell(colDesc).GetString(); else if (!esMinimarket) nombreReal = ultFam;
                        if (!string.IsNullOrEmpty(ultPeriodo) && !ultPeriodo.ToLower().Contains("año")) { lista.Add(new VentaItem { Sucursal = nombreSucursal, Periodo = ultPeriodo, ArticuloCodigo = (colCodigo != -1) ? row.Cell(colCodigo).GetString() : "", ArticuloNombre = nombreReal, Proveedor = ultProv, Familia = ultFam, TotalVenta = total, PorcentajeUtilidad = utilidad }); }
                    }
                    catch { continue; }
                }
            }
            return lista;
        }
        private void ParseDecimalFlexible(string t, out decimal r) { r = 0; if (string.IsNullOrWhiteSpace(t)) return; string l = t.Replace("%", "").Replace("$", "").Replace("₡", "").Trim(); if (decimal.TryParse(l, NumberStyles.Any, CultureInfo.InvariantCulture, out r)) return; var cr = CultureInfo.GetCultureInfo("es-CR"); if (decimal.TryParse(l, NumberStyles.Any, cr, out r)) return; }
    }
}