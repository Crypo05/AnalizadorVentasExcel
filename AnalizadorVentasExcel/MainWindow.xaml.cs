using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Diagnostics; // Para abrir enlaces web y reiniciar app
using System.Net.Http;    // Para verificar actualizaciones
using Microsoft.Win32;
using ClosedXML.Excel;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using SkiaSharp;
using LiveChartsCore.Measure;
using LiveChartsCore.Kernel; // ¡IMPORTANTE! Necesario para ChartPoint

namespace AnalizadorVentasExcel
{
    public partial class MainWindow : Window
    {
        // ==========================================
        // CONFIGURACIÓN DE LA APLICACIÓN
        // ==========================================
        private const string VersionActual = "1.0.0";
        // URL RAW del archivo de texto con la versión (ej: 1.0.1)
        private const string UrlVersionRemota = "https://raw.githubusercontent.com/TU_USUARIO/TU_REPO/main/version.txt";
        // URL de descarga del ejecutable nuevo
        private const string UrlDescarga = "https://github.com/TU_USUARIO/TU_REPO/raw/main/AnalizadorVentasExcel.exe";

        private List<VentaItem> _datosGlobales = new List<VentaItem>();
        private bool _cargandoFiltros = false;
        private CultureInfo _culturaCR;

        public MainWindow()
        {
            InitializeComponent();

            // 1. Limpieza de versiones antiguas tras actualización
            LimpiarVersionesAntiguas();

            ConfigurarCulturaManual();
            CargarOpcionesDesglose();

            // 2. Título con Créditos
            this.Title = $"Analizador Corporativo v{VersionActual} | Desarrollado por Mateo Sanabria Murcia";
        }

        private void LimpiarVersionesAntiguas()
        {
            try
            {
                string currentPath = Process.GetCurrentProcess().MainModule.FileName;
                string oldPath = currentPath + ".old";
                if (File.Exists(oldPath)) File.Delete(oldPath);
            }
            catch { /* Ignorar si no se puede borrar */ }
        }

        private void CargarOpcionesDesglose()
        {
            var opciones = new List<string> { "Proveedor", "Familia", "Sucursal", "Año Mes" };
            LstDesglose.ItemsSource = opciones;
        }

        private void ConfigurarCulturaManual()
        {
            _culturaCR = (CultureInfo)CultureInfo.CreateSpecificCulture("es-CR").Clone();
            _culturaCR.NumberFormat.CurrencySymbol = "₡";
            _culturaCR.NumberFormat.CurrencyDecimalDigits = 2;
            _culturaCR.NumberFormat.CurrencyGroupSeparator = ",";
            _culturaCR.NumberFormat.CurrencyDecimalSeparator = ".";

            CultureInfo.DefaultThreadCurrentCulture = _culturaCR;
            CultureInfo.DefaultThreadCurrentUICulture = _culturaCR;
        }

        // ==========================================
        // 1. SISTEMA DE ACTUALIZACIÓN
        // ==========================================
        private async void BtnActualizar_Click(object sender, RoutedEventArgs e)
        {
            BtnActualizar.IsEnabled = false;
            BtnActualizar.Content = "Verificando...";

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Añadimos timestamp para evitar caché
                    string versionRemota = await client.GetStringAsync(UrlVersionRemota + $"?t={DateTime.Now.Ticks}");
                    versionRemota = versionRemota.Trim();

                    // SI USAS LOCAL/PRUEBAS, COMENTA LO DE ARRIBA Y USA ESTO:
                    // string versionRemota = "1.0.0"; 

                    if (versionRemota != VersionActual)
                    {
                        var result = MessageBox.Show($"Nueva versión {versionRemota} disponible.\n¿Desea actualizar ahora?",
                            "Actualización", MessageBoxButton.YesNo, MessageBoxImage.Question);

                        if (result == MessageBoxResult.Yes)
                        {
                            BtnActualizar.Content = "Descargando...";

                            // Rutas para el SWAP de archivos
                            string rutaActual = Process.GetCurrentProcess().MainModule.FileName;
                            string rutaNueva = rutaActual + ".new";
                            string rutaVieja = rutaActual + ".old";

                            // Descargar
                            byte[] fileBytes = await client.GetByteArrayAsync(UrlDescarga);
                            File.WriteAllBytes(rutaNueva, fileBytes);

                            // Renombrar
                            if (File.Exists(rutaVieja)) File.Delete(rutaVieja);
                            File.Move(rutaActual, rutaVieja);
                            File.Move(rutaNueva, rutaActual);

                            MessageBox.Show("Actualización completada. La aplicación se reiniciará.", "Éxito");

                            // Reiniciar
                            Process.Start(rutaActual);
                            Application.Current.Shutdown();
                        }
                    }
                    else
                    {
                        MessageBox.Show("El sistema está actualizado.", "Estado", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al actualizar: {ex.Message}\nIntente más tarde.", "Error de Conexión");
            }
            finally
            {
                BtnActualizar.IsEnabled = true;
                BtnActualizar.Content = "🔄 Buscar Actualización";
            }
        }

        // ==========================================
        // 2. CARGA MASIVA (CARPETAS)
        // ==========================================
        private void BtnCargarCarpeta_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Seleccione UN archivo Excel dentro de la carpeta a procesar",
                Filter = "Excel Files|*.xlsx;*.xls",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    string carpeta = Path.GetDirectoryName(dialog.FileName);
                    string[] archivos = Directory.GetFiles(carpeta, "*.xls*");

                    if (archivos.Length == 0) return;

                    TxtEstadoArchivo.Text = $"Analizando {archivos.Length} archivos...";

                    string modo = (CmbTipoNegocio.SelectedItem as ComboBoxItem)?.Content.ToString();
                    var servicio = new ExcelService();
                    _datosGlobales.Clear();

                    int procesados = 0;
                    string errores = "";

                    foreach (string archivo in archivos)
                    {
                        try
                        {
                            string nombreSucursal = Path.GetFileNameWithoutExtension(archivo);
                            var datos = servicio.CargarDatos(archivo, modo, nombreSucursal);
                            _datosGlobales.AddRange(datos);
                            procesados++;
                        }
                        catch (Exception innerEx) { errores += $"\n{Path.GetFileName(archivo)}: {innerEx.Message}"; }
                    }

                    if (_datosGlobales.Any())
                    {
                        TxtEstadoArchivo.Text = $"Carga Exitosa: {procesados} archivos.";
                        TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Green;
                        InicializarFiltros();
                        AplicarFiltros();
                        if (!string.IsNullOrEmpty(errores)) MessageBox.Show($"Advertencia en algunos archivos:{errores}");
                    }
                    else
                    {
                        TxtEstadoArchivo.Text = "No se encontraron datos válidos.";
                        TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Red;
                    }
                }
                catch (Exception ex) { MessageBox.Show($"Error Fatal: {ex.Message}"); }
            }
        }

        // ==========================================
        // 3. GESTIÓN DE FILTROS
        // ==========================================
        private void InicializarFiltros()
        {
            _cargandoFiltros = true;
            LstFiltroSucursal.ItemsSource = _datosGlobales.Select(x => x.Sucursal).Distinct().OrderBy(x => x).ToList();
            LstFiltroSucursal.SelectAll();

            LstFiltroFecha.ItemsSource = _datosGlobales.Select(x => x.Periodo).Distinct().OrderByDescending(x => x).ToList();
            LstFiltroFecha.SelectAll();

            LstFiltroProveedor.ItemsSource = _datosGlobales.Select(x => x.Proveedor).Distinct().OrderBy(x => x).ToList();
            LstFiltroProveedor.SelectAll();

            ActualizarChecklistFamilias();
            _cargandoFiltros = false;
        }

        private void LstFiltroProveedor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_cargandoFiltros) return;
            ActualizarChecklistFamilias();
            AplicarFiltros();
        }

        private void ActualizarChecklistFamilias()
        {
            if (LstFiltroProveedor == null) return;
            bool prev = _cargandoFiltros;
            _cargandoFiltros = true;

            var provs = ObtenerSeleccionados(LstFiltroProveedor);
            var sucs = ObtenerSeleccionados(LstFiltroSucursal);

            var q = _datosGlobales.AsEnumerable();
            if (provs.Any()) q = q.Where(x => provs.Contains(x.Proveedor));
            if (sucs.Any()) q = q.Where(x => sucs.Contains(x.Sucursal));

            LstFiltroFamilia.ItemsSource = q.Select(x => x.Familia).Distinct().OrderBy(x => x).ToList();
            LstFiltroFamilia.SelectAll();
            _cargandoFiltros = prev;
        }

        private List<string> ObtenerSeleccionados(ListBox lb)
        {
            var l = new List<string>();
            if (lb.SelectedItems == null) return l;
            foreach (var item in lb.SelectedItems)
            {
                if (item is ListBoxItem lbi) l.Add(lbi.Content.ToString());
                else l.Add(item.ToString());
            }
            return l;
        }

        private void BtnSelectAllSucursal_Click(object sender, RoutedEventArgs e) => LstFiltroSucursal.SelectAll();
        private void BtnSelectAllFecha_Click(object sender, RoutedEventArgs e) => LstFiltroFecha.SelectAll();
        private void BtnSelectAllProv_Click(object sender, RoutedEventArgs e) => LstFiltroProveedor.SelectAll();
        private void BtnSelectAllFam_Click(object sender, RoutedEventArgs e) => LstFiltroFamilia.SelectAll();

        private void AplicarFiltros_Event(object sender, RoutedEventArgs e) { if (!_cargandoFiltros) AplicarFiltros(); }
        private void AplicarFiltros_Event(object sender, SelectionChangedEventArgs e) { if (!_cargandoFiltros) AplicarFiltros(); }

        // ==========================================
        // 4. MOTOR DE ANÁLISIS (PIVOT)
        // ==========================================
        private void AplicarFiltros()
        {
            if (GridResultados == null || CmbAgrupacion == null || LstDesglose == null) return;
            if (LstFiltroFecha == null) return;

            // Configuración
            var itemEjeX = CmbAgrupacion.SelectedItem as ComboBoxItem;
            var itemOp = CmbOperacion.SelectedItem as ComboBoxItem;
            if (itemEjeX == null || itemOp == null) return;

            string ejeX = itemEjeX.Content.ToString();
            string operacion = itemOp.Content.ToString();

            var dimensionesSerie = ObtenerSeleccionados(LstDesglose);
            bool hayDesglose = dimensionesSerie.Any();

            // Filtros
            var sSucs = ObtenerSeleccionados(LstFiltroSucursal);
            var sFechas = ObtenerSeleccionados(LstFiltroFecha);
            var sProvs = ObtenerSeleccionados(LstFiltroProveedor);
            var sFams = ObtenerSeleccionados(LstFiltroFamilia);

            if (!sSucs.Any()) { GridResultados.ItemsSource = null; return; }

            // Datos
            var datos = _datosGlobales.Where(x =>
                sSucs.Contains(x.Sucursal) && sFechas.Contains(x.Periodo) &&
                sProvs.Contains(x.Proveedor) && sFams.Contains(x.Familia)).ToList();

            double sumaGlobal = datos.Sum(x => (double)x.TotalVenta);

            // Generar Tabla
            List<ResumenDinamico> resumenTabla;

            if (hayDesglose)
            {
                resumenTabla = datos
                    .GroupBy(x => new { KeyX = ObtenerLlaveSimple(x, ejeX), KeySerie = ObtenerLlaveCompuesta(x, dimensionesSerie) })
                    .Select(g => new ResumenDinamico
                    {
                        Etiqueta = g.Key.KeyX,
                        DetalleSecundario = g.Key.KeySerie,
                        ValorNumerico = CalcularValor(g, operacion),
                        TipoFormato = operacion,
                        Participacion = (operacion.Contains("Suma") && sumaGlobal > 0) ? (CalcularValor(g, operacion) / sumaGlobal).ToString("P1", _culturaCR) : "-"
                    })
                    .OrderByDescending(x => x.ValorNumerico)
                    .ToList();
            }
            else
            {
                resumenTabla = datos.GroupBy(x => ObtenerLlaveSimple(x, ejeX))
                    .Select(g => new ResumenDinamico
                    {
                        Etiqueta = g.Key,
                        DetalleSecundario = "Total General",
                        ValorNumerico = CalcularValor(g, operacion),
                        TipoFormato = operacion,
                        Participacion = (operacion.Contains("Suma") && sumaGlobal > 0) ? (CalcularValor(g, operacion) / sumaGlobal).ToString("P1", _culturaCR) : "-"
                    })
                    .OrderByDescending(x => x.ValorNumerico)
                    .ToList();
            }

            if (ejeX == "Año Mes") resumenTabla = resumenTabla.OrderBy(x => x.Etiqueta).ThenByDescending(x => x.ValorNumerico).ToList();

            GridResultados.ItemsSource = resumenTabla;
            if (ColumnaValor != null) ColumnaValor.Header = operacion;

            TxtTituloReporte.Text = hayDesglose ? $"Análisis: {ejeX} vs Series" : $"Total por {ejeX}";
            TxtSubtitulo.Text = $"{datos.Count} registros analizados.";

            ActualizarGraficoMultiNivel(datos, ejeX, dimensionesSerie, operacion);
        }

        private double CalcularValor(IEnumerable<VentaItem> datos, string operacion)
        {
            if (operacion.Contains("Suma")) return (double)datos.Sum(x => x.TotalVenta);
            if (operacion.Contains("Promedio")) return (double)(datos.Any() ? datos.Average(x => x.PorcentajeUtilidad) : 0);
            return datos.Count();
        }

        private string ObtenerLlaveSimple(VentaItem item, string criterio)
        {
            switch (criterio)
            {
                case "Año Mes": return item.Periodo;
                case "Proveedor": return item.Proveedor;
                case "Familia": return item.Familia;
                case "Sucursal": return item.Sucursal;
                case "Articulo": return item.ArticuloNombre;
                default: return "General";
            }
        }

        private string ObtenerLlaveCompuesta(VentaItem item, List<string> dimensiones)
        {
            if (!dimensiones.Any()) return "";
            var partes = new List<string>();
            foreach (var dim in dimensiones) partes.Add(ObtenerLlaveSimple(item, dim));
            return string.Join(" - ", partes);
        }

        // ==========================================
        // 5. GRÁFICOS AVANZADOS (CLEAN & STRICT)
        // ==========================================
        private void ActualizarGraficoMultiNivel(List<VentaItem> datos, string ejeX, List<string> dimensionesSerie, string operacion)
        {
            if (GraficoVentas == null) return;
            GraficoVentas.Series = new ISeries[] { };

            // Estrategia estricta: El tooltip solo muestra lo que está exactamente bajo el mouse en el eje X
            GraficoVentas.TooltipFindingStrategy = TooltipFindingStrategy.CompareOnlyX;

            bool esTiempoEnX = (ejeX == "Año Mes");
            bool hayDesglose = dimensionesSerie.Any();

            // Definir Eje X
            var etiquetasX = datos.Select(x => ObtenerLlaveSimple(x, ejeX)).Distinct().ToList();
            if (esTiempoEnX) etiquetasX = etiquetasX.OrderBy(x => x).ToList();
            else etiquetasX = etiquetasX.OrderByDescending(lbl => CalcularValor(datos.Where(d => ObtenerLlaveSimple(d, ejeX) == lbl), operacion)).Take(20).ToList();

            var listaSeries = new List<ISeries>();

            // --- FORMATEADOR ANTI-CEROS ---
            Func<ChartPoint, string> tooltipCleaner = point => {
                // PrimaryValue obtiene el valor numérico (Y) del punto
                double val = point.PrimaryValue;
                // Si es casi cero, retornamos null (oculta la etiqueta del tooltip)
                if (Math.Abs(val) < 0.01) return null;
                return $"{point.Context.Series.Name}: {val.ToString("N0", _culturaCR)}";
            };
            // ------------------------------

            if (hayDesglose)
            {
                // Top Series Principales
                var topSeries = datos.GroupBy(x => ObtenerLlaveCompuesta(x, dimensionesSerie))
                                     .OrderByDescending(g => CalcularValor(g, operacion))
                                     .Take(10) // Limitamos colores
                                     .Select(g => g.Key).ToList();

                foreach (var nombreSerie in topSeries)
                {
                    var valores = new List<double>();
                    foreach (var xLabel in etiquetasX)
                    {
                        var subSet = datos.Where(d => ObtenerLlaveSimple(d, ejeX) == xLabel &&
                                                      ObtenerLlaveCompuesta(d, dimensionesSerie) == nombreSerie);
                        valores.Add(CalcularValor(subSet, operacion));
                    }

                    if (esTiempoEnX)
                    {
                        // Líneas para Tiempo
                        listaSeries.Add(new LineSeries<double>
                        {
                            Name = nombreSerie,
                            Values = valores,
                            LineSmoothness = 0,
                            GeometrySize = 10,
                            Stroke = new SolidColorPaint { StrokeThickness = 3 },
                            Fill = null,
                            TooltipLabelFormatter = tooltipCleaner
                        });
                    }
                    else
                    {
                        // Barras para Categorías
                        listaSeries.Add(new ColumnSeries<double>
                        {
                            Name = nombreSerie,
                            Values = valores,
                            TooltipLabelFormatter = tooltipCleaner
                        });
                    }
                }
            }
            else
            {
                // Sin desglose (Total)
                var valores = new List<double>();
                foreach (var xLabel in etiquetasX)
                {
                    var subSet = datos.Where(d => ObtenerLlaveSimple(d, ejeX) == xLabel);
                    valores.Add(CalcularValor(subSet, operacion));
                }
                listaSeries.Add(new ColumnSeries<double>
                {
                    Name = "Total",
                    Values = valores,
                    Fill = new SolidColorPaint(SKColors.DarkCyan),
                    TooltipLabelFormatter = tooltipCleaner
                });
            }

            GraficoVentas.Series = listaSeries.ToArray();
            GraficoVentas.XAxes = new Axis[] { new Axis { Labels = etiquetasX, LabelsRotation = 25, TextSize = 11 } };
            GraficoVentas.YAxes = new Axis[] { new Axis { Labeler = v => v.ToString("N0", _culturaCR) } };
        }
    }

    // ==========================================
    // CLASES DE MODELO
    // ==========================================
    public class VentaItem
    {
        public string Sucursal { get; set; }
        public string Periodo { get; set; }
        public string ArticuloCodigo { get; set; }
        public string ArticuloNombre { get; set; }
        public string Proveedor { get; set; }
        public string Familia { get; set; }
        public decimal TotalVenta { get; set; }
        public decimal PorcentajeUtilidad { get; set; }
    }

    public class ResumenDinamico
    {
        public string Etiqueta { get; set; }
        public string DetalleSecundario { get; set; }
        public double ValorNumerico { get; set; }
        public string TipoFormato { get; set; }
        public string Participacion { get; set; }
        public string ValorFormateado
        {
            get
            {
                var cr = CultureInfo.GetCultureInfo("es-CR");
                var nfi = (NumberFormatInfo)cr.NumberFormat.Clone();
                nfi.CurrencySymbol = "₡";
                if (TipoFormato.Contains("Suma")) return ValorNumerico.ToString("C2", nfi);
                if (TipoFormato.Contains("Promedio")) return ValorNumerico.ToString("P2", nfi);
                return ValorNumerico.ToString("N0", nfi);
            }
        }
    }

    // ==========================================
    // SERVICIO EXCEL (ROBUSTO Y HÍBRIDO)
    // ==========================================
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

                // 1. Detección Inteligente de Encabezados
                foreach (var row in ws.RowsUsed().Take(20))
                {
                    colFecha = -1; colCodigo = -1; colDesc = -1; colProv = -1; colFam = -1; colTotal = -1; colUtil = -1;
                    foreach (var cell in row.CellsUsed())
                    {
                        string val = cell.GetString().ToLower().Trim();
                        if (val.Contains("año") || val == "mes") colFecha = cell.Address.ColumnNumber;
                        else if (val == "artículo" || val == "articulo") colCodigo = cell.Address.ColumnNumber;
                        else if (val.Contains("desc")) colDesc = cell.Address.ColumnNumber;
                        else if (val.Contains("proveedor")) colProv = cell.Address.ColumnNumber;
                        else if (val.Contains("familia")) colFam = cell.Address.ColumnNumber;
                        else if (val == "total" || val == "total venta" || val == "total total") colTotal = cell.Address.ColumnNumber;
                        else if (val.Contains("utilidad")) colUtil = cell.Address.ColumnNumber;
                    }
                    if (colTotal != -1 && (colFam != -1 || colProv != -1)) { encabezadoRow = row; break; }
                }

                if (encabezadoRow == null) return lista;

                // 2. Modo Operativo
                bool esMinimarket = (colCodigo != -1);
                if (modo != null && modo.Contains("Minimarket")) esMinimarket = true;
                else if (modo != null && modo.Contains("Souvenir")) esMinimarket = false;

                string ultPeriodo = "", ultProv = "General", ultFam = "General";

                // 3. Lectura de Filas
                foreach (var row in ws.RowsUsed().Where(r => r.RowNumber() > encabezadoRow.RowNumber()))
                {
                    try
                    {
                        // Fill Down (Relleno hacia abajo)
                        if (colFecha != -1 && !row.Cell(colFecha).IsEmpty())
                        {
                            var c = row.Cell(colFecha);
                            ultPeriodo = c.DataType == XLDataType.DateTime ? c.GetDateTime().ToString("yyyy-MM") : c.GetString();
                        }
                        if (colProv != -1 && !row.Cell(colProv).IsEmpty())
                        {
                            string p = row.Cell(colProv).GetString();
                            if (!p.Contains("Total")) ultProv = p;
                        }
                        if (colFam != -1 && !row.Cell(colFam).IsEmpty()) ultFam = row.Cell(colFam).GetString();

                        // Filtros Anti-Duplicados
                        if (esMinimarket)
                        {
                            if (colCodigo != -1 && row.Cell(colCodigo).IsEmpty()) continue;
                        }
                        else
                        {
                            if (colFam != -1 && row.Cell(colFam).IsEmpty()) continue;
                            // En Souvenir, si el proveedor dice "Total" es una fila de suma, saltar
                            if (colProv != -1 && !row.Cell(colProv).IsEmpty() && row.Cell(colProv).GetString().Contains("Total")) continue;
                        }

                        // Lectura de Valores
                        var celdaTotal = row.Cell(colTotal);
                        if (celdaTotal.IsEmpty()) continue;

                        decimal total = 0;
                        if (celdaTotal.DataType == XLDataType.Number) total = (decimal)celdaTotal.GetDouble();
                        else decimal.TryParse(celdaTotal.GetString(), out total);

                        if (total == 0) continue;

                        decimal utilidad = 0;
                        if (colUtil != -1)
                        {
                            var cu = row.Cell(colUtil);
                            if (!cu.IsEmpty())
                            {
                                if (cu.DataType == XLDataType.Number) utilidad = (decimal)cu.GetDouble();
                                else decimal.TryParse(cu.GetString(), out utilidad);
                            }
                        }

                        // Guardado seguro
                        if (!string.IsNullOrEmpty(ultPeriodo) && !ultPeriodo.ToLower().Contains("año"))
                        {
                            lista.Add(new VentaItem
                            {
                                Sucursal = nombreSucursal,
                                Periodo = ultPeriodo,
                                ArticuloCodigo = (colCodigo != -1) ? row.Cell(colCodigo).GetString() : "",
                                ArticuloNombre = (colDesc != -1) ? row.Cell(colDesc).GetString() : (esMinimarket ? "Sin Nombre" : ultFam),
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
            return lista;
        }
    }
}