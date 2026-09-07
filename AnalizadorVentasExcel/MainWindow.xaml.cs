using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using AnalizadorVentasExcel.Modelos;
using AnalizadorVentasExcel.Servicios;
using LiveChartsCore;
using LiveChartsCore.Kernel;
using LiveChartsCore.Measure;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using Microsoft.Win32;
using SkiaSharp;

namespace AnalizadorVentasExcel
{
    public partial class MainWindow : Window
    {
        public const string VersionActual = "1.8.0";

        private AnalisisService? _motor;
        private bool _cargandoFiltros;
        private bool _modoExploracion;
        private bool _ocupado;
        private CultureInfo _culturaCR = CultureInfo.InvariantCulture;

        private List<AnalisisService.ProductoExplorado> _productosExplorados = new();

        /// <summary>Ventana del sistema de comparativa de precios; se reutiliza si sigue abierta.</summary>
        private VentanaComparativa? _ventanaComparativa;

        // Colecciones de los checklists. La selección vive aquí (en OpcionFiltro), no en
        // ListBox.SelectedItems, para que el buscador pueda ocultar elementos sin desmarcarlos.
        private readonly List<OpcionFiltro> _opSucursales = new();
        private readonly List<OpcionFiltro> _opPeriodos = new();
        private readonly List<OpcionFiltro> _opProveedores = new();
        private readonly List<OpcionFiltro> _opFamilias = new();
        private readonly List<OpcionFiltro> _opDesglose = new();

        /// <summary>La lista de familias depende de proveedores y sucursales; se reconstruye
        /// dentro del recálculo agrupado, no una vez por casilla marcada.</summary>
        private bool _familiasDesactualizadas;

        /// <summary>
        /// Los checklists disparan SelectionChanged una vez por elemento; con "Seleccionar
        /// todas" eso lanzaba un recálculo completo por cada sucursal/periodo/proveedor.
        /// El temporizador agrupa la ráfaga en un único recálculo.
        /// </summary>
        private readonly DispatcherTimer _temporizadorFiltros;

        public MainWindow()
        {
            InitializeComponent();
            LimpiarVersionesAntiguas();
            ConfigurarCulturaManual();
            CargarOpcionesDesglose();

            _temporizadorFiltros = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(120)
            };
            _temporizadorFiltros.Tick += (_, _) =>
            {
                _temporizadorFiltros.Stop();
                if (_familiasDesactualizadas)
                {
                    _familiasDesactualizadas = false;
                    ActualizarChecklistFamilias();
                }
                AplicarFiltros();
            };

            TxtVersion.Text = $"v{VersionActual}";
            Title = $"Analizador Corporativo v{VersionActual} | Desarrollado por Mateo Sanabria";

            GridResultados.SelectionChanged += GridResultados_SelectionChanged;
        }

        // ==========================================
        // CARGA
        // ==========================================
        private async void BtnCargarCarpeta_Click(object sender, RoutedEventArgs e)
        {
            if (_ocupado) return;

            var dialog = new OpenFileDialog
            {
                Title = "Seleccione un archivo Excel de la carpeta a analizar",
                Filter = "Excel|*.xlsx;*.xls",
                CheckFileExists = true
            };
            if (dialog.ShowDialog() != true) return;

            string? carpeta = Path.GetDirectoryName(dialog.FileName);
            if (string.IsNullOrEmpty(carpeta)) return;

            string[] archivos;
            try
            {
                archivos = Directory.GetFiles(carpeta, "*.xls*")
                                    .Where(a => !Path.GetFileName(a).StartsWith("~$", StringComparison.Ordinal))
                                    .OrderBy(a => a, StringComparer.CurrentCulture)
                                    .ToArray();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se pudo leer la carpeta:\n{ex.Message}", "Error");
                return;
            }

            if (archivos.Length == 0)
            {
                TxtEstadoArchivo.Text = "La carpeta no contiene archivos Excel.";
                TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Red;
                return;
            }

            // Se descarta el conjunto anterior antes de leer el nuevo, para no tener dos
            // cargas completas vivas a la vez.
            _motor = null;
            _modoExploracion = false;
            _productosExplorados = new List<AnalisisService.ProductoExplorado>();
            _temporizadorFiltros.Stop();

            var cronometro = Stopwatch.StartNew();
            EstablecerOcupado(true, $"Leyendo {archivos.Length} archivos...");

            try
            {
                var progreso = new Progress<string>(t => TxtEstadoArchivo.Text = t);
                string? modo = (CmbTipoNegocio.SelectedItem as ComboBoxItem)?.Content?.ToString();

                // La lectura corre fuera del hilo de UI: la ventana sigue respondiendo.
                var resultado = await new ExcelService().CargarCarpetaAsync(archivos, modo, progreso);
                cronometro.Stop();

                if (resultado.Datos.Count == 0)
                {
                    _motor = null;
                    TxtEstadoArchivo.Text = "No se encontraron datos.";
                    TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Red;
                    LimpiarVista();
                }
                else
                {
                    _motor = new AnalisisService(resultado.Datos);
                    TxtEstadoArchivo.Text =
                        $"Carga OK: {resultado.ArchivosLeidos} archivos, " +
                        $"{resultado.Datos.Count.ToString("N0", _culturaCR)} filas en " +
                        $"{cronometro.Elapsed.TotalSeconds:N1} s.";
                    TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Green;

                    InicializarFiltros(resultado.Datos);
                    AplicarFiltros();
                }

                if (resultado.Errores.Count > 0)
                    MessageBox.Show("Archivos con problemas:\n" + string.Join("\n", resultado.Errores), "Aviso");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error al cargar");
                TxtEstadoArchivo.Text = "Error en la carga.";
                TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.Red;
            }
            finally
            {
                EstablecerOcupado(false, null);
            }
        }

        private void EstablecerOcupado(bool ocupado, string? mensaje)
        {
            _ocupado = ocupado;
            BarraProgreso.Visibility = ocupado ? Visibility.Visible : Visibility.Collapsed;
            BtnCargarCarpeta.IsEnabled = !ocupado;
            BtnAuditar.IsEnabled = !ocupado;
            Cursor = ocupado ? System.Windows.Input.Cursors.Wait : null;
            if (mensaje != null)
            {
                TxtEstadoArchivo.Text = mensaje;
                TxtEstadoArchivo.Foreground = System.Windows.Media.Brushes.DimGray;
            }
        }

        private void LimpiarVista()
        {
            GridResultados.ItemsSource = null;
            GraficoVentas.Series = Array.Empty<ISeries>();
            foreach (var lb in new[] { LstFiltroSucursal, LstFiltroFecha, LstFiltroProveedor, LstFiltroFamilia })
                lb.ItemsSource = null;
        }

        // ==========================================
        // ANÁLISIS PRINCIPAL
        // ==========================================
        private void AplicarFiltros()
        {
            _modoExploracion = false;
            if (_motor == null || GridResultados == null || CmbAgrupacion == null) return;

            ColumnaParticipacion.Header = "% Part.";

            var ejeX = ConjuntoDatos.DesdeTexto((CmbAgrupacion.SelectedItem as ComboBoxItem)?.Content?.ToString());
            string? textoOperacion = (CmbOperacion.SelectedItem as ComboBoxItem)?.Content?.ToString();
            if (ejeX == null || textoOperacion == null) return;

            var sucursales = ObtenerSeleccionados(_opSucursales);
            if (sucursales.Count == 0) { GridResultados.ItemsSource = null; GraficoVentas.Series = Array.Empty<ISeries>(); return; }

            var peticion = new PeticionAnalisis
            {
                EjeX = ejeX.Value,
                Desglose = ObtenerSeleccionados(_opDesglose)
                    .Select(ConjuntoDatos.DesdeTexto)
                    .Where(d => d != null).Select(d => d!.Value).ToList(),
                Operacion = textoOperacion.Contains("Suma") ? Operacion.Suma
                          : textoOperacion.Contains("Promedio") ? Operacion.Promedio
                          : Operacion.Conteo,
                TextoOperacion = textoOperacion,
                Sucursales = sucursales,
                Periodos = ObtenerSeleccionados(_opPeriodos),
                Proveedores = ObtenerSeleccionados(_opProveedores),
                Familias = ObtenerSeleccionados(_opFamilias)
            };

            var resultado = _motor.Analizar(peticion);

            GridResultados.ItemsSource = resultado.Tabla;
            ColumnaValor.Header = textoOperacion;

            string nombreEje = (CmbAgrupacion.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
            TxtTituloReporte.Text = peticion.Desglose.Count > 0
                ? $"Análisis: {nombreEje} vs Series"
                : $"Total por {nombreEje}";
            TxtSubtitulo.Text = $"{resultado.FilasFiltradas.ToString("N0", _culturaCR)} registros filtrados.";

            DibujarGrafico(resultado);
        }

        private void DibujarGrafico(ResultadoAnalisis resultado)
        {
            GraficoVentas.TooltipFindingStrategy = TooltipFindingStrategy.CompareOnlyX;

            // Devolver null oculta la entrada del tooltip para los valores ~0, igual que antes.
            Func<ChartPoint, string> etiqueta = punto =>
            {
                double v = punto.Coordinate.PrimaryValue;
                return Math.Abs(v) < 0.01 ? null! : $"{punto.Context.Series.Name}: {v.ToString("N0", _culturaCR)}";
            };

            var series = new List<ISeries>(resultado.Series.Count);
            for (int i = 0; i < resultado.Series.Count; i++)
            {
                var s = resultado.Series[i];
                series.Add(SerieLinea(s.Nombre, s.Valores, ColorSerie(i), etiqueta));
            }

            GraficoVentas.Series = series;
            GraficoVentas.XAxes = new[] { new Axis { Labels = resultado.EtiquetasX, LabelsRotation = 25, TextSize = 11 } };
            GraficoVentas.YAxes = new[] { new Axis { Labeler = v => v.ToString("N0", _culturaCR) } };
        }

        /// <summary>
        /// Paleta fija para las series. Antes el trazo se creaba con
        /// <c>new SolidColorPaint { StrokeThickness = 3 }</c>, sin color: el valor por
        /// defecto de SKColor es transparente, así que la línea no se pintaba y sólo
        /// quedaban los puntos (de ahí el aspecto de diagrama de dispersión).
        /// </summary>
        private static readonly SKColor[] Paleta =
        {
            new SKColor(0x29, 0x80, 0xB9), // azul
            new SKColor(0xE7, 0x4C, 0x3C), // rojo
            new SKColor(0x27, 0xAE, 0x60), // verde
            new SKColor(0xE6, 0x7E, 0x22), // naranja
            new SKColor(0x8E, 0x44, 0xAD), // morado
            new SKColor(0x16, 0xA0, 0x85), // turquesa
            new SKColor(0xD3, 0x54, 0x00), // teja
            new SKColor(0x2C, 0x3E, 0x50), // azul oscuro
            new SKColor(0xC0, 0x39, 0x2B), // rojo oscuro
            new SKColor(0x7F, 0x8C, 0x8D)  // gris
        };

        private static SKColor ColorSerie(int indice) => Paleta[indice % Paleta.Length];

        /// <summary>Serie de línea con trazo y puntos visibles.</summary>
        private static LineSeries<double> SerieLinea(string nombre, double[] valores, SKColor color,
                                                     Func<ChartPoint, string> etiqueta)
            => new LineSeries<double>
            {
                Name = nombre,
                Values = valores,
                LineSmoothness = 0,
                GeometrySize = 9,
                Stroke = new SolidColorPaint(color) { StrokeThickness = 3 },
                GeometryFill = new SolidColorPaint(color),
                GeometryStroke = new SolidColorPaint(SKColors.White) { StrokeThickness = 2 },
                Fill = null, // sin área bajo la línea: con varias series se taparían entre sí
                YToolTipLabelFormatter = etiqueta
            };

        // ==========================================
        // EXPLORADOR DE PRODUCTOS
        // ==========================================
        private void BtnAuditar_Click(object sender, RoutedEventArgs e)
        {
            if (_motor == null) { MessageBox.Show("Primero cargue datos.", "Sin Datos"); return; }

            // Un recálculo pendiente por el temporizador de filtros sobrescribiría el
            // explorador justo después de abrirlo.
            _temporizadorFiltros.Stop();

            var periodos = ObtenerSeleccionados(_opPeriodos);
            var sucursales = ObtenerSeleccionados(_opSucursales);
            if (periodos.Count == 0) periodos = _motor.Datos.Periodos.Valores.ToList();
            if (sucursales.Count == 0) sucursales = _motor.Datos.Sucursales.Valores.ToList();

            _productosExplorados = _motor.Explorar(periodos, sucursales);
            if (_productosExplorados.Count == 0)
            {
                MessageBox.Show("No hay datos para los filtros seleccionados.");
                return;
            }

            _modoExploracion = true;
            GridResultados.ItemsSource = _productosExplorados.Select(p => p.Fila).ToList();

            TxtTituloReporte.Text = "📦 Explorador de Productos";
            TxtSubtitulo.Text = $"{_productosExplorados.Count.ToString("N0", _culturaCR)} productos únicos. " +
                                "Seleccione uno para comparar sucursales mes a mes.";
            ColumnaValor.Header = "Venta Total";
            ColumnaParticipacion.Header = "Disponibilidad";
            GraficoVentas.Series = Array.Empty<ISeries>();
        }

        private void GridResultados_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_modoExploracion) return;

            int indice = GridResultados.SelectedIndex;
            if (indice < 0 || indice >= _productosExplorados.Count) return;

            var producto = _productosExplorados[indice];
            var periodos = ObtenerSeleccionados(_opPeriodos);
            if (periodos.Count == 0) periodos = _motor!.Datos.Periodos.Valores.ToList();
            periodos = periodos.OrderBy(p => p, StringComparer.CurrentCulture).ToList();

            var (sucursales, valores) = _motor!.MargenPorSucursal(producto.NormalizadoId, periodos);
            DibujarGraficoComparativo(sucursales, valores, periodos, producto.Fila.Etiqueta);
        }

        private void DibujarGraficoComparativo(List<string> sucursales, List<double?[]> valores,
                                               List<string> meses, string nombreProducto)
        {
            GraficoVentas.TooltipFindingStrategy = TooltipFindingStrategy.CompareOnlyX;

            var series = new List<ISeries>(sucursales.Count);
            for (int i = 0; i < sucursales.Count; i++)
            {
                var color = ColorSerie(i);
                series.Add(new LineSeries<double?>
                {
                    Name = sucursales[i],
                    Values = valores[i],
                    LineSmoothness = 0,
                    GeometrySize = 9,
                    Stroke = new SolidColorPaint(color) { StrokeThickness = 3 },
                    GeometryFill = new SolidColorPaint(color),
                    GeometryStroke = new SolidColorPaint(SKColors.White) { StrokeThickness = 2 },
                    Fill = null,
                    YToolTipLabelFormatter = p => $"{p.Context.Series.Name}: {p.Coordinate.PrimaryValue:N2}%"
                });
            }

            GraficoVentas.Series = series;
            GraficoVentas.XAxes = new[] { new Axis { Labels = meses, LabelsRotation = 0, TextSize = 12, Name = "Comparativa Mensual" } };
            // El nombre del producto va en el subtítulo: puesto en el eje se recortaba.
            GraficoVentas.YAxes = new[] { new Axis { Labeler = v => $"{v:N0}%", Name = "Margen de Utilidad" } };

            TxtSubtitulo.Text = sucursales.Count > 0
                ? $"{nombreProducto} — margen mes a mes en {sucursales.Count} sucursal(es)."
                : $"{nombreProducto} — sin ventas en los periodos seleccionados.";
        }

        // ==========================================
        // FILTROS
        // ==========================================
        private void InicializarFiltros(ConjuntoDatos datos)
        {
            _cargandoFiltros = true;
            try
            {
                Rellenar(_opSucursales, LstFiltroSucursal,
                    datos.Sucursales.Valores.OrderBy(x => x, StringComparer.CurrentCulture), true,
                    marcaFamilias: true);

                Rellenar(_opPeriodos, LstFiltroFecha,
                    datos.Periodos.Valores.OrderByDescending(x => x, StringComparer.CurrentCulture), true);

                Rellenar(_opProveedores, LstFiltroProveedor,
                    datos.Proveedores.Valores.OrderBy(x => x, StringComparer.CurrentCulture), true,
                    marcaFamilias: true);
                AplicarBusqueda(LstFiltroProveedor, TxtBuscarProveedor.Text);

                ActualizarChecklistFamilias();
            }
            finally { _cargandoFiltros = false; }
        }

        /// <summary>
        /// Reconstruye un checklist. Se desuscribe de los elementos anteriores para no
        /// dejar manejadores colgando cada vez que se recarga una carpeta.
        /// </summary>
        private void Rellenar(List<OpcionFiltro> destino, ListBox lista, IEnumerable<string> valores,
                              bool marcados, bool marcaFamilias = false)
        {
            foreach (var vieja in destino)
            {
                vieja.PropertyChanged -= OpcionCambiada;
                vieja.PropertyChanged -= OpcionCambiadaConFamilias;
            }
            destino.Clear();

            foreach (var v in valores)
            {
                var op = new OpcionFiltro(v, marcados);
                op.PropertyChanged += marcaFamilias ? OpcionCambiadaConFamilias : OpcionCambiada;
                destino.Add(op);
            }

            lista.ItemsSource = null;
            lista.ItemsSource = destino;
        }

        private void OpcionCambiada(object? s, System.ComponentModel.PropertyChangedEventArgs e)
            => ProgramarRecalculo();

        private void OpcionCambiadaConFamilias(object? s, System.ComponentModel.PropertyChangedEventArgs e)
        {
            _familiasDesactualizadas = true;
            ProgramarRecalculo();
        }

        /// <summary>
        /// Las familias visibles dependen de los proveedores y sucursales elegidos.
        /// Se resuelve con máscaras sobre ids en una sola pasada.
        /// </summary>
        private void ActualizarChecklistFamilias()
        {
            if (_motor == null) return;
            var datos = _motor.Datos;

            var proveedores = new bool[datos.Proveedores.Count];
            foreach (var p in ObtenerSeleccionados(_opProveedores))
            {
                int id = datos.Proveedores.IdExistente(p);
                if (id >= 0) proveedores[id] = true;
            }

            var sucursales = new bool[datos.Sucursales.Count];
            foreach (var s in ObtenerSeleccionados(_opSucursales))
            {
                int id = datos.Sucursales.IdExistente(s);
                if (id >= 0) sucursales[id] = true;
            }

            var vistas = new bool[datos.Familias.Count];
            foreach (ref readonly var f in datos.Filas.AsSpan())
                if (proveedores[f.ProveedorId] && sucursales[f.SucursalId]) vistas[f.FamiliaId] = true;

            var lista = new List<string>();
            for (int i = 0; i < vistas.Length; i++) if (vistas[i]) lista.Add(datos.Familias[i]);
            lista.Sort(StringComparer.CurrentCulture);

            bool previo = _cargandoFiltros;
            _cargandoFiltros = true;
            try
            {
                Rellenar(_opFamilias, LstFiltroFamilia, lista, true);
                AplicarBusqueda(LstFiltroFamilia, TxtBuscarFamilia.Text);
            }
            finally { _cargandoFiltros = previo; }
        }

        private static List<string> ObtenerSeleccionados(List<OpcionFiltro> opciones)
        {
            var lista = new List<string>(opciones.Count);
            foreach (var o in opciones) if (o.Seleccionado) lista.Add(o.Nombre);
            return lista;
        }

        private void ProgramarRecalculo()
        {
            if (_cargandoFiltros || _motor == null) return;
            _temporizadorFiltros.Stop();
            _temporizadorFiltros.Start();
        }

        // ==========================================
        // BUSCADORES
        // ==========================================

        /// <summary>
        /// Oculta de la vista lo que no coincide, sin tocar el estado marcado: la selección
        /// está en OpcionFiltro y el filtro sólo afecta a la CollectionView.
        /// </summary>
        private static void AplicarBusqueda(ListBox lista, string? texto)
        {
            var vista = CollectionViewSource.GetDefaultView(lista.ItemsSource);
            if (vista == null) return;

            texto = texto?.Trim();
            if (string.IsNullOrEmpty(texto)) vista.Filter = null;
            else vista.Filter = o => o is OpcionFiltro op &&
                                     op.Nombre.Contains(texto, StringComparison.CurrentCultureIgnoreCase);
        }

        private void TxtBuscarProveedor_TextChanged(object sender, TextChangedEventArgs e)
        {
            PistaProveedor.Visibility = TxtBuscarProveedor.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
            AplicarBusqueda(LstFiltroProveedor, TxtBuscarProveedor.Text);
        }

        private void TxtBuscarFamilia_TextChanged(object sender, TextChangedEventArgs e)
        {
            PistaFamilia.Visibility = TxtBuscarFamilia.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
            AplicarBusqueda(LstFiltroFamilia, TxtBuscarFamilia.Text);
        }

        /// <summary>
        /// Marca o desmarca sólo lo que está visible en la lista. Con una búsqueda activa
        /// eso significa "sólo los resultados de la búsqueda", que es lo que hace útil la
        /// combinación buscar + Todas / Ninguna.
        /// </summary>
        private void MarcarVisibles(ListBox lista, bool marcado)
        {
            var vista = CollectionViewSource.GetDefaultView(lista.ItemsSource);
            if (vista == null) return;

            bool alguno = false;
            foreach (var o in vista)
            {
                if (o is not OpcionFiltro op || op.Seleccionado == marcado) continue;
                op.EstablecerSilencioso(marcado);
                op.NotificarSeleccion();   // refresca la casilla, sin recalcular por cada una
                alguno = true;
            }

            if (!alguno) return;
            if (lista == LstFiltroProveedor || lista == LstFiltroSucursal) _familiasDesactualizadas = true;
            ProgramarRecalculo();
        }

        private void BtnSelectAllSucursal_Click(object s, RoutedEventArgs e) => MarcarVisibles(LstFiltroSucursal, true);
        private void BtnSelectAllFecha_Click(object s, RoutedEventArgs e) => MarcarVisibles(LstFiltroFecha, true);
        private void BtnSelectAllProv_Click(object s, RoutedEventArgs e) => MarcarVisibles(LstFiltroProveedor, true);
        private void BtnSelectAllFam_Click(object s, RoutedEventArgs e) => MarcarVisibles(LstFiltroFamilia, true);
        private void BtnNingunoProv_Click(object s, RoutedEventArgs e) => MarcarVisibles(LstFiltroProveedor, false);
        private void BtnNingunaFam_Click(object s, RoutedEventArgs e) => MarcarVisibles(LstFiltroFamilia, false);

        private void AplicarFiltros_Event(object s, SelectionChangedEventArgs e) => ProgramarRecalculo();

        private void BtnAyuda_Click(object sender, RoutedEventArgs e)
            => new VentanaGuia { Owner = this }.ShowDialog();

        /// <summary>
        /// La comparativa de precios es un sistema aparte: lee otros archivos (listas de
        /// precios, sin ventas ni periodos) y no comparte los datos cargados aquí, así que
        /// vive en su propia ventana no modal para poder trabajar con las dos a la vez.
        /// </summary>
        private void BtnComparativa_Click(object sender, RoutedEventArgs e)
        {
            if (_ventanaComparativa == null || !_ventanaComparativa.IsLoaded)
            {
                _ventanaComparativa = new VentanaComparativa { Owner = this };
                _ventanaComparativa.Closed += (_, _) => _ventanaComparativa = null;
                _ventanaComparativa.Show();
            }
            else
            {
                if (_ventanaComparativa.WindowState == WindowState.Minimized)
                    _ventanaComparativa.WindowState = WindowState.Normal;
                _ventanaComparativa.Activate();
            }
        }

        // ==========================================
        // VARIOS
        // ==========================================
        private void CargarOpcionesDesglose()
        {
            Rellenar(_opDesglose, LstDesglose, new[] { "Proveedor", "Familia", "Sucursal", "Año Mes" }, marcados: false);
        }

        private void ConfigurarCulturaManual()
        {
            _culturaCR = (CultureInfo)CultureInfo.CreateSpecificCulture("es-CR").Clone();
            _culturaCR.NumberFormat.CurrencySymbol = "₡";
            CultureInfo.DefaultThreadCurrentCulture = _culturaCR;
            CultureInfo.DefaultThreadCurrentUICulture = _culturaCR;
            Thread.CurrentThread.CurrentCulture = _culturaCR;
            Thread.CurrentThread.CurrentUICulture = _culturaCR;
        }

        private void BtnActualizar_Click(object sender, RoutedEventArgs e)
            => new VentanaActualizacion { Owner = this }.ShowDialog();

        /// <summary>
        /// Borra el ejecutable anterior que dejó una actualización, y también el temporal
        /// de una descarga que se hubiera interrumpido.
        /// </summary>
        private static void LimpiarVersionesAntiguas()
        {
            ActualizacionService.LimpiarRespaldos();
            try
            {
                string? exe = Environment.ProcessPath;
                if (exe != null && File.Exists(exe + ".nuevo")) File.Delete(exe + ".nuevo");
            }
            catch { /* sin permisos o archivo en uso: no es crítico */ }
        }
    }
}
