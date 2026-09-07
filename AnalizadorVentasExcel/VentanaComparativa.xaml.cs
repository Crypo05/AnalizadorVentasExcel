using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Threading;
using AnalizadorVentasExcel.Modelos;
using AnalizadorVentasExcel.Servicios;
using LiveChartsCore;
using LiveChartsCore.Measure;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using Microsoft.Win32;
using SkiaSharp;

namespace AnalizadorVentasExcel
{
    /// <summary>
    /// Sistema secundario: compara el precio de un mismo producto entre sucursales.
    ///
    /// No comparte datos con el analizador de ventas. Lee otros archivos (las listas de
    /// precios que exporta la caja: código, descripción, costo, impuesto, utilidad y precio
    /// IVI) y cruza los productos por código de artículo, que es la única llave fiable:
    /// la descripción del mismo producto cambia de una tienda a otra.
    /// </summary>
    public partial class VentanaComparativa : Window
    {
        private ComparativaService? _motor;
        private bool _ocupado;
        private bool _cargandoFiltros;

        private readonly List<OpcionFiltro> _opSucursales = new();

        /// <summary>Columnas de precio insertadas en la tabla, una por sucursal comparada.</summary>
        private readonly List<DataGridColumn> _columnasSucursal = new();
        private List<string> _sucursalesComparadas = new();

        private List<FilaComparativa> _filas = new();
        private MetricaPrecio _metricaVista = MetricaPrecio.PrecioVenta;

        /// <summary>Índice de la tabla donde empiezan las columnas por sucursal.</summary>
        private const int PrimeraColumnaSucursal = 2;

        /// <summary>
        /// Marcar sucursales o escribir en el buscador dispara un evento por cada cambio;
        /// el temporizador agrupa la ráfaga en un único recálculo, igual que en la ventana
        /// principal.
        /// </summary>
        private readonly DispatcherTimer _temporizador;

        public VentanaComparativa()
        {
            InitializeComponent();

            _temporizador = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(150)
            };
            _temporizador.Tick += (_, _) =>
            {
                _temporizador.Stop();
                Comparar();
            };

            GridComparativa.SelectionChanged += GridComparativa_SelectionChanged;
        }

        // ==========================================
        // CARGA
        // ==========================================
        private async void BtnCargarPrecios_Click(object sender, RoutedEventArgs e)
        {
            if (_ocupado) return;

            var dialog = new OpenFileDialog
            {
                Title = "Seleccione un archivo de la carpeta con las listas de precios",
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
                Estado("La carpeta no contiene archivos Excel.", Brushes.Red);
                return;
            }

            _motor = null;
            _temporizador.Stop();
            LimpiarVista();

            var cronometro = Stopwatch.StartNew();
            EstablecerOcupado(true, $"Leyendo {archivos.Length} archivos...");

            try
            {
                var progreso = new Progress<string>(t => TxtEstadoArchivo.Text = t);
                var resultado = await new PreciosExcelService().CargarCarpetaAsync(archivos, progreso);
                cronometro.Stop();

                if (resultado.Datos.Count == 0)
                {
                    Estado("Ningún archivo tenía formato de lista de precios.", Brushes.Red);
                    MessageBox.Show(
                        "No se encontró el encabezado esperado en ninguno de los archivos.\n\n" +
                        "La comparativa necesita los reportes de precios, con las columnas " +
                        "\"Cód. Artículo\", \"Descripción\", \"Precio costo\", \"Imp. ventas\", " +
                        "\"Porc. utilidad\" y \"Precio IVI\".\n\n" +
                        "Los archivos de ventas se cargan en la ventana principal, no acá.",
                        "Formato no reconocido");
                    return;
                }

                _motor = new ComparativaService(resultado.Datos);

                Estado($"Carga OK: {resultado.ArchivosLeidos} sucursales, " +
                       $"{resultado.Datos.Count.ToString("N0", ResumenDinamico.FormatoCR)} precios en " +
                       $"{cronometro.Elapsed.TotalSeconds:N1} s.", Brushes.Green);

                InicializarFiltros(resultado.Datos);
                Comparar();

                if (resultado.Errores.Count > 0)
                    MessageBox.Show("Archivos con problemas:\n" + string.Join("\n", resultado.Errores), "Aviso");

                if (resultado.Descartados.Count > 0)
                    MessageBox.Show(
                        "Estos archivos se ignoraron porque no tienen formato de lista de precios:\n" +
                        string.Join("\n", resultado.Descartados), "Archivos ignorados");

                if (resultado.ArchivosLeidos < 2)
                    MessageBox.Show(
                        "Sólo se cargó una sucursal. Para comparar precios hacen falta al menos dos " +
                        "listas en la misma carpeta, una por sucursal.", "Falta una sucursal");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error al cargar");
                Estado("Error en la carga.", Brushes.Red);
            }
            finally
            {
                EstablecerOcupado(false, null);
            }
        }

        private void Estado(string texto, Brush color)
        {
            TxtEstadoArchivo.Text = texto;
            TxtEstadoArchivo.Foreground = color;
        }

        private void EstablecerOcupado(bool ocupado, string? mensaje)
        {
            _ocupado = ocupado;
            BarraProgreso.Visibility = ocupado ? Visibility.Visible : Visibility.Collapsed;
            BtnCargarPrecios.IsEnabled = !ocupado;
            Cursor = ocupado ? System.Windows.Input.Cursors.Wait : null;
            if (mensaje != null) Estado(mensaje, Brushes.DimGray);
        }

        private void LimpiarVista()
        {
            GridComparativa.ItemsSource = null;
            GraficoComparativa.Series = Array.Empty<ISeries>();
            _filas = new List<FilaComparativa>();
            foreach (var c in _columnasSucursal) GridComparativa.Columns.Remove(c);
            _columnasSucursal.Clear();
            _sucursalesComparadas = new List<string>();
            foreach (var t in new[] { TxtComparables, TxtConDiferencia, TxtIguales, TxtExclusivos, TxtPromedio })
                t.Text = "-";
        }

        // ==========================================
        // COMPARACIÓN
        // ==========================================
        private void Comparar()
        {
            if (_motor == null) return;

            var sucursales = ObtenerSeleccionados(_opSucursales);
            if (sucursales.Count == 0)
            {
                GridComparativa.ItemsSource = null;
                GraficoComparativa.Series = Array.Empty<ISeries>();
                TxtSubtitulo.Text = "Marque al menos una sucursal para comparar.";
                return;
            }

            _metricaVista = (MetricaPrecio)Math.Max(0, CmbMetrica.SelectedIndex);

            var peticion = new PeticionComparativa
            {
                Metrica = _metricaVista,
                Sucursales = sucursales,
                Presencia = (PresenciaMinima)Math.Max(0, CmbPresencia.SelectedIndex),
                SoloConDiferencia = ChkSoloDiferencia.IsChecked == true,
                UmbralPct = UmbralElegido(),
                Busqueda = TxtBuscarProducto.Text,
                Orden = (OrdenComparativa)Math.Max(0, CmbOrden.SelectedIndex)
            };

            var resultado = _motor.Comparar(peticion);

            if (!resultado.Sucursales.SequenceEqual(_sucursalesComparadas, StringComparer.Ordinal))
                ReconstruirColumnas(resultado.Sucursales);

            _filas = resultado.Filas;
            GridComparativa.ItemsSource = _filas;
            GraficoComparativa.Series = Array.Empty<ISeries>();

            EtiquetarColumnas(_metricaVista);
            MostrarResumen(resultado);
        }

        private decimal UmbralElegido()
        {
            string? tag = (CmbUmbral.SelectedItem as ComboBoxItem)?.Tag?.ToString();
            return decimal.TryParse(tag, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal v) ? v : 0m;
        }

        private void MostrarResumen(ResultadoComparativa r)
        {
            var formato = ResumenDinamico.FormatoCR;

            TxtComparables.Text = r.Comparables.ToString("N0", formato);
            TxtConDiferencia.Text = r.ConDiferencia.ToString("N0", formato);
            TxtIguales.Text = r.Iguales.ToString("N0", formato);
            TxtExclusivos.Text = r.Exclusivos.ToString("N0", formato);
            TxtPromedio.Text = r.Comparables > 0 ? FilaComparativa.FormatearPorcentaje(r.DiferenciaPromedioPct) : "-";

            TxtTitulo.Text = $"{ConjuntoPrecios.NombreDe(_metricaVista)} — {r.Sucursales.Count} sucursales";
            TxtSubtitulo.Text = _filas.Count == 0
                ? "Ningún producto cumple los filtros seleccionados."
                : $"{_filas.Count.ToString("N0", formato)} productos en la tabla. " +
                  "Seleccione uno para verlo sucursal por sucursal.";
        }

        /// <summary>
        /// La tabla lleva una columna por sucursal, así que se construyen en código: el
        /// enlace es por posición (Textos[i] / Colores[i]) contra los arreglos de la fila.
        /// </summary>
        private void ReconstruirColumnas(List<string> sucursales)
        {
            foreach (var c in _columnasSucursal) GridComparativa.Columns.Remove(c);
            _columnasSucursal.Clear();

            for (int i = 0; i < sucursales.Count; i++)
            {
                var estilo = new Style(typeof(TextBlock));
                estilo.Setters.Add(new Setter(TextBlock.ForegroundProperty, new Binding($"Colores[{i}]")));
                estilo.Setters.Add(new Setter(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Right));
                estilo.Setters.Add(new Setter(TextBlock.FontWeightProperty, FontWeights.SemiBold));
                estilo.Setters.Add(new Setter(TextBlock.MarginProperty, new Thickness(0, 0, 4, 0)));

                var columna = new DataGridTextColumn
                {
                    Header = sucursales[i],
                    Binding = new Binding($"Textos[{i}]"),
                    ElementStyle = estilo,
                    Width = new DataGridLength(1.1, DataGridLengthUnitType.Star),
                    MinWidth = 95,
                    // El enlace es por índice y el DataGrid no sabe ordenar por eso:
                    // el orden se elige en el panel izquierdo.
                    CanUserSort = false
                };

                GridComparativa.Columns.Insert(PrimeraColumnaSucursal + i, columna);
                _columnasSucursal.Add(columna);
            }

            _sucursalesComparadas = sucursales;
        }

        /// <summary>En utilidad no hay "más barata": la mejor sucursal es la de mayor margen.</summary>
        private void EtiquetarColumnas(MetricaPrecio metrica)
        {
            bool utilidad = metrica == MetricaPrecio.Utilidad;
            ColumnaMejor.Header = utilidad ? "Mayor utilidad" : "Más barata";
            ColumnaPeor.Header = utilidad ? "Menor utilidad" : "Más cara";
            ColumnaDiferencia.Header = utilidad ? "Dif. (puntos)" : "Dif.";
        }

        // ==========================================
        // GRÁFICO DEL PRODUCTO SELECCIONADO
        // ==========================================
        private void GridComparativa_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GridComparativa.SelectedItem is not FilaComparativa fila)
            {
                GraficoComparativa.Series = Array.Empty<ISeries>();
                return;
            }

            bool utilidad = _metricaVista == MetricaPrecio.Utilidad;
            var valores = fila.Valores.Select(v => v.HasValue ? (double?)v.Value : null).ToArray();

            GraficoComparativa.Series = new ISeries[]
            {
                new ColumnSeries<double?>
                {
                    Name = fila.Descripcion,
                    Values = valores,
                    Fill = new SolidColorPaint(new SKColor(0x8E, 0x44, 0xAD)),
                    DataLabelsPaint = new SolidColorPaint(new SKColor(0x2C, 0x3E, 0x50)),
                    DataLabelsPosition = DataLabelsPosition.Top,
                    DataLabelsFormatter = p => p.Coordinate.PrimaryValue.ToString("N0", ResumenDinamico.FormatoCR),
                    YToolTipLabelFormatter = p => FilaComparativa.Formatear((decimal)p.Coordinate.PrimaryValue, _metricaVista)
                }
            };

            GraficoComparativa.XAxes = new[]
            {
                new Axis { Labels = _sucursalesComparadas, TextSize = 12, LabelsRotation = 0 }
            };
            GraficoComparativa.YAxes = new[]
            {
                new Axis
                {
                    Name = utilidad ? "% de utilidad" : ConjuntoPrecios.NombreDe(_metricaVista),
                    Labeler = v => utilidad ? $"{v:N0}%" : v.ToString("N0", ResumenDinamico.FormatoCR)
                }
            };

            // El detalle completo de los nombres está en el tooltip de la descripción; acá
            // sólo se avisa, para que el subtítulo no se coma dos líneas de la ventana.
            TxtSubtitulo.Text = fila.DescripcionesDistintas
                ? $"{fila.Codigo} — {fila.Descripcion}  ⚠ el nombre cambia entre sucursales."
                : $"{fila.Codigo} — {fila.Descripcion}";
        }

        // ==========================================
        // FILTROS
        // ==========================================
        private void InicializarFiltros(ConjuntoPrecios datos)
        {
            _cargandoFiltros = true;
            try
            {
                foreach (var vieja in _opSucursales) vieja.PropertyChanged -= SucursalCambiada;
                _opSucursales.Clear();

                foreach (string nombre in datos.Sucursales.Valores.OrderBy(x => x, StringComparer.CurrentCulture))
                {
                    var opcion = new OpcionFiltro(nombre, true);
                    opcion.PropertyChanged += SucursalCambiada;
                    _opSucursales.Add(opcion);
                }

                LstSucursales.ItemsSource = null;
                LstSucursales.ItemsSource = _opSucursales;
            }
            finally { _cargandoFiltros = false; }
        }

        private void SucursalCambiada(object? s, System.ComponentModel.PropertyChangedEventArgs e)
            => ProgramarRecalculo();

        private static List<string> ObtenerSeleccionados(List<OpcionFiltro> opciones)
        {
            var lista = new List<string>(opciones.Count);
            foreach (var o in opciones) if (o.Seleccionado) lista.Add(o.Nombre);
            return lista;
        }

        private void ProgramarRecalculo()
        {
            if (_cargandoFiltros || _motor == null) return;
            _temporizador.Stop();
            _temporizador.Start();
        }

        private void MarcarTodas(bool marcado)
        {
            bool alguno = false;
            foreach (var o in _opSucursales)
            {
                if (o.Seleccionado == marcado) continue;
                o.EstablecerSilencioso(marcado);
                o.NotificarSeleccion();   // refresca la casilla, sin recalcular por cada una
                alguno = true;
            }
            if (alguno) ProgramarRecalculo();
        }

        private void BtnTodasSucursal_Click(object sender, RoutedEventArgs e) => MarcarTodas(true);
        private void BtnNingunaSucursal_Click(object sender, RoutedEventArgs e) => MarcarTodas(false);

        private void Opcion_Cambiada(object sender, RoutedEventArgs e) => ProgramarRecalculo();
        private void Opcion_Seleccionada(object sender, SelectionChangedEventArgs e) => ProgramarRecalculo();

        private void TxtBuscarProducto_TextChanged(object sender, TextChangedEventArgs e)
        {
            PistaProducto.Visibility = TxtBuscarProducto.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
            ProgramarRecalculo();
        }
    }
}
