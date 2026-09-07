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

        /// <summary>Si la tabla actual muestra costo y venta a la vez (dos columnas por sucursal).</summary>
        private bool _vistaCombinada;

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
            BtnExportar.IsEnabled = false;
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
                BtnExportar.IsEnabled = false;
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

            // Las columnas se rehacen si cambian las sucursales o si se entra o sale de la
            // vista combinada, que lleva dos columnas por sucursal en vez de una.
            bool combinada = ConjuntoPrecios.EsCombinada(_metricaVista);
            if (combinada != _vistaCombinada ||
                !resultado.Sucursales.SequenceEqual(_sucursalesComparadas, StringComparer.Ordinal))
                ReconstruirColumnas(resultado.Sucursales, combinada);

            _filas = resultado.Filas;
            GridComparativa.ItemsSource = _filas;
            GraficoComparativa.Series = Array.Empty<ISeries>();

            EtiquetarColumnas(_metricaVista);
            MostrarResumen(resultado);

            BtnExportar.IsEnabled = _filas.Count > 0;
        }

        // ==========================================
        // EXPORTACIÓN
        // ==========================================

        /// <summary>
        /// Exporta lo mismo que está viendo el usuario: las filas que quedaron después de
        /// los filtros, en el orden en que están, y con las mismas columnas (incluida una
        /// por sucursal comparada). Cambiar un filtro y volver a exportar da otro archivo.
        /// </summary>
        private void BtnExportar_Click(object sender, RoutedEventArgs e)
        {
            if (_filas.Count == 0)
            {
                MessageBox.Show("No hay nada que exportar: la tabla está vacía.", "Sin datos");
                return;
            }

            var dialogo = new SaveFileDialog
            {
                Title = "Guardar la comparativa",
                Filter = "Libro de Excel|*.xlsx",
                FileName = NombreSugerido(),
                DefaultExt = ".xlsx",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (dialogo.ShowDialog() != true) return;

            try
            {
                ExportadorExcel.ExportarComparativa(dialogo.FileName, _filas, _sucursalesComparadas, _metricaVista);

                Estado($"Exportados {_filas.Count.ToString("N0", ResumenDinamico.FormatoCR)} productos a " +
                       $"{Path.GetFileName(dialogo.FileName)}", Brushes.Green);

                var abrir = MessageBox.Show(
                    $"Se guardaron {_filas.Count.ToString("N0", ResumenDinamico.FormatoCR)} productos en:\n" +
                    $"{dialogo.FileName}\n\n¿Abrir el archivo ahora?",
                    "Exportación lista", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (abrir == MessageBoxResult.Yes)
                    Process.Start(new ProcessStartInfo(dialogo.FileName) { UseShellExecute = true });
            }
            catch (IOException ex)
            {
                // Lo más común con diferencia: el archivo quedó abierto en Excel.
                MessageBox.Show(
                    "No se pudo escribir el archivo. Si lo tenés abierto en Excel, cerralo y probá de nuevo.\n\n" +
                    ex.Message, "No se pudo guardar");
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("No hay permisos para escribir en esa carpeta. Probá guardarlo en el Escritorio.",
                                "No se pudo guardar");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error al exportar");
            }
        }

        private string NombreSugerido()
        {
            string metrica = _metricaVista switch
            {
                MetricaPrecio.Costo => "costos",
                MetricaPrecio.Utilidad => "utilidad",
                MetricaPrecio.CostoYVenta => "costos y precios",
                _ => "precios"
            };
            return $"Comparativa de {metrica} {DateTime.Now:yyyy-MM-dd}.xlsx";
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
        /// La tabla lleva una columna por sucursal —dos en la vista combinada, costo y
        /// venta—, así que se construyen en código: el enlace es por posición
        /// (Textos[i] / Colores[i]) contra los arreglos intercalados de la fila.
        /// </summary>
        private void ReconstruirColumnas(List<string> sucursales, bool combinada)
        {
            foreach (var c in _columnasSucursal) GridComparativa.Columns.Remove(c);
            _columnasSucursal.Clear();

            var subMetricas = combinada
                ? new[] { "costo", "venta" }
                : new[] { string.Empty };

            int posicion = PrimeraColumnaSucursal;

            for (int s = 0; s < sucursales.Count; s++)
            {
                for (int k = 0; k < subMetricas.Length; k++)
                {
                    int indice = s * subMetricas.Length + k;

                    var estilo = new Style(typeof(TextBlock));
                    estilo.Setters.Add(new Setter(TextBlock.ForegroundProperty, new Binding($"Colores[{indice}]")));
                    estilo.Setters.Add(new Setter(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Right));
                    estilo.Setters.Add(new Setter(TextBlock.FontWeightProperty, FontWeights.SemiBold));
                    estilo.Setters.Add(new Setter(TextBlock.MarginProperty, new Thickness(0, 0, 4, 0)));

                    var columna = new DataGridTextColumn
                    {
                        Header = combinada ? $"{sucursales[s]} · {subMetricas[k]}" : sucursales[s],
                        Binding = new Binding($"Textos[{indice}]"),
                        ElementStyle = estilo,
                        // Con la vista combinada las columnas se cuentan de a dos por
                        // sucursal: con anchos proporcionales quedarían aplastadas, así que
                        // se fijan y la tabla se desplaza en horizontal.
                        Width = combinada ? new DataGridLength(110) : new DataGridLength(1.1, DataGridLengthUnitType.Star),
                        MinWidth = combinada ? 100 : 95,
                        // El enlace es por índice y el DataGrid no sabe ordenar por eso:
                        // el orden se elige en el panel izquierdo.
                        CanUserSort = false
                    };

                    GridComparativa.Columns.Insert(posicion++, columna);
                    _columnasSucursal.Add(columna);
                }
            }

            AjustarColumnasFijas(combinada);
            _sucursalesComparadas = sucursales;
            _vistaCombinada = combinada;
        }

        /// <summary>
        /// Las columnas declaradas en el XAML usan anchos proporcionales, que reparten el
        /// espacio disponible. Eso funciona con pocas columnas, pero en la vista combinada
        /// (hasta veintipico) hay que pasarlas a ancho fijo o el DataGrid las comprime en
        /// vez de habilitar el desplazamiento horizontal.
        /// </summary>
        private void AjustarColumnasFijas(bool combinada)
        {
            ColumnaDiferenciaVenta.Visibility = combinada ? Visibility.Visible : Visibility.Collapsed;
            ColumnaDiferenciaVentaPct.Visibility = combinada ? Visibility.Visible : Visibility.Collapsed;

            // En la vista combinada las cuatro columnas de diferencia ya dicen quién está
            // más barata, y el color de cada celda lo remata: mostrar además "Más barata"
            // y "Más cara" sólo del costo confundiría.
            ColumnaMejor.Visibility = combinada ? Visibility.Collapsed : Visibility.Visible;
            ColumnaPeor.Visibility = combinada ? Visibility.Collapsed : Visibility.Visible;

            if (combinada)
            {
                ColumnaCodigo.Width = new DataGridLength(115);
                ColumnaDescripcion.Width = new DataGridLength(230);
                ColumnaDiferencia.Width = new DataGridLength(105);
                ColumnaDiferenciaPct.Width = new DataGridLength(90);
                ColumnaDiferenciaVenta.Width = new DataGridLength(105);
                ColumnaDiferenciaVentaPct.Width = new DataGridLength(90);
                ColumnaPresencia.Width = new DataGridLength(50);
                ColumnaAviso.Width = new DataGridLength(130);
            }
            else
            {
                ColumnaCodigo.Width = new DataGridLength(1.1, DataGridLengthUnitType.Star);
                ColumnaDescripcion.Width = new DataGridLength(3, DataGridLengthUnitType.Star);
                ColumnaDiferencia.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
                ColumnaDiferenciaPct.Width = new DataGridLength(0.8, DataGridLengthUnitType.Star);
                ColumnaPresencia.Width = new DataGridLength(0.5, DataGridLengthUnitType.Star);
                ColumnaAviso.Width = new DataGridLength(1.6, DataGridLengthUnitType.Star);
            }
        }

        /// <summary>En utilidad no hay "más barata": la mejor sucursal es la de mayor margen.</summary>
        private void EtiquetarColumnas(MetricaPrecio metrica)
        {
            if (ConjuntoPrecios.EsCombinada(metrica))
            {
                // Con las dos métricas juntas hay que decir de cuál es cada diferencia.
                ColumnaDiferencia.Header = "Dif. costo";
                ColumnaDiferenciaPct.Header = "Dif. costo %";
                return;
            }

            bool utilidad = metrica == MetricaPrecio.Utilidad;
            ColumnaMejor.Header = utilidad ? "Mayor utilidad" : "Más barata";
            ColumnaPeor.Header = utilidad ? "Menor utilidad" : "Más cara";
            ColumnaDiferencia.Header = utilidad ? "Dif. (puntos)" : "Dif.";
            ColumnaDiferenciaPct.Header = "Dif. %";
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
            bool combinada = ConjuntoPrecios.EsCombinada(_metricaVista);

            // En la vista combinada los valores vienen intercalados (costo, venta) por
            // sucursal, así que se parten en dos series que quedan una al lado de la otra.
            GraficoComparativa.Series = combinada
                ? new ISeries[]
                {
                    SerieProducto("Costo", Desintercalar(fila.Valores, 0, 2),
                                  new SKColor(0x8E, 0x44, 0xAD), MetricaPrecio.Costo),
                    SerieProducto("Venta", Desintercalar(fila.Valores, 1, 2),
                                  new SKColor(0x16, 0xA0, 0x85), MetricaPrecio.PrecioVenta)
                }
                : new ISeries[]
                {
                    SerieProducto(fila.Descripcion, Desintercalar(fila.Valores, 0, 1),
                                  new SKColor(0x8E, 0x44, 0xAD), _metricaVista)
                };

            GraficoComparativa.LegendPosition = combinada ? LegendPosition.Top : LegendPosition.Hidden;

            GraficoComparativa.XAxes = new[]
            {
                new Axis { Labels = _sucursalesComparadas, TextSize = 12, LabelsRotation = 0 }
            };
            GraficoComparativa.YAxes = new[]
            {
                new Axis
                {
                    Name = combinada ? "Colones" : utilidad ? "% de utilidad" : ConjuntoPrecios.NombreDe(_metricaVista),
                    Labeler = v => utilidad && !combinada ? $"{v:N0}%" : v.ToString("N0", ResumenDinamico.FormatoCR)
                }
            };

            // El detalle completo de los nombres está en el tooltip de la descripción; acá
            // sólo se avisa, para que el subtítulo no se coma dos líneas de la ventana.
            TxtSubtitulo.Text = fila.DescripcionesDistintas
                ? $"{fila.Codigo} — {fila.Descripcion}  ⚠ el nombre cambia entre sucursales."
                : $"{fila.Codigo} — {fila.Descripcion}";
        }

        /// <summary>Toma una de cada <paramref name="paso"/> posiciones, empezando en <paramref name="desde"/>.</summary>
        private static double?[] Desintercalar(decimal?[] valores, int desde, int paso)
        {
            var salida = new double?[(valores.Length - desde + paso - 1) / paso];
            for (int i = desde, n = 0; i < valores.Length; i += paso, n++)
                salida[n] = valores[i].HasValue ? (double?)valores[i]!.Value : null;
            return salida;
        }

        private static ColumnSeries<double?> SerieProducto(string nombre, double?[] valores,
                                                           SKColor color, MetricaPrecio metrica)
            => new ColumnSeries<double?>
            {
                Name = nombre,
                Values = valores,
                Fill = new SolidColorPaint(color),
                DataLabelsPaint = new SolidColorPaint(new SKColor(0x2C, 0x3E, 0x50)),
                DataLabelsPosition = DataLabelsPosition.Top,
                DataLabelsFormatter = p => p.Coordinate.PrimaryValue.ToString("N0", ResumenDinamico.FormatoCR),
                YToolTipLabelFormatter = p =>
                    $"{p.Context.Series.Name}: {FilaComparativa.Formatear((decimal)p.Coordinate.PrimaryValue, metrica)}"
            };

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
