using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using AnalizadorVentasExcel.Servicios;

namespace AnalizadorVentasExcel
{
    public partial class VentanaActualizacion : Window
    {
        private readonly ActualizacionService _servicio = new();
        // Se renueva en cada intento: si no, tras cancelar una descarga el siguiente
        // intento fallaría al instante con el token ya cancelado.
        private CancellationTokenSource _cancelacion = new();
        private InfoActualizacion? _info;
        private bool _descargando;

        public VentanaActualizacion()
        {
            InitializeComponent();
            Loaded += async (_, _) => await Consultar();
            Closing += (s, e) =>
            {
                if (_descargando)
                {
                    var r = MessageBox.Show("La descarga está en curso. ¿Cancelarla?", "Actualización",
                                            MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (r == MessageBoxResult.No) { e.Cancel = true; return; }
                }
                _cancelacion.Cancel();
            };
        }

        private async Task Consultar()
        {
            TxtVersiones.Text = $"Versión instalada: {MainWindow.VersionActual}";
            TxtProgreso.Text = $"Consultando {ActualizacionService.Repositorio}...";

            try
            {
                var r = await _servicio.ConsultarAsync(MainWindow.VersionActual, _cancelacion.Token);
                _info = r.Info;
                TxtProgreso.Text = string.Empty;
                BtnVerEnGitHub.Visibility = string.IsNullOrEmpty(_info?.UrlPagina) ? Visibility.Collapsed : Visibility.Visible;

                switch (r.Estado)
                {
                    case EstadoActualizacion.HayNueva:
                        TxtTitulo.Text = $"Hay una versión nueva: {_info!.Etiqueta}";
                        TxtVersiones.Text = $"Tenés la {MainWindow.VersionActual} — se puede actualizar a la {_info.Version}" +
                                            (_info.Tamano > 0 ? $" ({_info.Tamano / 1024d / 1024d:N0} MB de descarga)" : "");
                        TxtNotas.Text = string.IsNullOrWhiteSpace(_info.Notas)
                            ? "(El release no incluye notas.)" : _info.Notas;
                        BtnAccion.IsEnabled = !string.IsNullOrEmpty(_info.UrlDescarga);
                        if (!BtnAccion.IsEnabled)
                            TxtProgreso.Text = "El release publicado no trae el archivo AnalizadorVentasExcel.exe, " +
                                               "así que no se puede instalar automáticamente.";
                        break;

                    case EstadoActualizacion.AlDia:
                        TxtTitulo.Text = "Ya tenés la última versión";
                        TxtVersiones.Text = $"Versión {MainWindow.VersionActual}, igual que la publicada.";
                        TxtNotas.Text = _info?.Notas ?? string.Empty;
                        break;

                    case EstadoActualizacion.LocalMasNueva:
                        TxtTitulo.Text = "Tu versión es más reciente que la publicada";
                        TxtVersiones.Text = $"Instalada {MainWindow.VersionActual} — última publicada {_info!.Etiqueta}. " +
                                            "No hay nada que actualizar.";
                        TxtNotas.Text = _info.Notas;
                        break;
                }
            }
            catch (OperationCanceledException) { /* la ventana se cerró */ }
            catch (Exception ex)
            {
                TxtTitulo.Text = "No se pudo comprobar";
                TxtVersiones.Text = string.Empty;
                TxtNotas.Text = ex.Message;
                TxtProgreso.Text = string.Empty;
            }
        }

        private async void BtnAccion_Click(object sender, RoutedEventArgs e)
        {
            if (_info == null || _descargando) return;

            string? exe = Environment.ProcessPath;
            if (string.IsNullOrEmpty(exe))
            {
                MessageBox.Show("No se pudo determinar la ruta del programa.", "Actualización");
                return;
            }

            if (!ActualizacionService.PuedeEscribirJuntoAlExe(exe, out string motivo))
            {
                MessageBox.Show(
                    "No hay permiso de escritura en la carpeta del programa, así que no se puede " +
                    $"reemplazar el ejecutable.\n\nCarpeta: {Path.GetDirectoryName(exe)}\nDetalle: {motivo}\n\n" +
                    "Probá ejecutar el programa como administrador, o descargá la versión nueva a mano " +
                    "desde GitHub.", "Actualización", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // El temporal va en la misma carpeta para que el reemplazo sea un renombrado
            // dentro del mismo volumen y no una copia entre discos.
            string temporal = exe + ".nuevo";

            _cancelacion.Dispose();
            _cancelacion = new CancellationTokenSource();

            _descargando = true;
            BtnAccion.IsEnabled = false;
            BtnCerrar.Content = "Cancelar";
            BarraDescarga.Visibility = Visibility.Visible;

            try
            {
                var progreso = new Progress<(long recibidos, long total)>(p =>
                {
                    if (p.total > 0)
                    {
                        BarraDescarga.IsIndeterminate = false;
                        BarraDescarga.Value = 100d * p.recibidos / p.total;
                        TxtProgreso.Text = $"Descargando... {p.recibidos / 1024d / 1024d:N1} MB de {p.total / 1024d / 1024d:N1} MB";
                    }
                    else
                    {
                        BarraDescarga.IsIndeterminate = true;
                        TxtProgreso.Text = $"Descargando... {p.recibidos / 1024d / 1024d:N1} MB";
                    }
                });

                await _servicio.DescargarAsync(_info, temporal, progreso, _cancelacion.Token);

                TxtProgreso.Text = "Instalando...";
                string respaldo = ActualizacionService.Instalar(temporal, exe);

                _descargando = false;
                var r = MessageBox.Show(
                    $"Listo, quedó instalada la versión {_info.Etiqueta}.\n\n" +
                    "Hay que reiniciar el programa para usarla. ¿Reiniciar ahora?\n\n" +
                    $"(La versión anterior quedó guardada como {Path.GetFileName(respaldo)} y se borrará sola.)",
                    "Actualización completada", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (r == MessageBoxResult.Yes)
                {
                    ActualizacionService.Reiniciar(exe);
                    Application.Current.Shutdown();
                }
                else Close();
            }
            catch (OperationCanceledException)
            {
                TxtProgreso.Text = "Descarga cancelada.";
                BorrarTemporal(temporal);
            }
            catch (Exception ex)
            {
                BorrarTemporal(temporal);
                MessageBox.Show($"No se pudo completar la actualización:\n\n{ex.Message}\n\n" +
                                "El programa quedó intacto; podés seguir usándolo.",
                                "Actualización", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtProgreso.Text = "La actualización falló.";
            }
            finally
            {
                _descargando = false;
                BarraDescarga.Visibility = Visibility.Collapsed;
                BtnCerrar.Content = "Cerrar";
                BtnAccion.IsEnabled = _info != null && !string.IsNullOrEmpty(_info.UrlDescarga);
            }
        }

        private static void BorrarTemporal(string ruta)
        {
            try { if (File.Exists(ruta)) File.Delete(ruta); } catch { /* se limpia en el próximo arranque */ }
        }

        private void BtnVerEnGitHub_Click(object sender, RoutedEventArgs e)
        {
            if (_info == null || string.IsNullOrEmpty(_info.UrlPagina)) return;
            try { Process.Start(new ProcessStartInfo { FileName = _info.UrlPagina, UseShellExecute = true }); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "No se pudo abrir el navegador"); }
        }

        private void BtnCerrar_Click(object sender, RoutedEventArgs e)
        {
            if (_descargando) _cancelacion.Cancel();
            else Close();
        }
    }
}
