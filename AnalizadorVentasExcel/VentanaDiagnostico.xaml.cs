using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using AnalizadorVentasExcel.Modelos;
using AnalizadorVentasExcel.Servicios;

namespace AnalizadorVentasExcel
{
    public partial class VentanaDiagnostico : Window
    {
        private readonly IReadOnlyList<DiagnosticoArchivo> _diagnosticos;

        public VentanaDiagnostico(IReadOnlyList<DiagnosticoArchivo> diagnosticos)
        {
            InitializeComponent();

            _diagnosticos = diagnosticos;

            int cargados = diagnosticos.Count(d => d.Estado == EstadoCarga.Cargado);
            int conAvisos = diagnosticos.Count(d => d.Estado == EstadoCarga.ConAvisos);
            int noCargados = diagnosticos.Count(d => d.Estado == EstadoCarga.NoCargado);

            TxtResumen.Text = noCargados == 0 && conAvisos == 0
                ? $"{diagnosticos.Count.ToString("N0", ResumenDinamico.FormatoCR)} archivos, todos cargados sin avisos."
                : $"{diagnosticos.Count.ToString("N0", ResumenDinamico.FormatoCR)} archivos · " +
                  $"{cargados.ToString("N0", ResumenDinamico.FormatoCR)} cargados · " +
                  $"{conAvisos.ToString("N0", ResumenDinamico.FormatoCR)} con avisos · " +
                  $"{noCargados.ToString("N0", ResumenDinamico.FormatoCR)} no cargados";

            GridDiagnostico.ItemsSource = diagnosticos
                .OrderByDescending(d => d.Estado)
                .ThenBy(d => d.Archivo)
                .ToList();

            if (GridDiagnostico.Items.Count > 0)
                GridDiagnostico.SelectedIndex = 0;
        }

        private void GridDiagnostico_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GridDiagnostico.SelectedItem is DiagnosticoArchivo diagnostico)
            {
                TxtDetalleTitulo.Text = $"{diagnostico.Icono} {diagnostico.Archivo}";
                ListaDetalles.ItemsSource = diagnostico.Detalles;
            }
            else
            {
                TxtDetalleTitulo.Text = "Seleccioná un archivo para ver el detalle.";
                ListaDetalles.ItemsSource = null;
            }
        }

        private void BtnCopiarInforme_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Clipboard.SetText(DiagnosticoCarga.Informe(_diagnosticos));
            }
            catch
            {
                MessageBox.Show("No se pudo copiar al portapapeles. Probá de nuevo.", "Portapapeles");
                return;
            }

            string textoOriginal = BtnCopiarInforme.Content?.ToString() ?? "📋 Copiar informe";
            BtnCopiarInforme.Content = "✓ Copiado";

            var temporizador = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
            temporizador.Tick += (_, _) =>
            {
                BtnCopiarInforme.Content = textoOriginal;
                temporizador.Stop();
            };
            temporizador.Start();
        }

        private void BtnCerrar_Click(object sender, RoutedEventArgs e) => Close();
    }
}
