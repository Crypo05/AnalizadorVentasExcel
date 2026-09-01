using System.Windows;
using AnalizadorVentasExcel.Modelos;

namespace AnalizadorVentasExcel
{
    public partial class VentanaGuia : Window
    {
        public VentanaGuia()
        {
            InitializeComponent();

            ListaSecciones.ItemsSource = ContenidoGuia.Secciones();
            ListaRecetas.ItemsSource = ContenidoGuia.Recetas();
            ListaNotas.ItemsSource = ContenidoGuia.Notas();

            TxtPie.Text = $"Analizador Corporativo v{MainWindow.VersionActual} — Mateo Sanabria";
        }

        private void BtnCerrar_Click(object sender, RoutedEventArgs e) => Close();
    }
}
