using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AnalizadorVentasExcel.Modelos
{
    /// <summary>
    /// Un elemento marcable de un checklist de filtros.
    ///
    /// La selección vive aquí, en el modelo, y no en ListBox.SelectedItems. Es lo que
    /// permite buscar sin perder lo marcado: al filtrar la vista, los elementos ocultos
    /// desaparecerían de SelectedItems y se deseleccionarían solos.
    /// </summary>
    public sealed class OpcionFiltro : INotifyPropertyChanged
    {
        private bool _seleccionado;

        public OpcionFiltro(string nombre, bool seleccionado)
        {
            Nombre = nombre;
            _seleccionado = seleccionado;
        }

        public string Nombre { get; }

        public bool Seleccionado
        {
            get => _seleccionado;
            set
            {
                if (_seleccionado == value) return;
                _seleccionado = value;
                Notificar();
            }
        }

        /// <summary>Cambia el estado sin avisar, para actualizaciones masivas.</summary>
        public void EstablecerSilencioso(bool valor) => _seleccionado = valor;

        public void Notificar([CallerMemberName] string propiedad = "")
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propiedad));

        public void NotificarSeleccion()
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Seleccionado)));

        public event PropertyChangedEventHandler? PropertyChanged;

        public override string ToString() => Nombre;
    }
}
