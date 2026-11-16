using AnalizadorVentasExcel.Servicios;
using Xunit;

namespace AnalizadorVentasExcel.Tests
{
    public class HumoTests
    {
        [Fact]
        public void Compactar_quita_tildes_mayusculas_y_espacios_de_mas()
            => Assert.Equal("cafe con leche", CeldasExcel.Compactar("  CAFÉ   con  Leche "));
    }
}
