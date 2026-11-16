using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using AnalizadorVentasExcel.Modelos;
using AnalizadorVentasExcel.Servicios;
using static AnalizadorVentasExcel.Tests.LibroXlsx;

namespace AnalizadorVentasExcel.Tests
{
    /// <summary>Datos sintéticos y atajos compartidos por las pruebas de listas de precios.</summary>
    public abstract class PruebaConListas : PruebaConCarpeta
    {
        protected static readonly object?[] EncabezadoEstandar =
            { "Cód. Artículo", "Descripción", "Precio costo", "Imp. ventas", "Porc. utilidad - artículo", "Precio IVI - artículo" };

        protected static readonly object?[] EncabezadoUsuario =
            { "Descripción artículo", "Precio venta I.V.A.", "Familia", "Costo I.V.A." };

        protected string Bomba() => Escribir(Carpeta, "precios bomba.xlsx",
            EncabezadoEstandar,
            Fila("001", "CAFÉ NEGRO", 200, 13, 50, 800),
            Fila("002", "LECHE ENTERA 1 L", 400, 13, 30, 900),
            Fila("003", "PAN BLANCO", 300, 13, 40, 700),
            Fila("005", "COCA COLA 600 ML", 500, 13, 40, 1000));

        protected string Catarata() => Escribir(Carpeta, "precios catarata.xlsx",
            EncabezadoEstandar,
            Fila("001", "CAFE NEGRO GRANDE", 210, 13, 50, 850),
            Fila("002", "LECHE ENTERA 1L", 400, 13, 30, 950),
            Fila("004", "ARROZ 1 KG", 600, 13, 20, 1200),
            Fila("006", "COCA COLA 600 ML", 520, 13, 40, 1050));

        protected string Mirador() => Escribir(Carpeta, "precios mirador.xlsx",
            EncabezadoUsuario,
            Fila("  cafe   NEGRO ", 820, "0171", 230),
            Fila("Leche Entera 1 L", 910, "0171", 452),
            Fila("ARROZ 1 KG", 1250, "0050", 678),
            Fila("COCA COLA 600 ML", 1100, "0021", 600),
            Fila("PRODUCTO SOLO MIRADOR", 500, "0000", 300),
            Fila("PAN BLANCO GRANDE", 750, "0040", 340),
            Fila("DURACELL BATERIA AAA UND", 550, "0066", 400),
            Fila("DURACELL BATERIA AAA UND", 3500, "0066", 2000),
            new object?[] { null },
            Fila(null, 99999, null, 88888));

        /// <summary>Lista sin código con las descripciones dadas (precio de venta consecutivo).</summary>
        protected string ListaSinCodigo(string archivo, params string[] descripciones)
        {
            var filas = new List<object?[]> { EncabezadoUsuario };
            decimal precio = 1000;
            foreach (var d in descripciones) { filas.Add(Fila(d, precio, "0001", precio / 2)); precio += 100; }
            return Escribir(Carpeta, archivo, filas.ToArray());
        }

        protected static Task<ResultadoCargaPrecios> CargarPrecios(params string[] rutas)
            => new PreciosExcelService().CargarCarpetaAsync(rutas);

        protected static ResultadoComparativa Comparar(ConjuntoPrecios datos, MetricaPrecio metrica, params string[] sucursales)
            => new ComparativaService(datos).Comparar(new PeticionComparativa
            {
                Metrica = metrica,
                Sucursales = sucursales,
                Presencia = PresenciaMinima.Todos,
                Orden = OrdenComparativa.Codigo
            });

        protected static FilaComparativa FilaDe(ResultadoComparativa r, string codigo)
            => Assert.Single(r.Filas, f => f.Codigo == codigo);

        protected static decimal? Valor(ResultadoComparativa r, FilaComparativa f, string sucursal)
        {
            int i = r.Sucursales.IndexOf(sucursal);
            Assert.True(i >= 0, $"La sucursal {sucursal} no está en el resultado.");
            return f.Valores[i];
        }
    }
}
