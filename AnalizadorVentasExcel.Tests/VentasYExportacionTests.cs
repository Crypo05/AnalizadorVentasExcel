using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using AnalizadorVentasExcel.Modelos;
using AnalizadorVentasExcel.Servicios;
using ExcelDataReader;
using static AnalizadorVentasExcel.Tests.LibroXlsx;

namespace AnalizadorVentasExcel.Tests
{
    public class VentasTests : PruebaConCarpeta
    {
        private static Task<ResultadoCarga> CargarVentas(params string[] rutas)
            => new ExcelService().CargarCarpetaAsync(rutas, null);

        [Fact]
        public async Task D1_Ventas_con_codigo_y_costo_se_cargan_como_ventas()
        {
            string ruta = Escribir(Carpeta, "ventas.xlsx",
                Fila("Año mes", "Cód. Artículo", "Artículo desc.", "Proveedor", "Familia", "Precio costo", "Total", "% Utilidad"),
                Fila("2025-01", "001", "PROD A", "PROV A", "FAM A", 800, 1000, 0.2),
                Fila("2025-01", "002", "PROD B", "PROV A", "FAM A", 1800, 2500, 0.2));

            var r = await CargarVentas(ruta);
            var d = Assert.Single(r.Diagnosticos);

            Assert.NotEqual(EstadoCarga.NoCargado, d.Estado);
            Assert.NotEqual(ProblemaCarga.ReporteDelOtroSistema, d.Problema);
            Assert.Equal(2, r.Datos.Count);
            Assert.Equal(3500m, r.Datos.Filas.Sum(f => f.TotalVenta));
        }

        [Fact]
        public async Task D2_Formato_real_de_la_caja_descarta_totales_y_encabezados_repetidos()
        {
            string ruta = Escribir(Carpeta, "ventas.xlsx",
                Fila("Total"),
                Fila(null, null, null, null, null, "Total general"),
                Fila("Año mes", "Artículo", "Artículo desc.", "Proveedor oficial - razón comercial", "Familia", "Total", "% Utilidad"),
                Fila("2025-01", "001", "A", "P1", "F1", 1000, 0.3),
                Fila(null, "002", "B", "P1", "F1", 500, 0.2),
                Fila("2025-01 Total", null, null, null, null, 1500),
                Fila("Total general", null, null, null, null, 1500),
                Fila("Año mes", "Artículo", "Artículo desc.", "Proveedor oficial - razón comercial", "Familia", "Total", "% Utilidad"),
                Fila("2025-02", "001", "A", "P1", "F1", 700, 0.3));

            var r = await CargarVentas(ruta);

            Assert.Equal(3, r.Datos.Count);
            Assert.Equal(2200m, r.Datos.Filas.Sum(f => f.TotalVenta));
            var periodos = r.Datos.Filas.Select(f => r.Datos.Periodos[f.PeriodoId]).Distinct().OrderBy(p => p).ToArray();
            Assert.Equal(new[] { "2025-01", "2025-02" }, periodos);
        }

        [Fact]
        public async Task D3_Un_encabezado_incompleto_que_no_es_lista_de_precios_es_sin_encabezado()
        {
            string ruta = Escribir(Carpeta, "ventas.xlsx",
                Fila("Año mes", "Descripción", "Precio"),
                Fila("2025-01", "X", 100));

            var r = await CargarVentas(ruta);
            var d = Assert.Single(r.Diagnosticos);

            Assert.NotEqual(ProblemaCarga.ReporteDelOtroSistema, d.Problema);
            Assert.Equal(ProblemaCarga.SinEncabezado, d.Problema);
        }
    }

    public class ExportacionTests : PruebaConListas
    {
        private static List<object?[]> Leer(string ruta)
        {
            using var s = File.OpenRead(ruta);
            using var lector = ExcelReaderFactory.CreateReader(s);
            var filas = new List<object?[]>();
            while (lector.Read())
            {
                var fila = new object?[lector.FieldCount];
                for (int i = 0; i < fila.Length; i++) fila[i] = lector.GetValue(i);
                filas.Add(fila);
            }
            return filas;
        }

        private async Task<ResultadoCargaPrecios> Las3() => await CargarPrecios(Bomba(), Catarata(), Mirador());

        [Fact]
        public async Task F1_El_aviso_va_arriba_y_el_encabezado_queda_debajo()
        {
            var r = await Las3();
            var v = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "catarata", "mirador");
            var aviso = AvisoCruce.Lineas(r.Datos.CrucesPorDescripcion);
            string ruta = Path.Combine(Carpeta, "salida.xlsx");

            ExportadorExcel.ExportarComparativa(ruta, v.Filas, v.Sucursales, MetricaPrecio.PrecioVenta, aviso);
            var filas = Leer(ruta);

            Assert.Equal(aviso[0], filas[0][0]);
            Assert.Equal("Código", filas[aviso.Count + 1][0]);
        }

        [Fact]
        public async Task F2_Sin_aviso_el_encabezado_esta_en_la_primera_fila()
        {
            var r = await Las3();
            var v = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "catarata", "mirador");
            string ruta = Path.Combine(Carpeta, "salida.xlsx");

            ExportadorExcel.ExportarComparativa(ruta, v.Filas, v.Sucursales, MetricaPrecio.PrecioVenta);

            Assert.Equal("Código", Leer(ruta)[0][0]);
        }

        [Fact]
        public async Task F3_En_la_vista_costo_la_celda_de_mirador_queda_vacia_y_no_en_cero()
        {
            var r = await Las3();
            var v = Comparar(r.Datos, MetricaPrecio.Costo, "bomba", "catarata", "mirador");
            string ruta = Path.Combine(Carpeta, "salida.xlsx");

            ExportadorExcel.ExportarComparativa(ruta, v.Filas, v.Sucursales, MetricaPrecio.Costo);
            var filas = Leer(ruta);

            int colMirador = Array.FindIndex(filas[0], c => (c as string) == "mirador");
            Assert.True(colMirador >= 0);
            var f001 = Assert.Single(filas, f => (f[0] as string) == "001");

            Assert.True(f001[colMirador] is null || (f001[colMirador] as string) == "",
                $"La celda de mirador debería estar vacía y tiene: {f001[colMirador]}");
            int colBomba = Array.FindIndex(filas[0], c => (c as string) == "bomba");
            Assert.Equal(200d, Convert.ToDouble(f001[colBomba]));
        }
    }
}
