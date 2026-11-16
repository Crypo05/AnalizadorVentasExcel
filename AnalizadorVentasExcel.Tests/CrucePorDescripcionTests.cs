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
    public class CrucePorDescripcionTests : PruebaConListas
    {
        private const string SinCodigo = "(sin código)";

        private async Task<ResultadoCargaPrecios> Las3() => await CargarPrecios(Bomba(), Catarata(), Mirador());

        private static DiagnosticoArchivo DeMirador(ResultadoCargaPrecios r) => r.Diagnosticos[2];

        // ---------- A. Mensajes con el encabezado del usuario ----------

        [Fact]
        public async Task A1_Ventas_rechaza_la_lista_del_usuario_como_reporte_del_otro_sistema()
        {
            string ruta = Mirador();
            var r = await new ExcelService().CargarCarpetaAsync(new[] { ruta }, null);
            var d = Assert.Single(r.Diagnosticos);

            Assert.Equal(EstadoCarga.NoCargado, d.Estado);
            Assert.Equal(ProblemaCarga.ReporteDelOtroSistema, d.Problema);
            Assert.Contains("lista de precios", d.Mensaje);
            Assert.Contains("Cód. Artículo", d.Mensaje);
            Assert.Contains("Porc. utilidad", d.Mensaje);
            Assert.Contains("Precio costo", d.Mensaje);
            Assert.Contains("Costo I.V.A.", d.Mensaje);
            Assert.DoesNotContain("total", d.Mensaje, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Comparativa de Precios", d.Sugerencia);
            Assert.DoesNotContain("total", d.Sugerencia, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task A2_Precios_carga_la_lista_del_usuario_con_un_solo_aviso_de_columnas()
        {
            var r = await Las3();
            var d = DeMirador(r);

            Assert.Equal(EstadoCarga.ConAvisos, d.Estado);
            Assert.Equal(ProblemaCarga.CruceSinCodigo, d.Problema);

            var lineas = d.Detalles.Where(l =>
                l.Contains("Cód. Artículo") && l.Contains("Precio costo") && l.Contains("Costo I.V.A.") &&
                l.Contains("Imp. ventas") && l.Contains("Porc. utilidad")).ToList();
            var unica = Assert.Single(lineas);

            Assert.DoesNotContain("total", d.Mensaje, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("total", d.Sugerencia, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("total", unica, StringComparison.OrdinalIgnoreCase);
        }

        // ---------- B. Cruce por descripción ----------

        [Fact]
        public async Task B1_Conteos_del_cruce_de_mirador()
        {
            var r = await Las3();
            var c = Assert.Single(r.Datos.CrucesPorDescripcion);

            Assert.Equal("mirador", c.Sucursal);
            Assert.Equal(6, c.Productos);
            Assert.Equal(3, c.Cruzados);
            Assert.Equal(3, c.SinCruzar);
            Assert.Equal(1, c.NombresAmbiguos);
            Assert.Equal(2, c.RepetidosFilas);
            Assert.Equal(1, c.RepetidosNombres);
            Assert.Contains(c.EjemplosRepetidos, e => e.Contains("DURACELL BATERIA AAA UND"));
            Assert.Contains(c.EjemplosRepetidos, e => e.Contains("2"));
        }

        [Fact]
        public async Task B2_Vista_precio_de_venta_cruza_mirador_con_los_codigos()
        {
            var r = await Las3();
            var v = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "catarata", "mirador");

            var f1 = FilaDe(v, "001");
            Assert.Equal(820m, Valor(v, f1, "mirador"));
            Assert.Equal(800m, Valor(v, f1, "bomba"));
            Assert.Equal(850m, Valor(v, f1, "catarata"));

            Assert.Equal(910m, Valor(v, FilaDe(v, "002"), "mirador"));

            var f4 = FilaDe(v, "004");
            Assert.Equal(1250m, Valor(v, f4, "mirador"));
            Assert.Equal(1200m, Valor(v, f4, "catarata"));
        }

        [Fact]
        public async Task B3_No_entran_totales_al_pie_ni_repetidos()
        {
            var r = await Las3();
            var v = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "catarata", "mirador");

            Assert.DoesNotContain(v.Filas, f => f.Valores.Any(x => x == 99999m || x == 88888m));
            Assert.DoesNotContain(v.Filas, f => f.Descripcion.Contains("DURACELL", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public async Task B4_Lo_que_no_cruza_sale_en_filas_propias_sin_codigo()
        {
            var r = await Las3();
            var v = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "catarata", "mirador");

            var coca = Assert.Single(v.Filas, f => f.Codigo == SinCodigo &&
                f.Descripcion.Contains("COCA", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(1, coca.Presencia);
            Assert.Equal(1100m, Valor(v, coca, "mirador"));

            Assert.Null(Valor(v, FilaDe(v, "005"), "mirador"));
            Assert.Null(Valor(v, FilaDe(v, "006"), "mirador"));

            var pan = Assert.Single(v.Filas, f => f.Codigo == SinCodigo &&
                f.Descripcion.Contains("PAN BLANCO GRANDE", StringComparison.OrdinalIgnoreCase));
            Assert.Equal(750m, Valor(v, pan, "mirador"));
            Assert.Null(Valor(v, FilaDe(v, "003"), "mirador"));
        }

        [Fact]
        public async Task B5_El_mensaje_del_diagnostico_trae_los_conteos_y_el_riesgo()
        {
            var r = await Las3();
            var m = DeMirador(r).Mensaje;

            Assert.Contains("3 de 6", m);
            Assert.Contains("3 quedaron sin cruzar", m);
            Assert.Contains("incongruencias", m);
        }

        [Fact]
        public async Task B6_El_aviso_del_cruce_describe_la_lista_y_los_conteos()
        {
            var r = await Las3();
            var lineas = AvisoCruce.Lineas(r.Datos.CrucesPorDescripcion);

            Assert.Contains("mirador", lineas[0]);
            Assert.Contains("por descripción y no por código", lineas[0]);
            Assert.Contains(lineas, l => l.Contains("incongruencias"));
            Assert.Contains(lineas, l => l.Contains("3 de 6") && l.Contains("DURACELL BATERIA AAA UND"));
        }

        [Fact]
        public async Task B7_Las_listas_con_codigo_siguen_cruzando_por_codigo()
        {
            var r = await CargarPrecios(Bomba(), Catarata());
            Assert.False(r.Datos.HayCrucePorDescripcion);

            var v = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "catarata");
            Assert.Equal(6, v.Filas.Count);
            Assert.Equal(new[] { "001", "002", "003", "004", "005", "006" },
                v.Filas.Select(f => f.Codigo).OrderBy(c => c).ToArray());

            var f1 = FilaDe(v, "001");
            Assert.Equal(800m, Valor(v, f1, "bomba"));
            Assert.Equal(850m, Valor(v, f1, "catarata"));
        }

        [Fact]
        public async Task B8_El_cruce_por_descripcion_no_altera_el_cruce_por_codigo()
        {
            var sin = await CargarPrecios(Bomba(), Catarata());
            var con = await Las3();

            foreach (var metrica in new[] { MetricaPrecio.PrecioVenta, MetricaPrecio.Costo, MetricaPrecio.Utilidad })
            {
                var a = Comparar(sin.Datos, metrica, "bomba", "catarata");
                var b = Comparar(con.Datos, metrica, "bomba", "catarata", "mirador");
                foreach (var codigo in new[] { "001", "002", "003", "004", "005", "006" })
                {
                    var fa = FilaDe(a, codigo);
                    var fb = FilaDe(b, codigo);
                    Assert.Equal(Valor(a, fa, "bomba"), Valor(b, fb, "bomba"));
                    Assert.Equal(Valor(a, fa, "catarata"), Valor(b, fb, "catarata"));
                }
            }
        }

        [Fact]
        public async Task B9_Un_nombre_con_dos_codigos_distintos_no_se_cruza_con_ninguno()
        {
            string b = Escribir(Carpeta, "precios bomba.xlsx", EncabezadoEstandar,
                Fila("001", "CAFE NEGRO", 200, 13, 50, 800));
            string c = Escribir(Carpeta, "precios catarata.xlsx", EncabezadoEstandar,
                Fila("001", "CAFE NEGRO GRANDE", 210, 13, 50, 850));
            string m = ListaSinCodigo("precios mirador.xlsx", "CAFE NEGRO", "CAFE NEGRO GRANDE");

            var r = await CargarPrecios(b, c, m);
            var cruce = Assert.Single(r.Datos.CrucesPorDescripcion);

            Assert.Equal(0, cruce.Cruzados);
            Assert.Equal(2, cruce.NombresAmbiguos);

            var v = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "catarata", "mirador");
            Assert.Null(Valor(v, FilaDe(v, "001"), "mirador"));
        }

        [Fact]
        public async Task B10_Repetidos_que_solo_difieren_en_tildes_o_mayusculas_quedan_fuera()
        {
            string b = Escribir(Carpeta, "precios bomba.xlsx", EncabezadoEstandar,
                Fila("001", "CAFE NEGRO", 200, 13, 50, 800));
            string m = ListaSinCodigo("precios mirador.xlsx", "Café negro", "CAFE NEGRO", "PRODUCTO UNICO");

            var r = await CargarPrecios(b, m);
            var cruce = Assert.Single(r.Datos.CrucesPorDescripcion);

            Assert.Equal(2, cruce.RepetidosFilas);
            Assert.Equal(1, cruce.RepetidosNombres);
            Assert.Equal(0, cruce.Cruzados);

            var v = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "mirador");
            Assert.Null(Valor(v, FilaDe(v, "001"), "mirador"));
        }

        // ---------- C. «Costo I.V.A.» no entra como costo ----------

        [Fact]
        public async Task C1_El_costo_con_IVA_de_mirador_no_se_toma_como_costo()
        {
            var r = await Las3();
            int id = r.Datos.Sucursales.IdExistente("mirador");
            Assert.True(id >= 0);

            Assert.False(r.Datos.Tiene(id, MetricaPrecio.Costo));
            Assert.False(r.Datos.Tiene(id, MetricaPrecio.Utilidad));
            Assert.True(r.Datos.Tiene(id, MetricaPrecio.PrecioVenta));

            var propios = r.Datos.Filas.Where(f => f.SucursalId == id).ToList();
            Assert.NotEmpty(propios);
            Assert.All(propios, f => Assert.Equal(0m, f.Costo));
        }

        [Fact]
        public async Task C2_En_la_vista_costo_mirador_sale_sin_dato_y_no_como_la_mejor()
        {
            var r = await Las3();
            var v = Comparar(r.Datos, MetricaPrecio.Costo, "bomba", "catarata", "mirador");
            var f = FilaDe(v, "001");
            int i = v.Sucursales.IndexOf("mirador");

            Assert.Equal(FilaComparativa.TextoSinDato, f.Textos[i]);
            Assert.Null(f.Valores[i]);
            Assert.Equal("bomba", f.Mejor);
            Assert.Equal("catarata", f.Peor);
        }

        [Fact]
        public async Task C3_En_la_vista_utilidad_un_producto_con_una_sola_sucursal_no_es_comparable()
        {
            var r = await Las3();
            var v = Comparar(r.Datos, MetricaPrecio.Utilidad, "bomba", "catarata", "mirador");
            var f = FilaDe(v, "004");

            Assert.False(f.ComparablePrincipal);
            Assert.Equal("-", f.DiferenciaTexto);
            Assert.Equal("-", f.Mejor);
        }

        [Fact]
        public async Task C4_Los_comparables_cuentan_solo_sucursales_con_dato()
        {
            var r = await Las3();
            var costo = Comparar(r.Datos, MetricaPrecio.Costo, "bomba", "catarata", "mirador");
            var venta = Comparar(r.Datos, MetricaPrecio.PrecioVenta, "bomba", "catarata", "mirador");

            Assert.Equal(2, costo.Comparables);
            Assert.Equal(3, venta.Comparables);
        }

        // ---------- E. Lista estándar sigue cargando ----------

        [Fact]
        public async Task E1_Una_lista_estandar_carga_completa_y_sin_cruces()
        {
            var r = await CargarPrecios(Bomba());
            var d = Assert.Single(r.Diagnosticos);

            Assert.Equal(EstadoCarga.Cargado, d.Estado);
            Assert.Equal(4, d.FilasLeidas);
            Assert.False(r.Datos.HayCrucePorDescripcion);
            Assert.Empty(r.Datos.CrucesPorDescripcion);

            int id = r.Datos.Sucursales.IdExistente("bomba");
            Assert.True(id >= 0);
            Assert.True(r.Datos.Tiene(id, MetricaPrecio.Costo));
            Assert.True(r.Datos.Tiene(id, MetricaPrecio.Utilidad));
            Assert.True(r.Datos.Tiene(id, MetricaPrecio.PrecioVenta));
        }

        [Fact]
        public async Task E2_Una_lista_estandar_sin_impuesto_carga_y_lo_menciona()
        {
            string ruta = Escribir(Carpeta, "precios bomba.xlsx",
                Fila("Cód. Artículo", "Descripción", "Precio costo", "Porc. utilidad - artículo", "Precio IVI - artículo"),
                Fila("001", "CAFE NEGRO", 200, 50, 800));

            var r = await CargarPrecios(ruta);
            var d = Assert.Single(r.Diagnosticos);

            Assert.Equal(EstadoCarga.Cargado, d.Estado);
            Assert.Contains(d.Detalles, l => l.Contains("Imp. ventas"));
        }
    }
}
