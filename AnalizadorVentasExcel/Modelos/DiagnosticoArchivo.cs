using System;
using System.Collections.Generic;
using System.IO;

namespace AnalizadorVentasExcel.Modelos
{
    public enum EstadoCarga
    {
        Cargado = 0,
        ConAvisos = 1,
        NoCargado = 2
    }

    /// <summary>
    /// Qué le pasó a un archivo al cargarlo. El orden importa: está de menos a más grave,
    /// y un archivo que acumula varios problemas se queda con el mayor. Todo lo que está
    /// desde <see cref="SinFilasUtiles"/> en adelante impide cargar el archivo.
    /// </summary>
    public enum ProblemaCarga
    {
        Ninguno = 0,

        // --- Avisos: el archivo se cargó, pero conviene saberlo ---
        FilasDescartadas = 10,
        HojaEquivocada = 11,
        EncabezadoIncompleto = 12,
        SucursalDuplicada = 13,

        // --- Errores: el archivo no se cargó ---
        SinFilasUtiles = 20,
        SinEncabezado = 21,
        ReporteDelOtroSistema = 22,
        ArchivoVacio = 23,
        NoEsExcel = 24,
        ExcelDanado = 25,
        ConContrasena = 26,
        ArchivoEnUso = 27,
        SinPermisos = 28,
        ArchivoNoEncontrado = 29,
        Desconocido = 30
    }

    /// <summary>
    /// Veredicto de un archivo tras intentar cargarlo: qué pasó, con qué evidencia y qué
    /// hacer. Reemplaza al mensaje crudo de la excepción y, sobre todo, al silencio: antes
    /// un archivo sin encabezado devolvía cero filas y no dejaba rastro.
    /// </summary>
    public sealed class DiagnosticoArchivo
    {
        public DiagnosticoArchivo(string ruta, string sucursal)
        {
            Ruta = ruta;
            Archivo = Path.GetFileName(ruta);
            Sucursal = sucursal;
        }

        public string Ruta { get; }
        public string Archivo { get; }
        public string Sucursal { get; set; }

        /// <summary>El problema más grave registrado; <see cref="ProblemaCarga.Ninguno"/> si todo salió bien.</summary>
        public ProblemaCarga Problema { get; private set; } = ProblemaCarga.Ninguno;

        /// <summary>Qué pasó, en una frase y con la evidencia concreta (columnas, conteos, hoja).</summary>
        public string Mensaje { get; private set; } = "Cargado sin problemas.";

        /// <summary>Qué hacer al respecto. Vacío cuando es sólo informativo.</summary>
        public string Sugerencia { get; private set; } = string.Empty;

        public int FilasLeidas { get; set; }

        /// <summary>Todo lo registrado, en orden, incluidos los avisos que no llegaron a ser el principal.</summary>
        public List<string> Detalles { get; } = new();

        public EstadoCarga Estado =>
            Problema == ProblemaCarga.Ninguno ? EstadoCarga.Cargado
            : EsError(Problema) ? EstadoCarga.NoCargado
            : EstadoCarga.ConAvisos;

        public bool Cargado => Estado != EstadoCarga.NoCargado;

        public static bool EsError(ProblemaCarga p) => p >= ProblemaCarga.SinFilasUtiles;

        /// <summary>
        /// Anota un problema. Siempre queda en <see cref="Detalles"/>; pasa a ser el principal
        /// sólo si es más grave que el que ya había. Si no se da sugerencia, se toma la del
        /// catálogo.
        /// </summary>
        public void Registrar(ProblemaCarga problema, string mensaje, string? sugerencia = null)
        {
            if (problema == ProblemaCarga.Ninguno) return;

            Detalles.Add(mensaje);
            if (problema <= Problema) return;

            Problema = problema;
            Mensaje = mensaje;
            Sugerencia = sugerencia ?? CatalogoProblemas.De(problema).Sugerencia;
        }

        public string Icono => Estado switch
        {
            EstadoCarga.Cargado => "✅",
            EstadoCarga.ConAvisos => "⚠",
            _ => "❌"
        };

        public string EstadoTexto => Estado switch
        {
            EstadoCarga.Cargado => "Cargado",
            EstadoCarga.ConAvisos => "Con avisos",
            _ => "No cargado"
        };

        public string FilasTexto => Cargado ? FilasLeidas.ToString("N0", ResumenDinamico.FormatoCR) : "-";
    }

    /// <summary>Texto fijo de cada problema: título para la guía, explicación genérica y qué hacer.</summary>
    public sealed class EntradaCatalogo
    {
        public string Titulo { get; init; } = string.Empty;
        public string Explicacion { get; init; } = string.Empty;
        public string Sugerencia { get; init; } = string.Empty;
    }

    /// <summary>
    /// Una sola fuente para los textos: de acá salen la sugerencia que muestra la ventana de
    /// diagnóstico y la sección de la guía. Así nunca dicen cosas distintas. El mensaje
    /// concreto (con la evidencia: qué columna falta, cuántas filas) lo arma quien detecta
    /// el problema; acá va lo que no cambia.
    /// </summary>
    public static class CatalogoProblemas
    {
        private static readonly Dictionary<ProblemaCarga, EntradaCatalogo> Entradas = new()
        {
            [ProblemaCarga.ReporteDelOtroSistema] = new()
            {
                Titulo = "Es un reporte del otro sistema",
                Explicacion = "El archivo tiene el encabezado de un reporte de precios y se cargó en la ventana de " +
                              "ventas, o al revés. Los dos sistemas leen \"cualquier Excel de la carpeta\", así que " +
                              "es el error más fácil de cometer. El mensaje dice de cuál de los dos es.",
                Sugerencia = "Cargalo en la ventana que corresponde: las ventas en la principal, las listas de " +
                             "precios en ⚖️ Comparativa de Precios."
            },
            [ProblemaCarga.SinEncabezado] = new()
            {
                Titulo = "No se encontró el encabezado",
                Explicacion = "En las primeras 20 filas con datos no hay ninguna que tenga todas las columnas " +
                              "obligatorias. El mensaje señala la fila que más se parece y qué columna le falta.",
                Sugerencia = "Revisá que el reporte se haya exportado con todas las columnas y que el encabezado " +
                             "esté al principio de la hoja, no después de un título largo."
            },
            [ProblemaCarga.EncabezadoIncompleto] = new()
            {
                Titulo = "Falta una columna opcional",
                Explicacion = "El archivo se carga igual, pero sin esa columna: por ejemplo, sin \"% Utilidad\" " +
                              "todo entra con utilidad 0, y sin \"Proveedor\" todo queda como \"General\".",
                Sugerencia = "Si necesitás ese dato, volvé a exportar el reporte con esa columna incluida."
            },
            [ProblemaCarga.NoEsExcel] = new()
            {
                Titulo = "No es un libro de Excel",
                Explicacion = "La extensión dice .xls o .xlsx pero el contenido es otra cosa. Lo más común: la " +
                              "caja exporta una página HTML y la guarda como .xls. Excel la abre, pero el " +
                              "programa no. El mensaje dice qué es en realidad.",
                Sugerencia = "Abrilo en Excel y guardalo como \"Libro de Excel (.xlsx)\", o exportalo de nuevo " +
                             "desde la caja eligiendo formato Excel."
            },
            [ProblemaCarga.ExcelDanado] = new()
            {
                Titulo = "El archivo está dañado o incompleto",
                Explicacion = "Es un Excel, pero está cortado o corrupto y no se puede leer entero.",
                Sugerencia = "Volvé a exportarlo. Si lo copiaste por USB, WhatsApp o correo, puede haberse " +
                             "cortado en el camino."
            },
            [ProblemaCarga.ConContrasena] = new()
            {
                Titulo = "El archivo tiene contraseña",
                Explicacion = "Está protegido y el programa no puede abrirlo.",
                Sugerencia = "Quitale la contraseña en Excel (Archivo → Información → Proteger libro) y volvé " +
                             "a cargar la carpeta."
            },
            [ProblemaCarga.ArchivoEnUso] = new()
            {
                Titulo = "El archivo está abierto en otro programa",
                Explicacion = "Casi siempre está abierto en Excel, que lo bloquea mientras lo tiene en pantalla.",
                Sugerencia = "Cerralo y volvé a cargar la carpeta."
            },
            [ProblemaCarga.SinPermisos] = new()
            {
                Titulo = "No hay permiso para leer el archivo",
                Explicacion = "Windows no deja abrirlo: suele pasar con carpetas de red o de otro usuario.",
                Sugerencia = "Copiá la carpeta al Escritorio y cargala desde ahí."
            },
            [ProblemaCarga.ArchivoNoEncontrado] = new()
            {
                Titulo = "El archivo desapareció",
                Explicacion = "Estaba en la carpeta cuando se empezó a cargar y ya no está.",
                Sugerencia = "Volvé a cargar la carpeta."
            },
            [ProblemaCarga.ArchivoVacio] = new()
            {
                Titulo = "El archivo está vacío",
                Explicacion = "No tiene hojas, o no tiene ninguna fila con datos.",
                Sugerencia = "La exportación salió sin datos: repetila revisando el filtro de fechas."
            },
            [ProblemaCarga.SinFilasUtiles] = new()
            {
                Titulo = "El encabezado está bien, pero ninguna fila sirvió",
                Explicacion = "Todas las filas se descartaron: sin importe, sin código, o eran totales. El " +
                              "mensaje trae el desglose.",
                Sugerencia = "El reporte no trae movimientos para ese periodo, o el filtro de la exportación " +
                             "quedó vacío. Revisá las fechas y volvé a exportar."
            },
            [ProblemaCarga.FilasDescartadas] = new()
            {
                Titulo = "Se descartaron algunas filas",
                Explicacion = "El archivo se cargó, pero algunas filas no tenían con qué (sin importe, sin " +
                              "código, un total al pie). Es normal que sean unas pocas.",
                Sugerencia = string.Empty
            },
            [ProblemaCarga.HojaEquivocada] = new()
            {
                Titulo = "Los datos estaban en otra hoja",
                Explicacion = "La primera hoja del libro no tenía el encabezado, pero otra sí. El programa " +
                              "usó esa y avisa, por si no era la que querías.",
                Sugerencia = string.Empty
            },
            [ProblemaCarga.SucursalDuplicada] = new()
            {
                Titulo = "Dos archivos con la misma sucursal",
                Explicacion = "El nombre de la sucursal sale del nombre del archivo, quitándole la fecha y " +
                              "palabras como \"precios\". Dos archivos quedaron con el mismo nombre y sus " +
                              "filas se mezclan en una sola sucursal.",
                Sugerencia = "Renombrá uno de los dos archivos, o sacá de la carpeta el que sobra."
            },
            [ProblemaCarga.Desconocido] = new()
            {
                Titulo = "Error inesperado",
                Explicacion = "Falló algo que el programa no sabe explicar. El mensaje trae el texto técnico.",
                Sugerencia = "Copiá el informe con el botón de la ventana de diagnóstico y mandalo para revisarlo."
            }
        };

        public static EntradaCatalogo De(ProblemaCarga problema)
            => Entradas.TryGetValue(problema, out var e) ? e : new EntradaCatalogo();

        /// <summary>Los problemas en el orden en que conviene leerlos en la guía: primero los más comunes.</summary>
        public static IReadOnlyList<ProblemaCarga> OrdenParaLaGuia { get; } = new[]
        {
            ProblemaCarga.ReporteDelOtroSistema,
            ProblemaCarga.SinEncabezado,
            ProblemaCarga.ArchivoEnUso,
            ProblemaCarga.NoEsExcel,
            ProblemaCarga.FilasDescartadas,
            ProblemaCarga.SinFilasUtiles,
            ProblemaCarga.EncabezadoIncompleto,
            ProblemaCarga.HojaEquivocada,
            ProblemaCarga.SucursalDuplicada,
            ProblemaCarga.ArchivoVacio,
            ProblemaCarga.ExcelDanado,
            ProblemaCarga.ConContrasena,
            ProblemaCarga.SinPermisos,
            ProblemaCarga.ArchivoNoEncontrado,
            ProblemaCarga.Desconocido
        };
    }
}
