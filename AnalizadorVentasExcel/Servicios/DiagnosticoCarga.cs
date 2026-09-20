using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using AnalizadorVentasExcel.Modelos;
using ExcelDataReader.Exceptions;

namespace AnalizadorVentasExcel.Servicios
{
    /// <summary>Qué es en realidad un archivo, más allá de lo que diga su extensión.</summary>
    public enum TipoReal
    {
        Xlsx,
        Xls,
        Html,
        TextoCsv,
        Vacio,
        Otro
    }

    /// <summary>
    /// Traduce las excepciones de ExcelDataReader (y las de E/S de .NET) a un
    /// <see cref="DiagnosticoArchivo"/> que un usuario sin conocimientos técnicos pueda
    /// entender. El caso más común no es un archivo corrupto sino uno que nunca fue Excel:
    /// la caja registradora exporta HTML o CSV con extensión .xls, Excel lo abre igual (es
    /// tolerante) pero el lector no, así que hay que olfatear el contenido real para dar
    /// un mensaje útil en vez de repetir el texto crudo de la excepción.
    /// </summary>
    public static class DiagnosticoCarga
    {
        /// <summary>
        /// Adivina el tipo real de un archivo mirando sus primeros bytes, sin fiarse de la
        /// extensión. <see cref="FileShare.ReadWrite"/> porque a veces se llama mientras el
        /// archivo sigue abierto en Excel.
        /// </summary>
        public static TipoReal OlfatearTipo(string ruta)
        {
            byte[] buffer = new byte[512];
            int leidos;
            try
            {
                using var fs = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                leidos = fs.Read(buffer, 0, buffer.Length);
            }
            catch
            {
                return TipoReal.Otro;
            }

            if (leidos == 0) return TipoReal.Vacio;

            if (leidos >= 4 && buffer[0] == 0x50 && buffer[1] == 0x4B && buffer[2] == 0x03 && buffer[3] == 0x04)
                return TipoReal.Xlsx;

            if (leidos >= 8 && buffer[0] == 0xD0 && buffer[1] == 0xCF && buffer[2] == 0x11 && buffer[3] == 0xE0 &&
                buffer[4] == 0xA1 && buffer[5] == 0xB1 && buffer[6] == 0x1A && buffer[7] == 0xE1)
                return TipoReal.Xls;

            // El HTML "guardado como Excel" es el caso más común: se intenta decodificar como
            // texto y, si eso falla (bytes que no son UTF-8 válido), simplemente no es HTML.
            try
            {
                string texto = new UTF8Encoding(false, true).GetString(buffer, 0, leidos);
                string sinEspacios = texto.TrimStart('﻿', ' ', '\t', '\r', '\n');
                if (sinEspacios.StartsWith("<", StringComparison.Ordinal))
                {
                    string enMinusculas = sinEspacios.ToLowerInvariant();
                    if (enMinusculas.Contains("<html") || enMinusculas.Contains("<table") ||
                        enMinusculas.Contains("<!doctype") || enMinusculas.Contains("<?xml"))
                        return TipoReal.Html;
                }
            }
            catch (DecoderFallbackException)
            {
                // no es texto UTF-8 válido: sigue por las demás reglas
            }

            int imprimibles = 0;
            bool tieneSeparadorDeTexto = false;
            for (int i = 0; i < leidos; i++)
            {
                byte b = buffer[i];
                if ((b >= 0x20 && b <= 0x7E) || b == 0x09 || b == 0x0D || b == 0x0A) imprimibles++;
                if (b == (byte)';' || b == (byte)',' || b == (byte)'\t') tieneSeparadorDeTexto = true;
            }
            if ((double)imprimibles / leidos > 0.9 && tieneSeparadorDeTexto) return TipoReal.TextoCsv;

            return TipoReal.Otro;
        }

        /// <summary>
        /// Convierte una excepción de carga en un diagnóstico. Desenvuelve las envolturas
        /// que no aportan nada al usuario (<see cref="AggregateException"/> de tareas en
        /// paralelo, <see cref="TargetInvocationException"/> de invocación reflexiva) antes
        /// de clasificar la causa real.
        /// </summary>
        public static DiagnosticoArchivo Clasificar(Exception ex, string ruta, string sucursal)
        {
            var diagnostico = new DiagnosticoArchivo(ruta, sucursal);
            Exception causa = Desenvolver(ex);

            switch (causa)
            {
                case InvalidPasswordException:
                    diagnostico.Registrar(ProblemaCarga.ConContrasena, "El archivo tiene contraseña.");
                    break;

                case HeaderException:
                    ClasificarPorContenidoReal(diagnostico, ruta);
                    break;

                case CompoundDocumentException:
                case ExcelReaderException:
                case InvalidDataException:
                case EndOfStreamException:
                    diagnostico.Registrar(ProblemaCarga.ExcelDanado,
                        $"El archivo está dañado o incompleto. Detalle técnico: {causa.Message}");
                    break;

                case IOException io when io.HResult == unchecked((int)0x80070020) || io.HResult == unchecked((int)0x80070021):
                    diagnostico.Registrar(ProblemaCarga.ArchivoEnUso, "El archivo está abierto en otro programa, seguramente Excel.");
                    break;

                case UnauthorizedAccessException:
                    diagnostico.Registrar(ProblemaCarga.SinPermisos, "No hay permiso para leer el archivo.");
                    break;

                case FileNotFoundException:
                case DirectoryNotFoundException:
                    diagnostico.Registrar(ProblemaCarga.ArchivoNoEncontrado, "El archivo desapareció de la carpeta mientras se cargaba.");
                    break;

                case IOException io:
                    diagnostico.Registrar(ProblemaCarga.ExcelDanado,
                        $"No se pudo leer el archivo completo. Detalle técnico: {io.Message}");
                    break;

                default:
                    diagnostico.Registrar(ProblemaCarga.Desconocido,
                        $"Error inesperado: {causa.GetType().Name}: {causa.Message}");
                    break;
            }

            return diagnostico;
        }

        /// <summary>
        /// "Invalid file signature." de ExcelDataReader sólo dice que no reconoció el
        /// formato; olfatea el contenido para poder decirle al usuario qué es en realidad.
        /// </summary>
        private static void ClasificarPorContenidoReal(DiagnosticoArchivo diagnostico, string ruta)
        {
            string extension = Path.GetExtension(ruta);
            switch (OlfatearTipo(ruta))
            {
                case TipoReal.Html:
                    diagnostico.Registrar(ProblemaCarga.NoEsExcel,
                        $"No es un libro de Excel: es una página HTML guardada con extensión {extension}.");
                    break;
                case TipoReal.TextoCsv:
                    diagnostico.Registrar(ProblemaCarga.NoEsExcel,
                        $"No es un libro de Excel: es un archivo de texto (CSV) con extensión {extension}.");
                    break;
                case TipoReal.Vacio:
                    diagnostico.Registrar(ProblemaCarga.ArchivoVacio, "El archivo está vacío (0 bytes).");
                    break;
                case TipoReal.Xlsx:
                case TipoReal.Xls:
                    diagnostico.Registrar(ProblemaCarga.ExcelDanado,
                        "El archivo parece un Excel pero no se pudo leer: está dañado o incompleto.");
                    break;
                default:
                    diagnostico.Registrar(ProblemaCarga.NoEsExcel, "No es un libro de Excel: el contenido no se reconoce.");
                    break;
            }
        }

        private static Exception Desenvolver(Exception ex)
        {
            while (true)
            {
                if (ex is AggregateException ae && ae.InnerException != null) { ex = ae.InnerException; continue; }
                if (ex is TargetInvocationException tie && tie.InnerException != null) { ex = tie.InnerException; continue; }
                return ex;
            }
        }

        /// <summary>
        /// Detecta sucursales repetidas entre los archivos que sí se cargaron: sale del
        /// nombre del archivo, así que dos archivos con nombres parecidos terminan
        /// mezclando sus filas en una sola sucursal sin que se note en la tabla.
        /// </summary>
        public static void MarcarDuplicadas(IReadOnlyList<DiagnosticoArchivo> diagnosticos)
        {
            var grupos = diagnosticos
                .Where(d => d.Cargado)
                .GroupBy(d => d.Sucursal, StringComparer.OrdinalIgnoreCase);

            foreach (var grupo in grupos)
            {
                var miembros = grupo.ToList();
                if (miembros.Count <= 1) continue;

                foreach (var actual in miembros)
                {
                    string otros = string.Join(", ", miembros.Where(m => m != actual).Select(m => m.Archivo));
                    actual.Registrar(ProblemaCarga.SucursalDuplicada,
                        $"«{actual.Sucursal}» aparece también en {otros}: sus filas se mezclan en una sola sucursal.");
                }
            }
        }

        /// <summary>Texto plano para copiar al portapapeles: el mismo diagnóstico que muestra la ventana, sin formato.</summary>
        public static string Informe(IReadOnlyList<DiagnosticoArchivo> diagnosticos)
        {
            int cargados = diagnosticos.Count(d => d.Estado == EstadoCarga.Cargado);
            int conAvisos = diagnosticos.Count(d => d.Estado == EstadoCarga.ConAvisos);
            int noCargados = diagnosticos.Count(d => d.Estado == EstadoCarga.NoCargado);

            var sb = new StringBuilder();
            sb.AppendLine($"Diagnóstico de carga — {DateTime.Now:dd/MM/yyyy HH:mm}");
            sb.AppendLine($"{diagnosticos.Count} archivos: {cargados} cargados, {conAvisos} con avisos, {noCargados} no cargados");
            sb.AppendLine();

            for (int i = 0; i < diagnosticos.Count; i++)
            {
                var d = diagnosticos[i];
                string sufijoFilas = d.Cargado ? $" — {d.FilasTexto} filas" : string.Empty;

                if (d.Estado == EstadoCarga.Cargado)
                {
                    sb.AppendLine($"{d.Icono} {d.Archivo} ({d.Sucursal}){sufijoFilas}");
                }
                else
                {
                    sb.AppendLine($"{d.Icono} {d.Archivo} ({d.Sucursal}){sufijoFilas}");
                    sb.AppendLine($"   Qué pasó: {d.Mensaje}");
                    if (!string.IsNullOrEmpty(d.Sugerencia))
                        sb.AppendLine($"   Qué hacer: {d.Sugerencia}");

                    // El mensaje principal ya se imprimió como "Qué pasó": en la lista van
                    // sólo los demás avisos, para no leer lo mismo dos veces.
                    foreach (string detalle in d.Detalles)
                        if (detalle != d.Mensaje) sb.AppendLine($"   - {detalle}");
                }

                if (i < diagnosticos.Count - 1) sb.AppendLine();
            }

            return sb.ToString();
        }
    }
}
