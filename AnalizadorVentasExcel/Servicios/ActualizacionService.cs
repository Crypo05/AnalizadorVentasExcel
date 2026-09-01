using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AnalizadorVentasExcel.Servicios
{
    public sealed class InfoActualizacion
    {
        public Version Version { get; init; } = new(0, 0, 0);
        public string Etiqueta { get; init; } = string.Empty;
        public string UrlDescarga { get; init; } = string.Empty;
        public long Tamano { get; init; }
        public string Notas { get; init; } = string.Empty;
        public string UrlPagina { get; init; } = string.Empty;
    }

    public enum EstadoActualizacion { HayNueva, AlDia, LocalMasNueva }

    public sealed class ResultadoConsulta
    {
        public EstadoActualizacion Estado { get; init; }
        public InfoActualizacion? Info { get; init; }
    }

    /// <summary>
    /// Actualización desde los Releases de GitHub.
    ///
    /// Se consulta la API de releases y no el Version.txt del repositorio: ese archivo
    /// se quedó desactualizado (marca 1.0.2 cuando el último publicado es 1.3.0), mientras
    /// que los releases son lo que realmente se distribuye.
    ///
    /// El reemplazo aprovecha que Windows permite renombrar un ejecutable en uso: se mueve
    /// el actual a «.old», se pone el nuevo en su lugar y se reinicia. El «.old» lo borra
    /// MainWindow al arrancar la siguiente vez.
    /// </summary>
    public sealed class ActualizacionService
    {
        public const string Repositorio = "Crypo05/AnalizadorVentasExcel";
        private const string NombreAsset = "AnalizadorVentasExcel.exe";
        private const long TamanoMinimoRazonable = 1_000_000; // por debajo de esto no es la app

        private static readonly HttpClient Http = CrearCliente();

        private static HttpClient CrearCliente()
        {
            var c = new HttpClient { Timeout = TimeSpan.FromMinutes(30) };
            // GitHub rechaza las peticiones sin User-Agent.
            c.DefaultRequestHeaders.UserAgent.ParseAdd("AnalizadorVentasExcel-Updater");
            c.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
            return c;
        }

        /// <summary>
        /// Las etiquetas del repositorio mezclan estilos ("1.3.0" y "v1.2.0"), así que se
        /// normalizan antes de comparar.
        /// </summary>
        public static Version? NormalizarVersion(string? etiqueta)
        {
            if (string.IsNullOrWhiteSpace(etiqueta)) return null;

            string t = etiqueta.Trim();
            if (t.Length > 0 && (t[0] == 'v' || t[0] == 'V')) t = t.Substring(1);

            // Se corta en el primer carácter que no sea dígito o punto ("1.3.0-beta" -> "1.3.0").
            int fin = 0;
            while (fin < t.Length && (char.IsDigit(t[fin]) || t[fin] == '.')) fin++;
            t = t.Substring(0, fin).Trim('.');
            if (t.Length == 0) return null;

            // Version exige al menos mayor.menor.
            if (!t.Contains('.')) t += ".0";
            return Version.TryParse(t, out var v) ? v : null;
        }

        public async Task<ResultadoConsulta> ConsultarAsync(string versionActual, CancellationToken ct = default)
        {
            string url = $"https://api.github.com/repos/{Repositorio}/releases/latest";

            using var respuesta = await Http.GetAsync(url, ct).ConfigureAwait(false);
            if (!respuesta.IsSuccessStatusCode)
                throw new InvalidOperationException(
                    $"GitHub respondió {(int)respuesta.StatusCode} ({respuesta.ReasonPhrase}). " +
                    "Puede ser falta de conexión o el límite de consultas de GitHub.");

            string json = await respuesta.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            using var doc = JsonDocument.Parse(json);
            var raiz = doc.RootElement;

            string etiqueta = Texto(raiz, "tag_name");
            var version = NormalizarVersion(etiqueta)
                ?? throw new InvalidOperationException($"No se pudo interpretar la versión publicada ('{etiqueta}').");

            string urlDescarga = string.Empty;
            long tamano = 0;
            if (raiz.TryGetProperty("assets", out var assets) && assets.ValueKind == JsonValueKind.Array)
            {
                foreach (var a in assets.EnumerateArray())
                {
                    if (!string.Equals(Texto(a, "name"), NombreAsset, StringComparison.OrdinalIgnoreCase)) continue;
                    urlDescarga = Texto(a, "browser_download_url");
                    tamano = a.TryGetProperty("size", out var s) && s.TryGetInt64(out long v) ? v : 0;
                    break;
                }
            }

            var info = new InfoActualizacion
            {
                Version = version,
                Etiqueta = etiqueta,
                UrlDescarga = urlDescarga,
                Tamano = tamano,
                Notas = Texto(raiz, "body"),
                UrlPagina = Texto(raiz, "html_url")
            };

            var actual = NormalizarVersion(versionActual) ?? new Version(0, 0, 0);
            var estado = version > actual ? EstadoActualizacion.HayNueva
                       : version == actual ? EstadoActualizacion.AlDia
                       : EstadoActualizacion.LocalMasNueva;

            return new ResultadoConsulta { Estado = estado, Info = info };
        }

        private static string Texto(JsonElement e, string propiedad)
            => e.TryGetProperty(propiedad, out var v) && v.ValueKind == JsonValueKind.String
               ? (v.GetString() ?? string.Empty) : string.Empty;

        /// <summary>Descarga en streaming informando del avance (bytes recibidos, total).</summary>
        public async Task DescargarAsync(InfoActualizacion info, string destino,
                                         IProgress<(long recibidos, long total)>? progreso = null,
                                         CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(info.UrlDescarga))
                throw new InvalidOperationException(
                    $"El release {info.Etiqueta} no incluye el archivo {NombreAsset}.");

            using var respuesta = await Http
                .GetAsync(info.UrlDescarga, HttpCompletionOption.ResponseHeadersRead, ct)
                .ConfigureAwait(false);
            respuesta.EnsureSuccessStatusCode();

            long total = respuesta.Content.Headers.ContentLength ?? info.Tamano;

            using (var origen = await respuesta.Content.ReadAsStreamAsync(ct).ConfigureAwait(false))
            using (var salida = new FileStream(destino, FileMode.Create, FileAccess.Write, FileShare.None,
                                               bufferSize: 1 << 16, useAsync: true))
            {
                var buffer = new byte[1 << 16];
                long recibidos = 0;
                int leidos;
                while ((leidos = await origen.ReadAsync(buffer, ct).ConfigureAwait(false)) > 0)
                {
                    await salida.WriteAsync(buffer.AsMemory(0, leidos), ct).ConfigureAwait(false);
                    recibidos += leidos;
                    progreso?.Report((recibidos, total));
                }
            }

            ValidarEjecutable(destino, info.Tamano);
        }

        /// <summary>
        /// Comprueba que lo descargado sea de verdad el ejecutable y no una página de error
        /// ni una descarga a medias.
        /// </summary>
        public static void ValidarEjecutable(string ruta, long tamanoEsperado)
        {
            var fi = new FileInfo(ruta);
            if (!fi.Exists) throw new IOException("No se encontró el archivo descargado.");

            if (tamanoEsperado > 0 && fi.Length != tamanoEsperado)
                throw new IOException($"La descarga quedó incompleta: {fi.Length:N0} de {tamanoEsperado:N0} bytes.");

            if (fi.Length < TamanoMinimoRazonable)
                throw new IOException($"El archivo descargado es demasiado pequeño ({fi.Length:N0} bytes); " +
                                      "probablemente sea una página de error y no el programa.");

            using var fs = File.OpenRead(ruta);
            if (fs.ReadByte() != 'M' || fs.ReadByte() != 'Z')
                throw new IOException("El archivo descargado no es un ejecutable de Windows.");
        }

        /// <summary>
        /// Sustituye el ejecutable en uso por el descargado. Si algo falla después de
        /// apartar el original, lo devuelve a su sitio para no dejar la instalación rota.
        /// Devuelve la ruta del respaldo.
        /// </summary>
        public static string Instalar(string rutaDescargada, string exeActual)
        {
            string respaldo = exeActual + ".old";
            if (File.Exists(respaldo))
            {
                try { File.Delete(respaldo); }
                catch { respaldo = $"{exeActual}.old{DateTime.Now:yyyyMMddHHmmss}"; }
            }

            // Renombrar un ejecutable en uso sí está permitido; borrarlo no.
            File.Move(exeActual, respaldo);

            try
            {
                File.Move(rutaDescargada, exeActual);
            }
            catch
            {
                try { if (!File.Exists(exeActual)) File.Move(respaldo, exeActual); } catch { /* se informa el error original */ }
                throw;
            }

            return respaldo;
        }

        /// <summary>Borra los respaldos que dejaron actualizaciones anteriores.</summary>
        public static void LimpiarRespaldos()
        {
            try
            {
                string? exe = Environment.ProcessPath;
                if (exe == null) return;
                string? carpeta = Path.GetDirectoryName(exe);
                if (carpeta == null) return;

                foreach (var f in Directory.EnumerateFiles(carpeta, Path.GetFileName(exe) + ".old*"))
                {
                    try { File.Delete(f); } catch { /* sigue en uso: se intentará la próxima vez */ }
                }
            }
            catch { /* nunca debe impedir que el programa arranque */ }
        }

        /// <summary>Lanza el ejecutable ya actualizado.</summary>
        public static void Reiniciar(string exe)
            => Process.Start(new ProcessStartInfo { FileName = exe, UseShellExecute = true });

        public static string CarpetaEscribible(string exe)
            => Path.GetDirectoryName(exe) ?? Environment.CurrentDirectory;

        /// <summary>Comprueba por adelantado si se puede escribir junto al ejecutable.</summary>
        public static bool PuedeEscribirJuntoAlExe(string exe, out string motivo)
        {
            motivo = string.Empty;
            try
            {
                string prueba = Path.Combine(CarpetaEscribible(exe), $".escritura-{Guid.NewGuid():N}.tmp");
                using (var fs = File.Create(prueba)) { }
                File.Delete(prueba);
                return true;
            }
            catch (Exception ex)
            {
                motivo = ex.Message;
                return false;
            }
        }
    }
}
