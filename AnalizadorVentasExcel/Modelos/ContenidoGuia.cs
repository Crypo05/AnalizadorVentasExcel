using System.Collections.Generic;

namespace AnalizadorVentasExcel.Modelos
{
    public sealed class EntradaGuia
    {
        public string Nombre { get; init; } = string.Empty;
        public string Descripcion { get; init; } = string.Empty;
    }

    public sealed class SeccionGuia
    {
        public string Titulo { get; init; } = string.Empty;
        public string Resumen { get; init; } = string.Empty;
        public List<EntradaGuia> Entradas { get; init; } = new();
    }

    public sealed class RecetaGuia
    {
        public string Objetivo { get; init; } = string.Empty;
        public string Pasos { get; init; } = string.Empty;
    }

    /// <summary>Texto de la guía de usuario que abre el botón "?".</summary>
    public static class ContenidoGuia
    {
        public static List<SeccionGuia> Secciones() => new()
        {
            new SeccionGuia
            {
                Titulo = "1. Carga de datos",
                Resumen = "Todo empieza aquí: el programa lee una carpeta completa de archivos Excel, " +
                          "y trata cada archivo como una sucursal distinta.",
                Entradas =
                {
                    new EntradaGuia
                    {
                        Nombre = "Tipo de Negocio → Detección Automática",
                        Descripcion = "Opción recomendada. Decide sola cómo leer cada archivo: si el Excel trae una " +
                                      "columna \"Artículo\" con códigos, lo trata como Minimarket; si no, como Souvenir."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Tipo de Negocio → Minimarket (Detallado)",
                        Descripcion = "Fuerza el modo detallado: cada fila del Excel es un producto individual. El nombre " +
                                      "del producto sale de la columna \"Artículo desc.\" (o \"Descripción\" / \"Nombre\"). " +
                                      "Las filas sin código de artículo se descartan."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Tipo de Negocio → Souvenir (Agrupado)",
                        Descripcion = "Fuerza el modo agrupado: no hay productos individuales, cada fila se identifica por " +
                                      "su Familia y el nombre de la familia se usa como nombre del producto. Útil cuando el " +
                                      "reporte viene resumido por categoría."
                    },
                    new EntradaGuia
                    {
                        Nombre = "📂 Seleccionar Carpeta",
                        Descripcion = "Abre un diálogo de archivos. Seleccioná CUALQUIER Excel de la carpeta que querés " +
                                      "analizar: se cargan automáticamente TODOS los .xlsx y .xls de esa carpeta, no solo el " +
                                      "que elegiste. El nombre de cada archivo se convierte en el nombre de la sucursal " +
                                      "(por ejemplo, \"Mirador.xlsx\" → sucursal \"Mirador\"). Los archivos temporales de " +
                                      "Excel (los que empiezan con ~$) se ignoran."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Barra de progreso y mensaje de estado",
                        Descripcion = "Mientras lee, la ventana sigue respondiendo y el texto bajo el botón indica cuántos " +
                                      "archivos lleva. Al terminar muestra en verde cuántos archivos y cuántas filas se " +
                                      "cargaron y en cuántos segundos. Si algún archivo falla, aparece un aviso con el " +
                                      "detalle, pero los demás se cargan igual."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Formato que espera el Excel",
                        Descripcion = "El encabezado se busca en las primeras 20 filas con datos. Se reconocen las columnas: " +
                                      "\"Año mes\", \"Artículo\", \"Artículo desc.\", \"Proveedor\", \"Familia\", \"Total\" y " +
                                      "\"% Utilidad\". Si el periodo viene combinado (escrito solo en la primera fila de cada " +
                                      "mes, como en las tablas dinámicas), se arrastra hacia abajo automáticamente. " +
                                      "Las filas con Total = 0 no se cargan."
                    }
                }
            },

            new SeccionGuia
            {
                Titulo = "2. Filtros",
                Resumen = "Recortan qué datos entran al análisis. Todo lo que veas en la tabla y en el gráfico " +
                          "respeta estos cuatro filtros a la vez.",
                Entradas =
                {
                    new EntradaGuia
                    {
                        Nombre = "Sucursales",
                        Descripcion = "Una casilla por archivo cargado. Solo se analizan las sucursales marcadas. " +
                                      "Si las desmarcás todas, la tabla y el gráfico quedan vacíos."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Periodos",
                        Descripcion = "Un mes por casilla, en formato Año-Mes, ordenados del más reciente al más antiguo."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Proveedores",
                        Descripcion = "Todos los proveedores encontrados en los archivos. Tiene buscador."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Familias",
                        Descripcion = "Las categorías de producto. Tiene buscador. Esta lista NO es fija: se recalcula " +
                                      "según los proveedores y las sucursales que tengas marcados, y solo muestra las " +
                                      "familias que realmente existen en esa combinación. Al recalcularse, se marcan todas."
                    },
                    new EntradaGuia
                    {
                        Nombre = "🔎 Buscador (Proveedores y Familias)",
                        Descripcion = "Escribí parte del nombre para filtrar la lista mostrada. Lo importante: el buscador " +
                                      "solo cambia lo que SE VE, nunca lo que está marcado. Podés buscar algo, marcarlo, " +
                                      "borrar la búsqueda, buscar otra cosa y marcarla: todo lo anterior se conserva. " +
                                      "No distingue mayúsculas ni minúsculas."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Botones \"Todas\" / \"Todos\"",
                        Descripcion = "Marca todos los elementos VISIBLES en ese momento. Si no hay búsqueda activa, marca " +
                                      "toda la lista. Si hay una búsqueda activa, marca solo los resultados de esa búsqueda, " +
                                      "sin tocar el resto."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Botón \"Ninguna\" (Proveedores y Familias)",
                        Descripcion = "Desmarca los elementos visibles. Combinado con el buscador es la forma rápida de " +
                                      "aislar un grupo: ver más abajo la receta \"Aislar unas pocas familias\"."
                    }
                }
            },

            new SeccionGuia
            {
                Titulo = "3. Gráfico",
                Resumen = "Define cómo se agrupan y se dibujan los datos ya filtrados. Los tres controles se combinan " +
                          "entre sí y afectan también a la tabla.",
                Entradas =
                {
                    new EntradaGuia
                    {
                        Nombre = "Eje Principal (X)",
                        Descripcion = "Es el criterio con el que se agrupan los resultados; cada valor distinto es una " +
                                      "posición del eje horizontal y una fila de la tabla. Opciones: Año Mes (evolución en " +
                                      "el tiempo), Familia, Proveedor, Sucursal y Articulo."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Límite de 20 categorías",
                        Descripcion = "Cuando el eje X no es Año Mes, el gráfico dibuja solo las 20 categorías de mayor " +
                                      "valor, para que sea legible. La TABLA sí muestra todas: con eje Articulo podés tener " +
                                      "20 puntos en el gráfico y miles de filas en la tabla."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Desglose (Series)",
                        Descripcion = "Parte cada grupo del eje X en varias líneas. Podés marcar más de una dimensión a la " +
                                      "vez y se combinan: marcando Familia y Sucursal aparecen series como " +
                                      "\"CERVEZAS - Bomba\". Si no marcás nada, hay una sola línea llamada \"Total\"."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Límite de 10 series",
                        Descripcion = "El gráfico dibuja como máximo las 10 series de mayor valor, cada una con su color en " +
                                      "la leyenda de la derecha. La tabla, otra vez, muestra todas las combinaciones."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Operación → Suma Total (Colones)",
                        Descripcion = "Suma el dinero vendido de cada grupo. Es el modo normal."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Operación → Conteo (Transacciones)",
                        Descripcion = "Cuenta cuántas filas del Excel caen en cada grupo, en lugar de sumar dinero. Sirve " +
                                      "para separar \"vendo mucho dinero\" de \"vendo muchas veces\". Con Conteo, la columna " +
                                      "% Part. no aplica y aparece como \"-\"."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Leer el gráfico",
                        Descripcion = "Pasá el mouse por una posición del eje X y el tooltip muestra el valor de todas las " +
                                      "series en ese punto a la vez, para comparar directo. Las series con valor cero en " +
                                      "ese punto se omiten del tooltip."
                    }
                }
            },

            new SeccionGuia
            {
                Titulo = "4. Tabla de resultados",
                Resumen = "Es el mismo cálculo del gráfico pero completo y ordenado de mayor a menor.",
                Entradas =
                {
                    new EntradaGuia { Nombre = "Concepto", Descripcion = "El valor del Eje Principal (X) del grupo." },
                    new EntradaGuia
                    {
                        Nombre = "Detalle",
                        Descripcion = "La combinación de desglose de esa fila, o \"Total General\" si no hay desglose."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Suma Total / Conteo",
                        Descripcion = "El valor según la operación elegida. El encabezado de la columna cambia solo."
                    },
                    new EntradaGuia
                    {
                        Nombre = "% Utilidad",
                        Descripcion = "Promedio del margen de utilidad de las filas del grupo. Ojo: es un promedio simple, " +
                                      "no ponderado (ver la sección de notas al final)."
                    },
                    new EntradaGuia
                    {
                        Nombre = "% Part.",
                        Descripcion = "Cuánto pesa esa fila sobre el total de todo lo filtrado. Sirve para ver " +
                                      "concentración: si un proveedor tiene 31 %, casi un tercio de la venta depende de él."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Orden",
                        Descripcion = "De mayor a menor valor, salvo con eje Año Mes, que va cronológico."
                    }
                }
            },

            new SeccionGuia
            {
                Titulo = "5. Explorador de Productos (🔍 Auditar Anomalías)",
                Resumen = "Un modo aparte, pensado para revisar producto por producto y encontrar inconsistencias " +
                          "entre sucursales.",
                Entradas =
                {
                    new EntradaGuia
                    {
                        Nombre = "Qué hace el botón",
                        Descripcion = "Consolida TODOS los productos por nombre, sumando las sucursales, respetando los " +
                                      "filtros de Sucursales y Periodos que tengas puestos (los de Proveedor y Familia no " +
                                      "se aplican aquí)."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Columna Disponibilidad",
                        Descripcion = "En cuántas tiendas se vendió ese producto, y en la columna Detalle, cuáles. Un " +
                                      "producto que debería estar en las tres tiendas y aparece con \"1 Tiendas\" es " +
                                      "candidato a revisión."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Seleccionar un producto",
                        Descripcion = "Al hacer clic en una fila, el gráfico de abajo cambia a una comparativa: una línea " +
                                      "por sucursal con el margen de utilidad mes a mes de ese producto. El nombre del " +
                                      "producto aparece en el subtítulo."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Huecos y puntos sueltos",
                        Descripcion = "Un corte en la línea significa que esa sucursal no vendió el producto ese mes. Un " +
                                      "punto aislado significa que solo hay un mes con datos, y por eso no hay línea que " +
                                      "dibujar."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Para qué sirve en la práctica",
                        Descripcion = "Detectar el mismo producto cargado con descripciones distintas en cada tienda " +
                                      "(aparecerá como dos filas parecidas con 1 tienda cada una), y detectar márgenes muy " +
                                      "dispares para el mismo producto entre sucursales."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Cómo salir del explorador",
                        Descripcion = "Cambiá cualquier filtro, el eje X, el desglose o la operación y la vista vuelve al " +
                                      "análisis normal."
                    }
                }
            },

            new SeccionGuia
            {
                Titulo = "6. Comparativa de Precios (⚖️ botón morado)",
                Resumen = "Un sistema aparte, con su propia ventana y sus propios archivos. No usa los datos de " +
                          "ventas: lee las listas de precios de cada sucursal y compara, producto por producto, " +
                          "cuánto cuesta lo mismo en cada tienda. Se puede tener abierta a la vez que el análisis.",
                Entradas =
                {
                    new EntradaGuia
                    {
                        Nombre = "Qué archivos necesita",
                        Descripcion = "Los reportes de \"Comparativa de precios\", no los de ventas. Se reconocen por su " +
                                      "encabezado: \"Cód. Artículo\", \"Descripción\", \"Precio costo\", \"Imp. ventas\", " +
                                      "\"Porc. utilidad - artículo\" y \"Precio IVI - artículo\". Poné una lista por sucursal " +
                                      "en una misma carpeta y elegí cualquiera de ellas: se cargan todas. Si algún archivo " +
                                      "no tiene ese encabezado, se ignora y el programa te dice cuál."
                    },
                    new EntradaGuia
                    {
                        Nombre = "El nombre del archivo es el nombre de la sucursal",
                        Descripcion = "Se le quitan la fecha y la palabra \"precios\": \"precios la bomba 07-09-26.xlsx\" " +
                                      "queda como sucursal \"la bomba\". Con eso el encabezado de la columna es corto y legible."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Sucursales a comparar",
                        Descripcion = "Cada sucursal marcada es una columna de la tabla. Desmarcá las que no te interesen " +
                                      "para comparar solo dos, o dejalas todas para ver el panorama completo. Los totales " +
                                      "de arriba y el color de cada celda se recalculan con las sucursales marcadas."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Comparar por",
                        Descripcion = "\"Precio de venta (IVI)\" es lo que paga el cliente: sirve para detectar que el mismo " +
                                      "producto se vende más caro en una tienda. \"Precio de costo\" compara lo que costó " +
                                      "comprarlo, que es donde aparecen los problemas de negociación con el proveedor. " +
                                      "\"% de utilidad\" compara el margen. En precio y costo el mejor valor es el más bajo; " +
                                      "en utilidad, el más alto (y las columnas cambian de nombre para recordarlo)."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Productos",
                        Descripcion = "\"En 2 o más sucursales\" es lo normal: solo lo que se puede comparar. \"Sólo los que " +
                                      "están en todas\" deja los productos que maneja toda la cadena. \"Todos\" agrega los " +
                                      "exclusivos de una sola tienda, que aparecen con — en las demás columnas y con el aviso " +
                                      "\"Sólo en ...\": útil para ver qué le falta a una sucursal."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Sólo los que tienen diferencia · Diferencia mínima · Ordenar por · Buscar",
                        Descripcion = "Filtros de la vista. La casilla esconde los productos que valen igual en todas partes; " +
                                      "la \"Diferencia mínima\" deja solo los que se separan más de ese porcentaje; el orden " +
                                      "por defecto pone arriba las mayores diferencias porcentuales. El buscador mira el " +
                                      "código de barras y el nombre en cualquiera de las sucursales."
                    },
                    new EntradaGuia
                    {
                        Nombre = "Cómo leer la tabla",
                        Descripcion = "Una columna por sucursal, en verde el mejor valor y en rojo el peor (si el producto " +
                                      "vale lo mismo en todas, no se pinta nada). \"Dif.\" es la diferencia en colones entre " +
                                      "la sucursal más cara y la más barata, y \"Dif. %\" esa misma diferencia respecto del " +
                                      "valor más bajo. \"Suc.\" dice en cuántas de las sucursales comparadas está el producto."
                    },
                    new EntradaGuia
                    {
                        Nombre = "El aviso \"⚠ Nombres distintos\"",
                        Descripcion = "El cruce se hace por código de artículo, que es la única llave fiable: el mismo " +
                                      "producto suele estar escrito distinto en cada tienda (\"PASTILLAS ALEVE GELS UND\" " +
                                      "contra \"ALEVE GEL UND\"). Cuando los nombres no coinciden se marca la fila, y el " +
                                      "nombre de cada sucursal aparece al pasar el mouse por la columna Descripción. " +
                                      "Casi siempre es solo un tema de digitación, pero a veces revela que el mismo código " +
                                      "está usado para dos productos diferentes: por eso conviene mirarlo antes de sacar " +
                                      "conclusiones de una diferencia grande."
                    },
                    new EntradaGuia
                    {
                        Nombre = "El gráfico de abajo",
                        Descripcion = "Seleccioná una fila de la tabla y el gráfico muestra ese producto sucursal por " +
                                      "sucursal, con el valor sobre cada barra. Es la forma rápida de enseñarle a alguien " +
                                      "una diferencia concreta."
                    }
                }
            }
        };

        public static List<RecetaGuia> Recetas() => new()
        {
            new RecetaGuia
            {
                Objetivo = "Ver cómo evolucionan las ventas mes a mes, comparando sucursales",
                Pasos = "Eje X: Año Mes  •  Desglose: Sucursal  •  Operación: Suma Total"
            },
            new RecetaGuia
            {
                Objetivo = "Saber de qué proveedores dependemos más",
                Pasos = "Eje X: Proveedor  •  Desglose: ninguno  •  Operación: Suma Total. " +
                        "Mirá la columna % Part.: es el peso de cada proveedor sobre el total."
            },
            new RecetaGuia
            {
                Objetivo = "Ver qué categorías pesan más en cada tienda",
                Pasos = "Eje X: Familia  •  Desglose: Sucursal  •  Operación: Suma Total"
            },
            new RecetaGuia
            {
                Objetivo = "Encontrar los productos más vendidos",
                Pasos = "Eje X: Articulo  •  Desglose: ninguno. El gráfico muestra el top 20 y la tabla la lista completa."
            },
            new RecetaGuia
            {
                Objetivo = "Comparar el mismo producto entre sucursales, mes a mes",
                Pasos = "Botón 🔍 Auditar Anomalías  →  clic en el producto de la tabla. " +
                        "Filtrá antes los periodos si querés acotar el rango."
            },
            new RecetaGuia
            {
                Objetivo = "Ver la estacionalidad de una categoría (ej. cervezas en verano)",
                Pasos = "En el buscador de Familias escribí \"CERVEZA\"  →  \"Ninguna\"  →  marcá las que te interesan  →  " +
                        "borrá la búsqueda  →  Eje X: Año Mes."
            },
            new RecetaGuia
            {
                Objetivo = "Aislar unas pocas familias o proveedores",
                Pasos = "Buscá el término  →  botón \"Ninguna\" (desmarca solo lo visible)  →  buscá lo que sí querés  →  " +
                        "botón \"Todas\" (marca solo esos resultados)."
            },
            new RecetaGuia
            {
                Objetivo = "Distinguir productos caros de productos de mucha rotación",
                Pasos = "Eje X: Articulo con Operación: Suma Total, y después la misma vista con Operación: Conteo. " +
                        "Lo que sube mucho en Conteo y poco en Suma es rotación barata."
            },
            new RecetaGuia
            {
                Objetivo = "Analizar una sola sucursal a fondo",
                Pasos = "Desmarcá las demás en Sucursales  →  Eje X: Proveedor o Familia  →  Desglose: Año Mes " +
                        "para ver la evolución dentro de esa tienda."
            },
            new RecetaGuia
            {
                Objetivo = "Ver qué familias vende cada proveedor",
                Pasos = "Eje X: Proveedor  •  Desglose: Familia. Recordá que solo se dibujan las 10 series mayores, " +
                        "pero la tabla las trae todas."
            },
            new RecetaGuia
            {
                Objetivo = "¿En qué productos le estamos cobrando de más (o de menos) que la otra sucursal?",
                Pasos = "⚖️ Comparativa de Precios  •  cargá la carpeta de listas de precios  •  Comparar por: Precio de " +
                        "venta (IVI)  •  Productos: En 2 o más sucursales  •  Ordenar por: Mayor diferencia %. Arriba " +
                        "quedan los casos más gruesos; revisá el aviso de nombres distintos antes de decidir, porque a " +
                        "veces son dos presentaciones distintas con el mismo código."
            },
            new RecetaGuia
            {
                Objetivo = "¿Estamos comprando lo mismo a precios distintos según la tienda?",
                Pasos = "⚖️ Comparativa de Precios  •  Comparar por: Precio de costo  •  Diferencia mínima: 10 % o más. " +
                        "Lo que salga es materia de negociación con el proveedor o error de digitación en la caja."
            },
            new RecetaGuia
            {
                Objetivo = "¿Qué productos vende la otra sucursal que nosotros ni tenemos?",
                Pasos = "⚖️ Comparativa de Precios  •  Productos: Todos (incluye exclusivos)  •  Ordenar por: Descripción. " +
                        "Las filas con — y el aviso \"Sólo en ...\" son las que existen en una sola tienda."
            }
        };

        public static List<EntradaGuia> Notas() => new()
        {
            new EntradaGuia
            {
                Nombre = "La comparativa de precios no comparte datos con el análisis",
                Descripcion = "Son dos sistemas distintos y cada uno carga su propia carpeta. Cargar ventas no llena la " +
                              "comparativa, ni al revés. Tampoco hay periodos: la lista de precios es una foto del día en " +
                              "que se exportó, así que la fecha del archivo es toda la referencia temporal que hay."
            },
            new EntradaGuia
            {
                Nombre = "En la comparativa, un código = un producto",
                Descripcion = "Si el mismo código aparece dos veces dentro de la lista de una sucursal, se toma la primera " +
                              "aparición. Y las filas sin costo, sin precio y sin utilidad (las tres en cero) no se cargan: " +
                              "no hay nada que comparar y ensuciarían el ranking de diferencias."
            },
            new EntradaGuia
            {
                Nombre = "El % de Utilidad es un promedio simple",
                Descripcion = "Se promedia el margen de todas las filas del grupo dándoles el mismo peso, sin ponderar por " +
                              "el monto vendido. Un producto con ₡500 de venta y un margen atípico pesa igual que uno con " +
                              "₡5.000.000. Además, en los archivos hay filas con valores como 100 o 2 en la columna de " +
                              "utilidad (que equivaldrían a 10.000 % y 200 %), y esas inflan mucho el promedio: por eso " +
                              "algunas familias muestran porcentajes de tres cifras. Tomá esa columna como un indicador " +
                              "para investigar, no como un dato contable."
            },
            new EntradaGuia
            {
                Nombre = "Filas con Total = 0",
                Descripcion = "No se cargan. Si un producto aparece en el Excel con venta cero, no existirá en el análisis."
            },
            new EntradaGuia
            {
                Nombre = "El nombre del archivo es el nombre de la sucursal",
                Descripcion = "Si renombrás un archivo, cambia el nombre de la sucursal en todo el programa. Conviene " +
                              "mantener nombres cortos y estables."
            },
            new EntradaGuia
            {
                Nombre = "Los filtros se aplican todos a la vez",
                Descripcion = "Una fila entra al análisis solo si su sucursal, su periodo, su proveedor y su familia están " +
                              "marcados. Si algo no aparece, revisá los cuatro filtros: lo más común es que la lista de " +
                              "Familias se haya recalculado al cambiar los proveedores."
            }
        };
    }
}
