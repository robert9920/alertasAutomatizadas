# App Alertas

Aplicación de escritorio para Windows que arma y envía el **Reporte de proyecto** al
cliente: texto fijo, indicadores de avance, Curva S y tabla de entregables pendientes
de conformidad. El correo se ve y se edita en pantalla antes de enviarlo, proyecto por
proyecto, con un clic en **Enviar**.

---

## 1. Puesta en marcha

### Versión portable (.exe)

1. Copia la carpeta `App Alertas` completa donde quieras.
2. Abre `App Alertas.exe`.
3. Rellena `BD_Reportes.xlsx` (está al lado del ejecutable) y pulsa **Actualizar**.

No requiere instalar Python ni permisos de administrador.

### Desde el código

Sirve una ventana normal de **PowerShell**: `python` y `conda` ya están en el PATH, y la
política de ejecución (`RemoteSigned`) permite lanzar los `.ps1` de este proyecto sin
cambiar nada.

**Preparar el entorno, una sola vez:**

```powershell
conda create -n app-alertas python=3.12 -y
conda activate app-alertas
cd "C:\Trabajo\Desarrollos\Solicitud Automatizaciones\PMO\Juan\D09-TA06-26-0024-V1\App Alertas"
pip install -r requirements.txt
```

> **Por qué Python 3.12 y no 3.13.** Las librerías que se instalaron al principio quedaron en
> el *site-packages del usuario* (`…\AppData\Roaming\Python\Python313\site-packages`), y esa
> ruta va **antes** que la de cualquier entorno en `sys.path`. Un entorno con 3.13 heredaría
> esas copias sin avisar; con 3.12 el aislamiento es real y no hay que desinstalar nada.

**Cada vez que se toque el código:**

```powershell
conda activate app-alertas
cd "…\App Alertas"

python App_Alertas.py       # probar el cambio en la aplicación
python pruebas.py           # 44 comprobaciones sobre el Excel modelo
.\preparar_entrega.ps1      # recompila el .exe y regenera el ZIP de entrega
```

`preparar_entrega.ps1` se encarga solo: recompila únicamente si algún `.py` es más nuevo que
el `.exe`, ejecuta `pruebas.py` y `revisar_entrega.py`, y **aborta sin generar el ZIP** si
algo falla. Antes de una entrega nueva conviene subir `__version__` en `app/__init__.py`,
porque de ahí sale el nombre del paquete.

Para arrancar sin conda (usando el Python del sistema) basta con
`pip install -r requirements.txt` y `python App_Alertas.py`; también vale
`python -m app.main`.

### Generar solo el .exe

```powershell
.\build.ps1                 # un solo ejecutable portable
.\build.ps1 -Carpeta        # carpeta con dependencias sueltas (arranca más rápido)
```

El resultado queda en `dist\App Alertas\`. Para el paquete que se envía a otra persona, usa
`preparar_entrega.ps1` (ver el apartado 7).

### Subir el proyecto a GitHub

El `.gitignore` ya está preparado. **Inicializa el repositorio dentro de `App Alertas`**, no
en la carpeta padre: ahí está `Referencia\LE.xlsx`, con datos reales de un cliente.

```powershell
cd "…\App Alertas"
git init
git add .
git status                  # revisa la lista antes del primer commit
git commit -m "App Alertas: reporte de avance y entregables"
git remote add origin https://github.com/<usuario>/<repositorio>.git
git push -u origin main
```

Quedan fuera del repositorio `BD_Reportes.xlsx` (su hoja `SMTP` guarda la contraseña en texto
plano y las otras hojas, correos de clientes), cualquier otro Excel, `logs\`, `dist\`,
`entrega\` y `build\`. Quien clone el repositorio regenera su base de datos con
`python crear_bd.py`.

---

## 2. El Excel de base de datos (`BD_Reportes.xlsx`)

Es la única configuración de la aplicación. Tras editarlo, pulsa **Actualizar** (o F5):
no hace falta cerrar la app.

### Hoja `Proyectos`

| Campo | Obligatorio | Descripción |
|---|---|---|
| Código Proyecto | Sí | Clave que enlaza con el resto de hojas. |
| Nombre Proyecto | No | Si se deja vacío se toma de la hoja `CARATULA` del Excel del proyecto. |
| Cliente | No | Igual que el anterior. |
| Ruta L | Sí | Carpeta donde está el Excel de control de entregables. |
| Nombre de Excel | Sí | Nombre del archivo **sin** la extensión `.xlsx`. |
| Hoja LET | No | Solo si la hoja no se llama `LET` ni empieza por `LET`. |
| Hoja EV | No | Igual que el anterior. |
| Estado Filtro | No | Estatus a filtrar. Por defecto `En revisión del cliente`. |
| Activo | No | `Sí` / `No`. Los inactivos aparecen en gris. |

### Hoja `Destinatarios`

| Campo | Descripción |
|---|---|
| Código Proyecto | Enlaza el correo con su proyecto. |
| Nombre | Opcional, solo informativo. |
| Correo | Dirección de destino. |
| Tipo Destinatario | `Para`, `CC` o `CCO`. **Si se deja vacío se usa `Para`.** |
| Activo | `No` deja la fila desmarcada al abrir el proyecto; siempre se puede marcar en la app. |

### Hojas `MapeoLET` y `MapeoEV` (opcionales)

Sirven para fijar dónde está cada campo cuando un Excel se sale del formato habitual.

- **`MapeoLET`**: se escribe la **letra de la columna** (`M`, `n`, `BS`… da igual
  mayúsculas o minúsculas), más `Fila Encabezado` y `Fila Inicio Datos` si hacen falta.
- **`MapeoEV`**: se escribe el **número de fila** de cada serie, la `Columna Inicio
  Semanas` y la `Celda SPI` (por ejemplo `H44`).

**Cada campo que se deje vacío se resuelve solo**, en este orden:

1. Lo indicado en el mapeo de la BD.
2. Búsqueda del nombre del campo dentro de la hoja, sin distinguir tildes,
   mayúsculas ni espacios de más (y tolerando variantes parecidas).
3. La posición del archivo modelo `Referencia/LE.xlsx`.

En la pestaña **Datos y filtros** el botón *Ver cómo se leyó el Excel* muestra, campo por
campo, qué referencia se usó y de cuál de los tres orígenes salió.

### Hoja `SMTP`

`Servidor`, `Puerto`, `Seguridad` (`STARTTLS` / `SSL` / `Ninguna`), `Usuario`,
`Contraseña`, `Remitente`, `Nombre Remitente`, `Responder a`, `CCO fijo` y `Timeout (s)`.

> ⚠️ **La contraseña se guarda en texto plano.** Protege el acceso a este archivo.
> Si el correo corporativo es Microsoft 365 y el envío falla con error de
> autenticación, es porque el tenant tiene SMTP AUTH deshabilitado: pide a TI que lo
> habilite para la cuenta o que genere una contraseña de aplicación.

El botón **Probar SMTP** de la barra superior comprueba la conexión sin enviar nada.

### Hoja `Plantilla`

Los textos fijos del correo. Admiten HTML simple (`<b>`, `<i>`, `<u>`, `<br>`) y estos
comodines:

`{fecha_hoy}` `{codigo_proyecto}` `{nombre_proyecto}` `{cliente}` `{semana}`
`{avance_planificado}` `{avance_real}` `{desviacion}` `{spi}` `{n_entregables}`

Asunto por defecto:
`Gestión: Reporte de proyecto {fecha_hoy} | {codigo_proyecto} - {nombre_proyecto}`

### Hoja `Config`

Está dividida en secciones. Cada parámetro lleva su ayuda en la tercera columna.

| Sección | Qué controla |
|---|---|
| **Gráfico · tamaño de la imagen y ejes** | Ancho, alto, DPI, rangos de los dos ejes Y, semanas a mostrar y si se dibujan las etiquetas de datos. |
| **Gráfico · tamaños de letra** | Etiquetas de las líneas y de las barras, porcentajes de cada eje Y, semanas y meses del eje X, leyenda y título. |
| **Gráfico · colores de las series** | Las 3 líneas acumuladas y las 3 barras semanales. |
| **Gráfico · colores de las etiquetas** | Las 6 por separado. |
| **Gráfico · eje X y título** | Fondo y texto de los recuadros de semanas y meses, y color del título. |
| **Correo · tabla de entregables** | Color del encabezado, color del encabezado de estatus y decimales de los indicadores. |
| **Correo · firma** | Ruta de la firma y ancho con el que se inserta. |
| **Aplicación** | Tema claro u oscuro. |

Dos convenios importantes:

- **Los tamaños de letra vacíos significan «automático»**: la aplicación los ajusta al
  número de semanas, que es lo que mantiene legible un proyecto de 36 semanas. En cuanto
  escribes un número, manda ese número.
- **Los colores de etiqueta vacíos se heredan**: las de las líneas toman el color de su
  propia línea y las de las barras salen en negro.

Los porcentajes de los ejes Y y el texto de la leyenda no siguen a `Color texto eje X`:
se mantienen en gris para que oscurecer las bandas no los deje invisibles.

### La firma del correo

Guarda tu firma como **`firma.png`** (también vale `.jpg`) **en la misma carpeta que
`App Alertas.exe`**: la aplicación la encuentra sola, sin configurar nada. Se inserta
después del «Muchas gracias».

Si la tienes en otro sitio —por ejemplo una carpeta de red compartida— escribe la ruta
completa del archivo en `Config` → `Ruta firma`. El ancho se ajusta con `Ancho firma px`
(330 por defecto); el alto se calcula solo, respetando la proporción.

La firma se vuelve a leer cada vez que pulsas **Actualizar**, así que puedes cambiarla sin
cerrar la aplicación.

---

## 3. Cómo se calculan los datos

- **Semana de corte**: la última semana que tiene valor a la vez en `% Previsto` y en
  `% Real`. Es la que aparece como «Semana N» en el texto del correo.
- **Avance planificado** = `% Previsto Acum` de esa semana.
- **Avance real** = `% Real Acum` de esa semana.
- **Desviación** = avance real − avance planificado, calculada **sobre los valores ya
  redondeados** para que los tres números del correo sean coherentes entre sí.
  Con `Decimales avance = 0` se muestran como enteros.
- **SPI**: celda indicada en `MapeoEV`; si no, se busca la etiqueta `SPI` en la hoja y
  se toma el primer número a su derecha.
- **Tabla de entregables**: filas de la hoja `LET` cuyo `ESTATUS DEL ENTREGABLE LC`
  coincide con el filtro. Si el valor configurado no existe en esa hoja, la app avisa y
  deja elegir cualquiera de los estatus presentes mediante casillas.
- **Días de espera** = hoy − `FECHA ÚLTIMO ENVÍO A CLIENTE`, en días de calendario. Se
  calcula al generar el correo, no al leer el Excel, así que sigue siendo correcto aunque
  la aplicación lleve días abierta. Sin fecha de envío, la celda queda vacía.

### Etiquetas de la Curva S

Cada semana, la etiqueta de la línea que va por encima se dibuja arriba y la de la que va
por debajo, abajo. La comparación es contra `% Previsto Acum`, usando como referencia el
`% Real Acum` y, cuando este ya no tiene datos, el `% Tendencia Acum`.

- Si las dos cifras coinciden (las líneas se superponen) se dibuja **una sola**.
- La **primera etiqueta de `% Tendencia Acum` no se dibuja**, porque repite el último
  valor de `% Real Acum`.
- Una etiqueta que caiga demasiado abajo se sube, para que no se salga del gráfico.

---

## 4. Uso diario

1. Elige el proyecto en la lista de la izquierda.
2. Revisa la pestaña **Datos y filtros**: semana, indicadores, estatus incluidos y qué
   entregables entrarán en la tabla (se puede desmarcar cualquiera).
3. Vuelve a **Correo** y ajusta lo que haga falta: el cuerpo es editable como en un
   procesador de textos, y la pestaña **HTML** permite retoques finos.
4. Revisa a la derecha el asunto, los destinatarios (casilla de envío y `Para`/`CC`/`CCO`)
   y añade adjuntos si corresponde.
5. **ENVIAR CORREO** → aparece un resumen para confirmar antes de que salga nada.

Otros botones útiles:

- **Regenerar correo**: rehace el correo desde la plantilla y descarta las ediciones
  manuales (pide confirmación).
- **Guardar .eml**: guarda el correo como borrador para abrirlo en Outlook y revisarlo.
- **Historial**: correos enviados desde la aplicación.
- **Tema**: alterna claro y oscuro. El cuerpo del correo siempre se ve sobre fondo
  blanco, como lo verá el cliente.

---

## 5. Estructura del proyecto

```
App Alertas/
├─ app/
│  ├─ main.py             arranque de la aplicación
│  ├─ config.py           rutas, config.json y registro
│  ├─ constantes.py       nombres de campos, posiciones de respaldo y textos por defecto
│  ├─ bd.py               lectura y validación de BD_Reportes.xlsx
│  ├─ excel_compat.py     apertura robusta de libros (evita el fallo de tablas dinámicas)
│  ├─ excel_utils.py      normalización de texto y búsqueda tolerante
│  ├─ let_reader.py       entregables y estatus de la hoja LET
│  ├─ ev_reader.py        series de avance, semana de corte, KPIs y SPI de la hoja EV
│  ├─ lectura.py          abre el Excel del proyecto una sola vez y junta todo
│  ├─ chart.py            Curva S en PNG
│  ├─ email_builder.py    asunto, cuerpo HTML y texto plano
│  ├─ sender.py           mensaje MIME y envío SMTP
│  ├─ historial.py        registro de envíos
│  └─ ui/                 interfaz (PySide6)
├─ recursos/              icono
├─ logs/                  app.log e historial.csv
├─ App_Alertas.py         lanzador (punto de entrada del ejecutable)
├─ preparar_entrega.ps1   genera el ZIP para otra persona
├─ revisar_entrega.py     comprueba que el paquete no lleve datos tuyos
├─ crear_bd.py            genera BD_Reportes.xlsx
├─ .gitignore             deja fuera credenciales, Excel y compilados
├─ pruebas.py             pruebas de humo
├─ requirements.txt       dependencias
└─ build.ps1              empaquetado con PyInstaller
```

`App_Alertas.py` se mantiene deliberadamente mínimo: prepara la salida estándar (sin
consola, `--windowed` deja `stdout` y `stderr` en `None` y cualquier aviso de una
librería abortaría el arranque) y deja por escrito en `logs\arranque_error.log`
cualquier fallo, incluidos los errores de importación, que de otro modo solo producirían
un cuadro de diálogo sin información.

La app escribe `logs\` y `config.json` junto al ejecutable, para seguir siendo portable.
Si esa carpeta es de solo lectura (colocada en `C:\Program Files`, o abierta desde una
carpeta de red ajena) pasa automáticamente a `%LOCALAPPDATA%\App Alertas\`, de modo que
nunca deje de abrirse por un problema de permisos.

---

## 6. Problemas frecuentes

| Síntoma | Causa y solución |
|---|---|
| «No se encuentra el archivo …» al abrir un proyecto | `Ruta L` o `Nombre de Excel` mal escritos en la hoja `Proyectos`. |
| Todos los indicadores salen en «-» | El Excel del proyecto nunca se guardó con Excel, así que no tiene valores calculados en caché. Ábrelo, guárdalo y pulsa **Actualizar**. |
| No aparece ningún entregable | El estatus configurado no existe en esa hoja. La app lo avisa; marca el estatus correcto en **Datos y filtros**. |
| Un campo se lee de la columna equivocada | Fíjalo en `MapeoLET` / `MapeoEV` y pulsa **Actualizar**. *Ver cómo se leyó el Excel* indica qué se está usando. |
| Error de autenticación al enviar | SMTP AUTH deshabilitado o MFA en la cuenta. Ver la nota de la hoja `SMTP`. |
| La app no arranca | Revisa `logs\arranque_error.log` (fallos de arranque) y `logs\app.log`. |
| «Windows protegió su PC» al abrir el .exe | El ejecutable no está firmado digitalmente. *Más información → Ejecutar de todas formas*. Solo ocurre la primera vez. |

---

## 7. Cómo compartir la aplicación

### Generar el paquete

```powershell
.\preparar_entrega.ps1              # usa el .exe existente si está al día
.\preparar_entrega.ps1 -Recompilar  # fuerza una compilación limpia antes
.\preparar_entrega.ps1 -Carpeta     # versión en carpeta, si el antivirus bloquea el .exe único
```

Deja en `entrega\App Alertas <versión>.zip` un paquete con el programa, una
`BD_Reportes.xlsx` **vacía** y la documentación. Antes de comprimir ejecuta
`revisar_entrega.py`, que aborta si la base de datos todavía tuviera proyectos,
destinatarios, credenciales o rutas de tu equipo.

### Qué enviar

Solo ese `.zip`. Contiene:

| Archivo | Para qué |
|---|---|
| `App Alertas.exe` | La aplicación |
| `BD_Reportes.xlsx` | Plantilla vacía con las 8 hojas, listas desplegables y ayudas |
| `LEEME.md` | Este manual |
| `PRIMEROS PASOS.txt` | Guía de una página para arrancar |

**Lo que NO se envía:**

- La carpeta `Referencia\`. Las posiciones del archivo modelo están codificadas en
  `app/constantes.py`; ese Excel no se lee en tiempo de ejecución.
- **Tu** `BD_Reportes.xlsx`, porque su hoja `SMTP` guarda tu contraseña en texto plano.
- El código fuente, salvo que la otra persona vaya a modificarlo.

### Qué necesita la otra persona

**Nada instalado.** El ejecutable lleva dentro el intérprete de Python, las DLL de Qt y el
runtime de Visual C++. Solo requiere **Windows 10 u 11 de 64 bits**. Excel le hará falta
únicamente para editar la base de datos, no para que la aplicación funcione.

Lo mismo vale para ti: para *ejecutar* el `.exe` no necesitas Python. Solo hace falta para
recompilarlo (`build.ps1`) o ejecutar las pruebas.

### Qué debe configurar

1. Descomprimir en una carpeta **con permiso de escritura** (Documentos, Escritorio, una
   carpeta de red propia). Si la coloca en `C:\Program Files`, la app seguirá abriendo pero
   escribirá sus registros en `%LOCALAPPDATA%\App Alertas\`.
2. Hoja `Proyectos`: sus proyectos. **`Ruta L` en formato UNC** (`\\servidor\carpeta`) y no
   con la letra mapeada (`L:\...`): la misma letra puede apuntar a otra carpeta —o no
   existir— en su equipo. Y debe tener permiso de lectura sobre esas carpetas.
3. Hoja `Destinatarios`: los correos del cliente.
4. Hoja `SMTP`: **su propia** cuenta. Cada persona envía desde la suya.

### Avisos de Windows

La primera vez, SmartScreen mostrará «Windows protegió su PC» porque el ejecutable no está
firmado digitalmente: *Más información → Ejecutar de todas formas*. Evitarlo requeriría
comprar un certificado de firma de código.

Algunos antivirus corporativos marcan los ejecutables de un solo archivo de PyInstaller.
Si ocurre, `.\preparar_entrega.ps1 -Carpeta` genera una versión en carpeta que suele pasar
sin problemas y además arranca más rápido.

### Actualizaciones posteriores

Para entregar una versión nueva basta con que sustituya **`App Alertas.exe`**. Su
`BD_Reportes.xlsx` no se toca, así que no pierde su configuración. Sube el número en
`__version__` de `app/__init__.py` para que el ZIP quede identificado.

---

Ejecuta `python pruebas.py` para verificar de una vez la lectura del Excel modelo, la
tolerancia a cambios de estructura, el mapeo explícito y el armado del correo.
