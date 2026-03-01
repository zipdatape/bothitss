# Notificador de Bajas de Usuarios Hitss (App C#)

Aplicación de escritorio en C# (Windows Forms) que replica el flujo del proceso UiPath: detecta correos de cese en Outlook, procesa el Excel y la base CSV, y envía la notificación por correo. Todo desde la misma aplicación, sin .bat ni UiPath.

## Requisitos

- **Windows** con .NET 8.0 Runtime (o SDK si compilas tú).
- **Microsoft Outlook** instalado y configurado (misma cuenta que usarás para leer y enviar).
- Carpetas y archivos de configuración (rutas que configuras en la app).

## Cómo ejecutar

### Opción 1: Publicar y copiar a otro equipo

1. En Visual Studio o desde línea de comandos:
   ```bash
   cd NotificadorBajasHitssApp
   dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true
   ```
2. La salida estará en `bin\Release\net8.0-windows\win-x64\publish\`.
3. Copia toda la carpeta `publish` al otro equipo (o el `.exe` si usas single file y no hay dependencias extra).
4. En el otro equipo, instala [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) si no lo tiene.
5. Ejecuta `NotificadorBajasHitss.exe`.

### Opción 2: Ejecutar desde Visual Studio

1. Abre la solución o el proyecto en Visual Studio.
2. Pulsa F5 o “Iniciar”.

## Uso de la aplicación

1. **Al abrir**, se carga la configuración desde `config.json` (en la misma carpeta que el .exe). Si no existe, se usan valores por defecto.
2. **Configuración**:  
   - Revisa y edita los campos. En las **carpetas** puedes usar el botón **...** para elegir la ruta con el selector de Windows.  
   - **Asunto a buscar**: texto fijo que debe aparecer en el asunto (ej. **CESE DE PERSONAL - **). El correo suele llegar como "CESE DE PERSONAL - 26/02/2026"; la app solo filtra por el texto fijo.  
   - **Carpeta Outlook**: ruta dentro de Outlook (ej.: `Bandeja de entrada\C.H_BAJAS`).
3. **Guardar configuración**: guarda los cambios en `config.json` para la próxima ejecución.
4. **Ejecutar proceso**:  
   - Busca en Outlook (carpeta indicada) correos no leídos con adjunto cuyo asunto contenga **Asunto correo a buscar** + fecha (dd/MM/yyyy).  
   - Si encuentra uno: guarda el adjunto, lo mueve a la carpeta de usuario, procesa el Excel y la base CSV, hace backup y envía el correo de notificación con la tabla de bajas.  
   - Si no encuentra: envía un aviso al destinatario indicando que no se encontró el correo.
5. El **log** en la parte inferior muestra el progreso y posibles errores.

## Configuración (campos)

| Campo | Descripción |
|-------|-------------|
| Proceso (nombre) | Nombre del proceso (informativo). |
| Días a restar (fecha) | Días a restar a hoy para la fecha de búsqueda del asunto (ej.: 1 = ayer). |
| Carpeta temporal | Donde se guardan temporalmente los adjuntos. |
| Carpeta usuario (Excel) | Donde se deja el Excel de bajas (nombre: dd.MM.yy.xlsx). |
| Asunto correo a buscar | Texto fijo que debe contener el asunto (ej.: **CESE DE PERSONAL - **). La fecha en el correo es dinámica. |
| Carpeta Outlook | Carpeta de Outlook donde buscar (ej.: `Bandeja de entrada\C.H_BAJAS`). |
| Nombre hoja Excel | Hoja del Excel de bajas a leer. |
| Carpeta BASE | Carpeta del archivo CSV base. |
| Archivo base CSV | Nombre del CSV (ej.: BASE HITSS.csv). |
| Carpeta backup | Carpeta donde se guardan los backups del CSV. |
| Prefijo archivo backup | Prefijo del nombre del backup (se añade fecha). |
| Correo destinatario | Dirección que recibe la notificación y los avisos de error. |
| Asunto correo notificación | Asunto del correo de notificación. |

## Despliegue en otros equipos

- Copia la carpeta publicada (o el ejecutable si usas single file).
- Asegúrate de que esté instalado **.NET 8 Desktop Runtime** (o el runtime que hayas usado al publicar).
- En el primer arranque se creará `config.json` al guardar la configuración; configura las rutas y la carpeta de Outlook según ese equipo.
- **Outlook** debe estar instalado y configurado; la app usa la cuenta por defecto para leer y enviar.

## Actualizaciones automáticas y releases en GitHub

Al **iniciar** la aplicación se consulta si hay una versión más nueva en el repositorio configurado (`zipdatape/bothitss`). La detección usa la **última release publicada** en GitHub, no solo los tags.

### Qué hacer para que los usuarios reciban la actualización

1. **Crear una Release en GitHub** (no basta con push ni con crear solo un tag):
   - Repo: **https://github.com/zipdatape/bothitss**
   - Ir a **Releases** → **Create a new release**.

2. **Tag de la versión** (obligatorio): crear un tag nuevo, ej. **v1.0.1** (con la "v" delante). La app compara este número con la versión instalada; si el tag es mayor, ofrece actualizar.

3. **Assets opcionales**: si subes un **.exe** como asset de la release, la app puede descargarlo e instalarlo. Si no, al aceptar actualizar se abre la página de la release para descargar manualmente.

4. **Publicar la release**: Publish release. A partir de ahí, instalaciones con versión menor verán "Nueva versión disponible" al abrir.

| Acción en GitHub | Efecto |
|------------------|--------|
| Crear **Release** con tag **v1.0.1** (o superior) | La app detecta actualización al iniciar. |
| Añadir **.exe** como asset | La app puede descargar e instalar al aceptar. |
| Solo push de código (sin release) | No se detecta actualización. |

URL que usa la app: `https://api.github.com/repos/zipdatape/bothitss/releases/latest`.

## Notas

- El flujo es equivalente al del proceso UiPath (carpeta C.H_BAJAS, asunto + fecha, adjuntos, base CSV, backup y correo de notificación).
- No requiere .bat: todo se ejecuta desde la interfaz (botón “Ejecutar proceso”).
- La configuración se guarda en `config.json` en la carpeta de la aplicación, por lo que cada equipo puede tener su propia configuración.
