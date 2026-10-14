# Configuración de importación, fotos, QR y alertas por correo

## Base de datos

Ejecuta `alertas_analitica.sql` una vez en la base `biomant`. Añade las rutas de imágenes y la tabla que evita repetir correos para la misma fecha y estado de alerta.

## Importar inventario

La importación está disponible para Administradores desde **Importar Inventario**. Acepta `.xlsx` y PDF con texto seleccionable. El Excel toma la primera hoja y la fila 1 como encabezados. Columnas requeridas:

`id_equipo, nombre_equipo, marca, modelo, numero_serie, nivel_riesgo, ubicacion, estado, tipo_categoria`

Columnas opcionales: `proximo_mto_preventivo` (formato `AAAA-MM-DD`) y `foto_equipo`. Estados permitidos: `Operativo`, `En mantenimiento`, `Fuera de servicio`. Categorías: `Biomédico`, `Tecnología`. Riesgos: `I`, `IIA`, `IIB`, `III`. IDs ya existentes se omiten; las filas inválidas se reportan al final. Se admiten hasta 5.000 filas y archivos de 15 MB.

Para importar un PDF, debe contener una tabla de texto con una fila de encabezados y celdas separadas por `|`, punto y coma, tabuladores o espacios múltiples. El sistema no aplica OCR a PDF escaneados.

## Fotos y códigos QR

Los Administradores suben fotos JPEG, PNG o WebP de hasta 5 MB desde la ficha de cada equipo. Se almacenan en `uploads/equipos` y la base guarda la ruta pública. Para que un QR abra la ficha desde otros dispositivos, configura `APP_PUBLIC_URL` con la dirección HTTP(S) accesible desde esos equipos, por ejemplo la URL de la intranet. El QR abre BioMant y presenta el detalle del equipo después del inicio de sesión.

## Notificaciones por correo

Configura estas variables en `.env` y reinicia el servidor:

```env
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_SECURE=false
SMTP_USER=usuario
SMTP_PASS=contraseña
MAIL_FROM=BioMant <biomant@example.com>
MAINTENANCE_ALERT_CRON=0 8 * * *
APP_TIMEZONE=America/Bogota
APP_PUBLIC_URL=http://servidor-biomant:3000
```

El trabajo diario consulta los preventivos con fecha igual o anterior a 15 días desde `CURRENT_DATE` y envía un resumen a los correos de usuarios con rol Administrador o Técnico. Envía una vez al entrar en el rango de próximos y otra si llega a vencido; cambiar la fecha programada crea una nueva alerta. Si SMTP no está configurado, el servidor deja las notificaciones desactivadas y lo informa en consola.

## Filtros

El filtro de ubicación usa la columna `ubicacion` disponible en el inventario; se muestra como departamento/ubicación porque el esquema actual no tiene una columna separada de hospital o departamento.
