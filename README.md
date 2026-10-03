# Despachos · mensajes de WhatsApp

App web de un solo archivo (`index.html`) basada en `Logistica_Viajes_DSY.xlsm`.
Se ingresan los datos del viaje (Fecha, Área, Cliente, Sector, Producto, Transporte, 60º, Capacidad, Procedencia,
Destino, Muelle, Orden Pedido, Hora Orden, Conductor, Placa, Responsable, Hora Inicio) y se envían 6 mensajes,
armados con las mismas fórmulas del libro (msgAsigna, msgCancela, msgDescarga, msgDesigna, msgTurno, Recordatorio):

| Botón | Destinatario | Fuente del celular |
|---|---|---|
| Asignar / Anular / Descargar / Recordar | Conductor | tbChofer |
| Solicitar | Transportista | tbDuenos |
| Coordinar | Coordinador del terminal | tbCoordinador |

Controles: campos obligatorios por mensaje, formato de celular +593, aviso si el destino es el número de respaldo,
si el conductor/placa/transportista no está en las hojas de apoyo, si la capacidad no coincide con la unidad,
si ya se envió; saludo según la hora real; mensaje editable; registro de envíos exportable a CSV.

## Uso
1. `pip install openpyxl`
2. `python scripts/exportar_apoyo.py Logistica_Viajes_DSY.xlsm` → genera `data/apoyo.json`
3. Sube el repo a GitHub y activa Pages (o abre `index.html` desde un servidor).

> ⚠️ `data/apoyo.json` contiene cédulas y celulares de conductores: usa un repositorio **privado**
> (GitHub Pages en repos privados requiere plan de pago) o aloja la página en otro sitio.
