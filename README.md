# Registro de No Conforme

Formulario en línea para capturar no conformes. Las listas desplegables salen de `VALIDACIONES.xlsx` y, al escribir la orden, se llenan **Cliente**, **Código** y **Producto** desde la hoja `rep-ov-espec` de Seguimiento de OT.

## Cómo usarlo

1. Abre el sitio publicado en GitHub Pages.
2. Escribe la orden (ID Orden de Venta u Orden de Compra).
3. Si la orden tiene varios productos, elige uno.
4. Completa el resto y pulsa **Guardar Registro**.
5. **Exportar registros** descarga un CSV de lo guardado en ese navegador.

Los registros se quedan en el navegador (no en GitHub). Si cambias de equipo o de Chrome, no van a aparecer.

## Actualizar los datos desde Excel

En la carpeta del proyecto:

```
python scripts/actualizar-datos.py
```

Lee `VALIDACIONES.xlsx` y `SEGUIMIENTO DE OT1 (1).xlsx` desde el Escritorio y regenera `data/validaciones.json` y `data/ordenes.json`. Al repositorio solo se suben esos JSON; los Excel se quedan en la PC.
