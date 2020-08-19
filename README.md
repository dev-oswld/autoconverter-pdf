# Visio a PDF 
Script creado en Python para convertir archivos de Visio a PDF por lotes (batch)

## Notas
```python
# Ignorar todas las alertas y cuadros de dialogo
# https://docs.microsoft.com/en-us/office/vba/api/visio.application.alertresponse

visio.AlertResponse = 7
```

```python
# Sintaxis de argumentos
# https://docs.microsoft.com/en-us/office/vba/api/visio.document.exportasfixedformat

document.ExportAsFixedFormat(1, output, 1, 0)
document.Close()
```
- [Docs pywin32](http://timgolden.me.uk/pywin32-docs/contents.html)