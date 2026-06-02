# Changelog — main

## [No publicado]

### Cambiado
- Se renombró "Tipo de Proceso" a "Tipo de Estudio" en toda la aplicación (UI, backend, headers del Sheet)
- Se cambió el campo de selección única (select) a selección múltiple (checkboxes) para permitir elegir varios tipos de estudio por solicitud
- Se actualizó la validación para verificar que al menos un tipo de estudio esté seleccionado
- Se actualizó la lógica de anexo obligatorio: si al menos uno de los tipos seleccionados requiere anexo, el documento es obligatorio
- Se actualizó el nombre del archivo en Drive para incluir todos los tipos de estudio seleccionados (separados por coma)
- Se renombró la constante `TIPOS_PROCESO` a `TIPOS_ESTUDIO` en el backend y `getTiposProceso()` a `getTiposEstudio()`
