# Excel to PDF Converter - Versión RPA

## 🚀 Características para Automation Anywhere

### ✅ Mejoras Implementadas

1. **Sistema de Logging Avanzado**
   - Logs con timestamps detallados
   - Rotación automática de archivos
   - Niveles configurables (DEBUG, INFO, WARNING, ERROR)

2. **Manejo Robusto de Timeouts**
   - Timeouts configurables por operación
   - Detección automática de procesos colgados
   - Limpieza automática de procesos Excel huérfanos

3. **Configuración Externa**
   - Archivo `rpa_config.json` para configuración
   - Variables de entorno soportadas
   - Configuración sin hardcoding

4. **Códigos de Salida Estándar**
   - 0: Éxito
   - 1: Archivo no encontrado
   - 2: Error de Excel
   - 3: Timeout
   - 4: Error de permisos
   - 99: Error general

5. **Validaciones Robustas**
   - Verificación de Excel instalado
   - Validación de permisos de archivos
   - Verificación de espacio en disco
   - Detección de procesos huérfanos

6. **Procesamiento por Lotes**
   - Conversión de carpetas completas
   - Listas de archivos desde JSON
   - Reportes detallados de resultados

## 📁 Archivos Generados

- `impresionPDF_improved.py` - Script principal mejorado
- `rpa_config.json` - Archivo de configuración
- `batch_converter.py` - Procesador por lotes
- `excel_to_pdf_rpa_YYYYMMDD.log` - Archivos de log

## 🔧 Uso desde Automation Anywhere

### Comando Básico
```powershell
python impresionPDF_improved.py [nivel_log] [archivo_excel] [archivo_pdf]
```

### Ejemplos
```powershell
# Conversión básica
python impresionPDF_improved.py INFO "C:\datos\archivo.xlsx" "C:\salida\archivo.pdf"

# Con logging detallado
python impresionPDF_improved.py DEBUG "C:\datos\archivo.xlsx"

# Procesamiento por lotes
python batch_converter.py "C:\carpeta_excel" "C:\carpeta_pdf"
```

## ⚙️ Configuración

Edita `rpa_config.json` para personalizar:

```json
{
  "excel": {
    "timeout_seconds": 300,
    "retry_attempts": 3,
    "retry_delay": 5,
    "visible": false,
    "display_alerts": false
  },
  "pdf": {
    "default_quality": 600,
    "default_orientation": "auto",
    "margins_cm": 1.0,
    "fit_to_width": true
  },
  "rpa": {
    "exit_codes": {
      "success": 0,
      "file_not_found": 1,
      "excel_error": 2,
      "timeout": 3,
      "permission_error": 4,
      "general_error": 99
    },
    "log_level": "INFO"
  },
  "cleanup": {
    "kill_orphaned_excel": true,
    "max_excel_processes": 5
  }
}
```

## 📊 Monitoreo en AA

### Variables de Salida Recomendadas
- `ExitCode` - Código de salida del script
- `OutputPath` - Ruta del PDF generado
- `LogPath` - Ruta del archivo de log
- `ProcessingTime` - Tiempo de procesamiento

### Logs de Monitoreo
Los logs se guardan automáticamente con formato:
```
2024-01-15 10:30:45 - ExcelToPDF_RPA - INFO - Iniciando conversión: archivo.xlsx
2024-01-15 10:30:47 - ExcelToPDF_RPA - INFO - Conversión exitosa: archivo.pdf
```

## 🛠️ Troubleshooting

### Problemas Comunes

1. **Excel no responde**
   - El script detecta y mata procesos huérfanos automáticamente
   - Ajusta `timeout_seconds` en configuración

2. **Permisos insuficientes**
   - Verifica permisos de lectura/escritura
   - Ejecuta AA como administrador si es necesario

3. **Archivos bloqueados**
   - El script valida archivos antes de procesar
   - Implementa reintentos automáticos

### Códigos de Error
- **Exit Code 1**: Archivo Excel no encontrado
- **Exit Code 2**: Error de Excel (no instalado, corrompido)
- **Exit Code 3**: Timeout (Excel no responde)
- **Exit Code 4**: Permisos insuficientes
- **Exit Code 99**: Error general (ver logs)

## 📈 Rendimiento

### Optimizaciones Implementadas
- Context managers para manejo automático de recursos
- Limpieza automática de procesos Excel
- Reintentos inteligentes con delays
- Validación previa de archivos y permisos
- Procesamiento por lotes eficiente

### Métricas Típicas
- **Archivos pequeños** (< 1MB): ~5-10 segundos
- **Archivos medianos** (1-10MB): ~15-30 segundos
- **Archivos grandes** (> 10MB): ~30-60 segundos

## 🔄 Integración con AA

### Workflow Recomendado
1. **Validación previa**: Verificar existencia de archivos
2. **Ejecución**: Llamar script con parámetros
3. **Verificación**: Comprobar exit code
4. **Logs**: Revisar logs en caso de error
5. **Limpieza**: Archivos temporales si es necesario

### Variables AA Sugeridas
```
vInputFolder = "C:\input"
vOutputFolder = "C:\output"
vLogLevel = "INFO"
vConfigFile = "rpa_config.json"
```
