# Excel to PDF Converter - RPA Edition

Convertidor robusto de archivos Excel a PDF optimizado para automatización RPA y Automation Anywhere.

## 🚀 Características

- ✅ **Conversión Excel a PDF** con configuración optimizada
- ✅ **Sistema de logging avanzado** con timestamps y rotación
- ✅ **Manejo robusto de errores** con códigos de salida estándar
- ✅ **Configuración externa** sin hardcoding
- ✅ **Procesamiento por lotes** para múltiples archivos
- ✅ **Limpieza automática** de procesos Excel huérfanos
- ✅ **Validaciones completas** de permisos y espacio en disco
- ✅ **Compatible con Windows** y entornos de producción

## 📁 Estructura del Proyecto

```
├── impresionPDF_improved.py    # Script principal mejorado
├── batch_converter.py          # Procesador por lotes
├── rpa_config.json            # Archivo de configuración
├── README_RPA.md              # Documentación técnica detallada
└── README.md                  # Este archivo
```

## 🔧 Instalación

### Requisitos
- Python 3.7+
- Microsoft Excel instalado
- Windows (para COM automation)

### Dependencias
```bash
pip install pywin32 psutil
```

## 📖 Uso Básico

### Comando Simple
```bash
python impresionPDF_improved.py
```

### Con Parámetros
```bash
python impresionPDF_improved.py INFO "archivo.xlsx" "salida.pdf"
```

### Procesamiento por Lotes
```bash
python batch_converter.py "carpeta_excel" "carpeta_salida"
```

## ⚙️ Configuración

Edita `rpa_config.json` para personalizar:

```json
{
  "excel": {
    "timeout_seconds": 300,
    "retry_attempts": 3,
    "visible": false
  },
  "pdf": {
    "default_quality": 600,
    "default_orientation": "auto",
    "margins_cm": 1.0
  },
  "rpa": {
    "exit_codes": {
      "success": 0,
      "file_not_found": 1,
      "excel_error": 2,
      "timeout": 3,
      "permission_error": 4,
      "general_error": 99
    }
  }
}
```

## 🤖 Uso en Automation Anywhere

### Variables Recomendadas
- `InputFile` - Archivo Excel de entrada
- `OutputFile` - Archivo PDF de salida
- `LogLevel` - Nivel de logging (DEBUG, INFO, WARNING, ERROR)

### Comando en AA
```powershell
python impresionPDF_improved.py %LogLevel% "%InputFile%" "%OutputFile%"
```

### Códigos de Salida
- `0`: Éxito
- `1`: Archivo no encontrado
- `2`: Error de Excel
- `3`: Timeout
- `4`: Error de permisos
- `99`: Error general

## 📊 Ejemplo de Salida

```
2024-01-15 10:30:45 - ExcelToPDF_RPA - INFO - === Iniciando Excel to PDF Converter RPA ===
2024-01-15 10:30:47 - ExcelToPDF_RPA - INFO - Excel inicializado correctamente
2024-01-15 10:30:48 - ExcelToPDF_RPA - INFO - Libro abierto: datos.xlsx, Hojas: 3
2024-01-15 10:30:52 - ExcelToPDF_RPA - INFO - Conversión exitosa: salida.pdf
2024-01-15 10:30:52 - ExcelToPDF_RPA - INFO - Tiempo: 4.2s, Tamaño: 1250.5KB
```

## 🛠️ Troubleshooting

### Problemas Comunes

1. **Excel no responde**
   - El script detecta y mata procesos huérfanos automáticamente
   - Ajusta `timeout_seconds` en configuración

2. **Permisos insuficientes**
   - Verifica permisos de lectura/escritura
   - Ejecuta como administrador si es necesario

3. **Archivos bloqueados**
   - El script valida archivos antes de procesar
   - Implementa reintentos automáticos

## 📈 Rendimiento

- **Archivos pequeños** (< 1MB): ~5-10 segundos
- **Archivos medianos** (1-10MB): ~15-30 segundos  
- **Archivos grandes** (> 10MB): ~30-60 segundos

## 📄 Licencia

MIT License - Libre para uso comercial y personal.

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:
1. Fork el proyecto
2. Crea una rama para tu feature
3. Commit tus cambios
4. Push a la rama
5. Abre un Pull Request

## 📞 Soporte

Para soporte técnico o preguntas sobre integración con RPA, abre un issue en GitHub.
