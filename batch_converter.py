"""
Batch Excel to PDF Converter para RPA
Procesa múltiples archivos Excel de forma eficiente
"""

import os
import sys
import json
from pathlib import Path
from impresionPDF_improved import ExcelToPdfConverter, logger, config

def process_folder(folder_path: str, output_folder: str = None, pattern: str = "*.xlsx"):
    """
    Procesa todos los archivos Excel en una carpeta
    
    Args:
        folder_path: Carpeta con archivos Excel
        output_folder: Carpeta de salida (opcional)
        pattern: Patrón de archivos (ej: "*.xlsx", "*.xls")
    """
    folder_path = Path(folder_path)
    
    if not folder_path.exists():
        logger.error(f"Carpeta no encontrada: {folder_path}")
        return {}
    
    # Buscar archivos Excel
    excel_files = list(folder_path.glob(pattern))
    logger.info(f"Encontrados {len(excel_files)} archivos Excel")
    
    if not excel_files:
        logger.warning("No se encontraron archivos Excel")
        return {}
    
    # Configurar carpeta de salida
    if output_folder is None:
        output_folder = folder_path / "pdf_output"
    
    output_folder = Path(output_folder)
    output_folder.mkdir(exist_ok=True)
    
    # Procesar archivos
    with ExcelToPdfConverter() as converter:
        results = {}
        
        for excel_file in excel_files:
            try:
                pdf_file = output_folder / f"{excel_file.stem}.pdf"
                result = converter.convert_excel_to_pdf(str(excel_file), str(pdf_file))
                results[str(excel_file)] = result
                logger.info(f"✅ {excel_file.name} -> {pdf_file.name}")
                
            except Exception as e:
                logger.error(f"❌ Error en {excel_file.name}: {e}")
                results[str(excel_file)] = f"ERROR: {e}"
        
        return results

def process_file_list(file_list_path: str, output_folder: str = None):
    """
    Procesa archivos desde una lista en archivo JSON
    
    Args:
        file_list_path: Archivo JSON con lista de archivos
        output_folder: Carpeta de salida
    """
    try:
        with open(file_list_path, 'r', encoding='utf-8') as f:
            file_list = json.load(f)
    except Exception as e:
        logger.error(f"Error cargando lista de archivos: {e}")
        return {}
    
    if output_folder:
        output_folder = Path(output_folder)
        output_folder.mkdir(exist_ok=True)
    
    with ExcelToPdfConverter() as converter:
        return converter.convert_batch(file_list, str(output_folder) if output_folder else None)

def main():
    """Función principal para procesamiento por lotes"""
    if len(sys.argv) < 2:
        print("Uso:")
        print("  python batch_converter.py <carpeta> [carpeta_salida]")
        print("  python batch_converter.py --list <archivo_lista.json> [carpeta_salida]")
        return
    
    logger.info("=== Iniciando Procesamiento por Lotes ===")
    
    try:
        if sys.argv[1] == "--list" and len(sys.argv) > 2:
            # Procesar desde lista de archivos
            file_list_path = sys.argv[2]
            output_folder = sys.argv[3] if len(sys.argv) > 3 else None
            
            results = process_file_list(file_list_path, output_folder)
            
        else:
            # Procesar carpeta
            folder_path = sys.argv[1]
            output_folder = sys.argv[2] if len(sys.argv) > 2 else None
            
            results = process_folder(folder_path, output_folder)
        
        # Resumen de resultados
        total = len(results)
        successful = len([r for r in results.values() if not str(r).startswith("ERROR")])
        
        logger.info(f"=== Resumen ===")
        logger.info(f"Total archivos: {total}")
        logger.info(f"Exitosos: {successful}")
        logger.info(f"Con errores: {total - successful}")
        
        # Guardar reporte
        report_file = f"batch_report_{int(time.time())}.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        
        logger.info(f"Reporte guardado en: {report_file}")
        
    except Exception as e:
        logger.error(f"Error en procesamiento por lotes: {e}")
        sys.exit(1)

if __name__ == "__main__":
    import time
    main()
