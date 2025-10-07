"""Excel to PDF Converter - Versión RPA para Automation Anywhere"""

import win32com.client as win32
import os
import sys
import json
import time
import signal
import threading
import logging
import psutil
from pathlib import Path
from datetime import datetime
from contextlib import contextmanager
from typing import Optional, List, Dict, Union
import pythoncom

class RPALogger:
    def __init__(self, log_level=logging.INFO):
        self.logger = logging.getLogger('ExcelToPDF_RPA')
        self.logger.setLevel(log_level)
        
        if not self.logger.handlers:
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            
            file_handler = logging.FileHandler(
                f'excel_to_pdf_rpa_{datetime.now().strftime("%Y%m%d")}.log',
                encoding='utf-8'
            )
            file_handler.setFormatter(formatter)
            
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(formatter)
            
            self.logger.addHandler(file_handler)
            self.logger.addHandler(console_handler)
    
    def info(self, msg): self.logger.info(msg)
    def error(self, msg): self.logger.error(msg)
    def warning(self, msg): self.logger.warning(msg)
    def debug(self, msg): self.logger.debug(msg)

class RPAConfig:
    def __init__(self, config_file='rpa_config.json'):
        self.config_file = config_file
        self.config = self._load_default_config()
        self._load_config()
    
    def _load_default_config(self):
        return {
            "excel": {
                "timeout_seconds": 300,
                "retry_attempts": 3,
                "retry_delay": 5,
                "visible": False,
                "display_alerts": False
            },
            "pdf": {
                "default_quality": 600,
                "default_orientation": "auto",
                "margins_cm": 1.0,
                "fit_to_width": True
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
                "kill_orphaned_excel": True,
                "max_excel_processes": 5
            }
        }
    
    def _load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    file_config = json.load(f)
                    self.config.update(file_config)
                    print(f"Configuración cargada desde {self.config_file}")
            except Exception as e:
                print(f"Error cargando configuración: {e}")
    
    def get(self, key_path: str, default=None):
        keys = key_path.split('.')
        value = self.config
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
        return value

config = RPAConfig()
logger = RPALogger(getattr(logging, config.get('rpa.log_level', 'INFO')))

class TimeoutError(Exception):
    pass

class ExcelProcessManager:
    @staticmethod
    def kill_orphaned_excel_processes():
        try:
            excel_processes = [p for p in psutil.process_iter(['pid', 'name']) 
                             if p.info['name'] and 'excel' in p.info['name'].lower()]
            
            max_processes = config.get('cleanup.max_excel_processes', 5)
            
            if len(excel_processes) > max_processes:
                logger.warning(f"Demasiados procesos Excel ({len(excel_processes)}), limpiando...")
                for proc in excel_processes[max_processes:]:
                    try:
                        proc.kill()
                        logger.info(f"Proceso Excel {proc.info['pid']} terminado")
                    except:
                        pass
        except Exception as e:
            logger.error(f"Error limpiando procesos Excel: {e}")
    
    @staticmethod
    def is_excel_installed():
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Quit()
            return True
        except:
            return False

class ExcelToPdfConverter:
    def __init__(self):
        self.excel = None
        self.initialized = False
        self.logger = logger
        self.config = config
        
    def initialize_excel(self):
        try:
            if config.get('cleanup.kill_orphaned_excel', True):
                ExcelProcessManager.kill_orphaned_excel_processes()
            
            if not ExcelProcessManager.is_excel_installed():
                raise Exception("Microsoft Excel no está instalado o no es accesible")
            
            pythoncom.CoInitialize()
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')
            self.excel.Visible = config.get('excel.visible', False)
            self.excel.DisplayAlerts = config.get('excel.display_alerts', False)
            self.excel.ScreenUpdating = False
            
            self.initialized = True
            self.logger.info("Excel inicializado correctamente")
            
        except Exception as e:
            self.logger.error(f"Error inicializando Excel: {e}")
            raise
    
    def validate_file_permissions(self, file_path: str) -> bool:
        try:
            if not os.path.exists(file_path):
                return False
            
            with open(file_path, 'rb') as f:
                f.read(1)
            
            parent_dir = os.path.dirname(file_path)
            test_file = os.path.join(parent_dir, f"test_write_{int(time.time())}.tmp")
            try:
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                return True
            except:
                return False
                
        except Exception as e:
            self.logger.error(f"Error validando permisos: {e}")
            return False
    
    def configure_print_settings(self, ws):
        try:
            ps = ws.PageSetup
            
            orientation = config.get('pdf.default_orientation', 'auto')
            if orientation == 'auto':
                used_range = ws.UsedRange
                num_cols = used_range.Columns.Count
                num_rows = used_range.Rows.Count
                ps.Orientation = 2 if num_cols > num_rows or num_cols > 10 else 1
            else:
                ps.Orientation = 2 if orientation.lower() == 'landscape' else 1
            
            ps.Zoom = False
            ps.FitToPagesWide = 1 if config.get('pdf.fit_to_width', True) else False
            ps.FitToPagesTall = False
            
            margin_cm = config.get('pdf.margins_cm', 1.0)
            margin_points = self.excel.CentimetersToPoints(margin_cm)
            ps.LeftMargin = margin_points
            ps.RightMargin = margin_points
            ps.TopMargin = margin_points
            ps.BottomMargin = margin_points
            ps.HeaderMargin = self.excel.CentimetersToPoints(0.5)
            ps.FooterMargin = self.excel.CentimetersToPoints(0.5)
            
            ps.CenterHorizontally = True
            ps.CenterVertically = False
            ps.PrintQuality = config.get('pdf.default_quality', 600)
            
            if ws.PageSetup.PrintArea == "":
                used_range = ws.UsedRange
                ws.PageSetup.PrintArea = used_range.Address
                
        except Exception as e:
            self.logger.error(f"Error configurando impresión: {e}")
            raise
    
    def convert_excel_to_pdf(self, excel_path: str, pdf_path: Optional[str] = None, 
                           optimize_orientation: bool = True) -> str:
        wb = None
        start_time = time.time()
        
        try:
            excel_path = os.path.abspath(excel_path)
            if not self.validate_file_permissions(excel_path):
                raise FileNotFoundError(f"Archivo no encontrado o sin permisos: {excel_path}")
            
            if pdf_path is None:
                pdf_path = str(Path(excel_path).with_suffix('.pdf'))
            pdf_path = os.path.abspath(pdf_path)
            
            free_space = psutil.disk_usage(os.path.dirname(pdf_path)).free
            if free_space < 100 * 1024 * 1024:
                raise Exception("Espacio insuficiente en disco")
            
            self.logger.info(f"Iniciando conversión: {excel_path} -> {pdf_path}")
            
            if not self.initialized:
                self.initialize_excel()
            
            retry_attempts = config.get('excel.retry_attempts', 3)
            retry_delay = config.get('excel.retry_delay', 5)
            
            for attempt in range(retry_attempts):
                try:
                    wb = self.excel.Workbooks.Open(excel_path)
                    break
                except Exception as e:
                    if attempt < retry_attempts - 1:
                        self.logger.warning(f"Intento {attempt + 1} falló, reintentando en {retry_delay}s: {e}")
                        time.sleep(retry_delay)
                    else:
                        raise
                
            self.logger.info(f"Libro abierto: {wb.Name}, Hojas: {wb.Worksheets.Count}")
            
            for i in range(1, wb.Worksheets.Count + 1):
                ws = wb.Worksheets(i)
                self.logger.debug(f"Configurando hoja: {ws.Name}")
                self.configure_print_settings(ws)
            
            self.logger.info("Exportando a PDF...")
            wb.ExportAsFixedFormat(0, pdf_path)
            
            if not os.path.exists(pdf_path):
                raise Exception("PDF no se generó correctamente")
            
            elapsed_time = time.time() - start_time
            file_size = os.path.getsize(pdf_path)
            
            self.logger.info(f"Conversión exitosa: {pdf_path}")
            self.logger.info(f"Tiempo: {elapsed_time:.2f}s, Tamaño: {file_size/1024:.1f}KB")
            
            return pdf_path
                
        except FileNotFoundError as e:
            self.logger.error(f"Archivo no encontrado: {e}")
            sys.exit(config.get('rpa.exit_codes.file_not_found', 1))
        except PermissionError as e:
            self.logger.error(f"Error de permisos: {e}")
            sys.exit(config.get('rpa.exit_codes.permission_error', 4))
        except Exception as e:
            if "timeout" in str(e).lower():
                self.logger.error(f"Timeout en la operación: {e}")
                sys.exit(config.get('rpa.exit_codes.timeout', 3))
            self.logger.error(f"Error durante conversión: {e}")
            sys.exit(config.get('rpa.exit_codes.general_error', 99))
        finally:
            if wb:
                try:
                    wb.Close(False)
                    self.logger.debug("Libro cerrado")
                except:
                    pass
    
    def convert_batch(self, file_list: List[str], output_dir: Optional[str] = None) -> Dict[str, str]:
        results = {}
        
        for excel_file in file_list:
            try:
                if output_dir:
                    filename = Path(excel_file).stem + '.pdf'
                    pdf_path = os.path.join(output_dir, filename)
                else:
                    pdf_path = None
                
                result = self.convert_excel_to_pdf(excel_file, pdf_path)
                results[excel_file] = result
                self.logger.info(f"✅ {excel_file} -> {result}")
                
            except Exception as e:
                self.logger.error(f"❌ Error en {excel_file}: {e}")
                results[excel_file] = f"ERROR: {e}"
        
        return results
    
    def close(self):
        if self.excel and self.initialized:
            try:
                self.excel.ScreenUpdating = True
                self.excel.Quit()
                self.excel = None
                pythoncom.CoUninitialize()
                self.initialized = False
                self.logger.info("Excel cerrado correctamente")
            except Exception as e:
                self.logger.error(f"Error cerrando Excel: {e}")
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

def main():
    try:
        if len(sys.argv) > 1:
            log_level = sys.argv[1].upper()
            if hasattr(logging, log_level):
                logger.logger.setLevel(getattr(logging, log_level))
        
        logger.info("=== Iniciando Excel to PDF Converter RPA ===")
        
        with ExcelToPdfConverter() as converter:
            if len(sys.argv) > 2:
                excel_file = sys.argv[2]
                pdf_file = sys.argv[3] if len(sys.argv) > 3 else None
            else:
                excel_file = r"C:\Users\erlin\Downloads\Datos.xlsx"
                pdf_file = r"C:\Users\erlin\Downloads\salida_mejorada.pdf"
            
            result = converter.convert_excel_to_pdf(excel_file, pdf_file)
            logger.info(f"✅ Proceso completado exitosamente: {result}")
            
            sys.exit(config.get('rpa.exit_codes.success', 0))
            
    except Exception as e:
        logger.error(f"Error en proceso principal: {e}")
        sys.exit(config.get('rpa.exit_codes.general_error', 99))

if __name__ == "__main__":
    main()
