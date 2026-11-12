"""
Automação do Diálogo de Certificado do Windows
Envia ENTER via Win32 API direta para confirmar o certificado digital selecionado.
"""
import time
import threading
import ctypes

def confirmar_dialogo_certificado_background(delay_segundos: float = 2.0):
    """
    Executa em background o envio de tecla ENTER via Win32 keybd_event,
    sem disparar fail-safe de mouse.
    """
    def _worker():
        time.sleep(delay_segundos)
        print("[*] Enviando tecla ENTER via Win32 API para confirmar Certificado Digital...")
        VK_RETURN = 0x0D
        KEYEVENTF_KEYUP = 0x0002
        
        # Pressionar e soltar ENTER
        ctypes.windll.user32.keybd_event(VK_RETURN, 0, 0, 0)
        time.sleep(0.05)
        ctypes.windll.user32.keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
        
        time.sleep(1.0)
        ctypes.windll.user32.keybd_event(VK_RETURN, 0, 0, 0)
        time.sleep(0.05)
        ctypes.windll.user32.keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
    
    t = threading.Thread(target=_worker, daemon=True)
    t.start()
