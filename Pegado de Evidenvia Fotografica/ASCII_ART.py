import shutil
import time
import re

# Habilita códigos ANSI en Windows (opcional si no está instalado)
try:
    from colorama import init as _colorama_init
    _colorama_init()
except Exception:
    pass

# ============================
# CONFIGURACIÓN GENERAL
# ============================

# Paleta de colores 
# #ECD925 → código 226 en ANSI 256
HEX_COLOR = "\033[38;5;226m"  # color #ECD925 en ANSI 256
RESET = "\033[0m"

# ============================
# FIGURITA MAMALONA PERSONALIZABLE
# ============================

ASCII_ART = f"""
{HEX_COLOR}
                                                                                                    
   @@@@@@@@@                     @@@@@@@@@          @@@@@@@@                 @@@@@@@@@@@@@@@@@      
   @@@@@@@@@                    @@@@@@@@@@             @@@@@@@@          @@@@@@@@@@@@@@@@@@@@@@@    
    @@@@@@@@@                   @@@@@@@@@    @@@@@       @@@@@@@@     @@@@@@@@@@@@@@@@@@@@@@@@@@    
    @@@@@@@@@@                 @@@@@@@@@    @@@@@@@@     @@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@     
     @@@@@@@@@                @@@@@@@@@    @@@@@@@@@@     @@@@@@@@@@@@@@@@@@@@@             @@@     
      @@@@@@@@@               @@@@@@@@@    @@@@@@@@@      @@@@@@@@@@@@@@@@@@                        
      @@@@@@@@@@             @@@@@@@@@     @@@@@@@@      @@@@@@@@@@@@@@@@@                          
       @@@@@@@@@            @@@@@@@@@@      @@@@@@      @@@@@@@@@@@@@@@@@                           
       @@@@@@@@@@           @@@@@@@@@        @@@      @@@@@@@@@@@@@@@@@@                            
        @@@@@@@@@          @@@@@@@@@                @@@@@@@@@@@@@@@@@@@@                            
         @@@@@@@@@         @@@@@@@@              @@@@@@@@@@@@@@    @@@@                             
         @@@@@@@@@        @@@@@@@@@              @@@@@@@@@@@@@     @@@@                             
          @@@@@@@@@      @@@@@@@@@        @@@      @@@@@@@@@@@     @@@@                             
           @@@@@@@@@     @@@@@@@@@      @@@@@@      @@@@@@@@@     @@@@@                             
           @@@@@@@@@    @@@@@@@@@      @@@@@@@@       @@@@@@@     @@@@@@                            
            @@@@@@@@@   @@@@@@@@      @@@@@@@@@@@       @@@@     @@@@@@@                            
            @@@@@@@@@  @@@@@@@@       @@@@@@@@@@@@@      @@     @@@@@@@@@                           
             @@@@@@@@@ @@@@@@@@       @@@@@@@@@@@@@@           @@@@@@@@@@@                          
              @@@@@@@@@@@@@@@@         @@@@@@@@@@@@@@@        @@@@@@@@@@@@@@                        
              @@@@@@@@@@@@@@@@           @@@@@@@@@@@@         @@@@@@@@@@@@@@@@@               @@    
               @@@@@@@@@@@@@@               @@@@                @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@    
               @@@@@@@@@@@@@                           @@@@        @@@@@@@@@@@@@@@@@@@@@@@@@@@@@    
                @@@@@@@@@@@@                        @@@@@@@@          @@@@@@@@@@@@@@@@@@@@@@@@@@@   
                 @@@@@@@@@@                       @@@@@@@@@@@@             @@@@@@@@@@@@@@@@@@@
{RESET}
"""

# ============================
# FUNCIONES PRINCIPALES
# ============================

def get_width():
    try:
        return shutil.get_terminal_size().columns
    except:
        return 80

def strip_ansi(text):
    ansi = re.compile(r'\x1B\[[0-?]*[ -/]*[@-~]')
    return ansi.sub('', text)

def center_text(text):
    width = get_width()
    lines = text.split("\n")
    centered = []

    for line in lines:
        real = strip_ansi(line)
        padding = max(0, (width - len(real)) // 2)
        centered.append(" " * padding + line)

    return "\n".join(centered)

def print_ascii(text, typing=False, speed=0.001):
    centered = center_text(text)
    if not typing:
        print(centered)
    else:
        for char in centered:
            print(char, end="", flush=True)
            time.sleep(speed)

# ============================
# EJECUCIÓN
# ============================

if __name__ == "__main__":
    print_ascii(ASCII_ART, typing=False)

# ============================
# FIN DEL ARCHIVO