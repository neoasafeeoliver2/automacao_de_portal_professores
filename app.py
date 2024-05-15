from docx import Document
import re
import pyperclip
import pyautogui
from time import sleep
itens = []
def read_table_from_docx(file_path):
    doc = Document(file_path)
    tables = doc.tables
    for table in tables:
        for row in table.rows[1:]:
            for cell in row.cells:
                codigo = [re.sub(r'[\n\r]+', '', cell.text)]
                itens.append(codigo)

file_path = "22-27.04.docx"
read_table_from_docx(file_path)
lista = itens[1]
strings = ''
for elemento in itens[1]:
    strings += elemento + ' '
codigos = re.findall(r'\((.*?)\)', strings)

def processo_colar(click):
    sleep(1)
    pyautogui.click(click, duration=1)
    pyautogui.hotkey('tab')
    sleep(0.2)
    pyautogui.hotkey('tab')
    

txt1 = ''.join(itens[2]) 
txt2 = ''.join(itens[4])
txt3 = ''.join(itens[3])
# VERIFICANDO CÓDIGOS

# primeiro código:

# VERIFICANDO EO
if 'EO' in codigos[0][4:]:
        sleep(3)
        pyautogui.click(179,334, duration=1.5)
        pyautogui.hotkey('down')
        pyautogui.hotkey('enter')
        sleep(1)
        if '01' in codigos[0][4:]:  # ELE PRECISA LER A PARTIR DO 3 CARACTERE
            #entrar no campo
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
        if '02' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
        if '03' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=3)
            pyautogui.hotkey('enter')            
        if '04' in codigos[0][4:]:    
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=4)
            pyautogui.hotkey('enter')            
        if '05' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=5)
            pyautogui.hotkey('enter')         
        if '06' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=6)
            pyautogui.hotkey('enter') 
            pyautogui.press('home', presses=2)            
        if '07' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=7)
            pyautogui.hotkey('enter') 
            pyautogui.press('home', presses=2)                    
# VERIFICANDO CG
if 'CG' in codigos[0][4:]:
        sleep(3)
        pyautogui.click(179,334, duration=1.5)
        pyautogui.press('down', presses=2)
        pyautogui.hotkey('enter') 
        sleep(1)
        if '01' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
        if '02' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
        if '03' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=3)
            pyautogui.hotkey('enter')             
        if '04' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=4)
            pyautogui.hotkey('enter')           
        if '05' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=5)
            pyautogui.hotkey('enter') 
# VERIFICANDO TS
if 'TS' in codigos[0][4:]:
        sleep(3)
        pyautogui.click(179,334, duration=1.5)
        pyautogui.press('down', presses=3)
        pyautogui.hotkey('enter') 
        sleep(1)
        if '01' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
        if '02' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
        if '03' in codigos[0][4:]:
            pyautogui.hotkey('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.hotkey('tab', presses=3)
            pyautogui.hotkey('enter') 
# VERIFICANDO EF
if 'EF' in codigos[0][4:]:
        sleep(4)
        pyautogui.click(179,334, duration=1.5)
        pyautogui.press('down', presses=4)
        pyautogui.hotkey('enter')
        sleep(1) 
        if '01' in codigos[0][4:]:  # ELE PRECISA LER A PARTIR DO 3 CARACTERE
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
        if '02' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
        if '03' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=3)
            pyautogui.hotkey('enter')           
        if '04' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=4)
            pyautogui.hotkey('enter')              
        if '05' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=5)
            pyautogui.hotkey('enter')            
        if '06' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=6)
            pyautogui.hotkey('enter')  
            pyautogui.press('home', presses=2)           
        if '07' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tapress', presses=7)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
        if '08' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=8)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
        if '09' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=9)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
# VERIFICANDO ET
if 'ET' in codigos[0][4:]:
        sleep(3)
        pyautogui.click(179,334, duration=1.5)
        pyautogui.hotkey('down', presses=5)
        pyautogui.hotkey('enter') 
        sleep(1)
        if '01' in codigos[0][4:]:  # ELE PRECISA LER A PARTIR DO 3 CARACTERE
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.hotkey('tab')
            pyautogui.hotkey('enter')
        if '02' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
        if '03' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=3)
            pyautogui.hotkey('enter')           
        if '04' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=4)
            pyautogui.hotkey('enter')              
        if '05' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=5)
            pyautogui.hotkey('enter')            
        if '06' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=6)
            pyautogui.hotkey('enter') 
            pyautogui.press('home', presses=2)            
        if '07' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tapress', presses=7)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
        if '08' in codigos[0][4:]:
            pyautogui.press('tab', presses=2)
            pyautogui.hotkey('enter')
            sleep(0.5)
            pyautogui.press('tab', presses=8)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)

sleep(1)

if len(codigos) == 2:
    #selecionando segundo código:
    sleep(2)
    pyautogui.click(179,334, duration=1.5)
    # VERIFICANDO EO
    if 'EO' in codigos[1][4:]:
            sleep(2)
            pyautogui.click(164,365, duration=1)
            sleep(2)
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
            sleep(1)
            if '01' in codigos[1][4:]:  # ELE PRECISA LER A PARTIR DO 3 CARACTERE
                #entrar no campo
                pyautogui.press('tab', presses=3)
                sleep(0.5)
                pyautogui.hotkey('enter')
            if '02' in codigos[1][4:]:
                pyautogui.press('tab', presses=4)
                sleep(0.5)
                pyautogui.hotkey('enter')
            if '03' in codigos[1][4:]:
                pyautogui.press('tab', presses=5)
                sleep(0.5)
                pyautogui.hotkey('enter')      
            if '04' in codigos[1][4:]:
                pyautogui.press('tab', presses=6)
                sleep(0.5)
                pyautogui.hotkey('enter')                       
            if '05' in codigos[1][4:]:
                pyautogui.press('tab', presses=7)
                sleep(0.5)
                pyautogui.hotkey('enter')     
            if '06' in codigos[1][4:]:   
                pyautogui.press('tab', presses=8)
                sleep(0.5)                
                pyautogui.hotkey('enter')    
                pyautogui.press('home', presses=2)    
            if '07' in codigos[1][4:]:
                pyautogui.press('tab', presses=9)
                sleep(0.5)
                pyautogui.hotkey('enter') 
                pyautogui.press('home', presses=2)  
    # VERIFICANDO CG
    if 'CG' in codigos[1][4:]:
            sleep(2)
            pyautogui.click(162,403, duration=1)
            sleep(2)
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
            sleep(1)
            if '01' in codigos[1][4:]:
                # Entrando no campo
                pyautogui.press('tab', presses=3)
                sleep(0.5)
                # selecionando
                pyautogui.hotkey('enter')
            if '02' in codigos[1][4:]:
                pyautogui.press('tab', presses=4)
                sleep(0.5)
                pyautogui.hotkey('enter')
            if '03' in codigos[1][4:]:
                pyautogui.press('tab', presses=5)
                sleep(0.5)
                pyautogui.hotkey('enter')        
            if '04' in codigos[1][4:]:
                pyautogui.press('tab', presses=6)
                sleep(0.5)
                pyautogui.hotkey('enter')      
            if '05' in codigos[1][4:]:
                pyautogui.press('tab', presses=7)
                sleep(0.5)
                pyautogui.hotkey('enter')            
    # VERIFICANDO TS
    if 'TS' in codigos[1][4:]:
            sleep(2)
            pyautogui.click(173,434, duration=1)
            sleep(2)
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
            sleep(1)
            if '01' in codigos[1][4:]:
                # Entrando no campo
                pyautogui.press('tab', presses=3)
                sleep(0.5)
                # selecionando
                pyautogui.hotkey('enter')
            if '02' in codigos[1][4:]:
                pyautogui.press('tab', presses=4)
                sleep(0.5)
                pyautogui.hotkey('enter')
            if '03' in codigos[1][4:]:
                pyautogui.press('tab', presses=5)
                sleep(0.5)
                pyautogui.hotkey('enter')
    # VERIFICANDO EF
    if 'EF' in codigos[1][4:]:
        if 'EF' in codigos[0][4:]:
            sleep(2)
            pyautogui.click(179,334, duration=1.5)
            pyautogui.click(179,334, duration=1.5)
            pyautogui.press('tab', presses=2)
        else:
            pyautogui.click(184,449, duration=1)
            sleep(2)
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
            sleep(1)
        if '01' in codigos[1][4:]:  # ELE PRECISA LER A PARTIR DO 3 CARACTERE
            # Entrando no campo
            pyautogui.press('tab', presses=1)
            sleep(0.5)
            # selecionando
            pyautogui.hotkey('enter') 
        if '02' in codigos[1][4:]:
            # Entrando no campo
            pyautogui.press('tab', presses=2)
            sleep(0.5)
            # selecionando
            pyautogui.hotkey('enter')
        if '03' in codigos[1][4:]:
                # Entrando no campo
            pyautogui.press('tab', presses=3)
            sleep(0.5)
            # selecionando
            pyautogui.hotkey('enter')            
        if '04' in codigos[1][4:]:
            pyautogui.press('tab', presses=4)
            sleep(0.5)
            pyautogui.hotkey('enter')             
        if '05' in codigos[1][4:]:
            pyautogui.press('tab', presses=5)
            sleep(0.5)
            pyautogui.hotkey('enter')            
        if '06' in codigos[1][4:]:
            pyautogui.press('tab', presses=6)
            sleep(0.5)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)           
        if '07' in codigos[1][4:]:
            pyautogui.press('tab', presses=7)
            sleep(0.5)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
        if '08' in codigos[1][4:]:
            pyautogui.hotkey('tab', presses=8)
            sleep(0.5)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
        if '09' in codigos[1][4:]:
            pyautogui.hotkey('tab', presses=9)
            sleep(0.5)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
    # VERIFICANDO ET
    if 'ET' in codigos[1][4:]:
        sleep(2)
        pyautogui.click(193,473, duration=1)
        sleep(2)
        pyautogui.press('tab', presses=2)
        pyautogui.press('enter')
        sleep(1)
        if '01' in codigos[1][4:]:  
            #entrar no campo
            pyautogui.press('tab', presses=3)
            sleep(0.5)
            pyautogui.hotkey('enter')
            sleep(1)    
        if '02' in codigos[1][4:]:
            pyautogui.press('tab', presses=4)
            sleep(0.5)
            pyautogui.hotkey('enter')
        if '03' in codigos[1][4:]:
            pyautogui.press('tab', presses=5)
            sleep(0.5)
            pyautogui.hotkey('enter')
        if '04' in codigos[1][4:]:
            pyautogui.press('tab', presses=6)
            sleep(0.5)
            pyautogui.hotkey('enter')            
        if '05' in codigos[1][4:]:
            pyautogui.press('tab', presses=7)
            sleep(0.5)
            pyautogui.hotkey('enter')
        if '06' in codigos[1][4:]:
            pyautogui.press('tab', presses=8)
            sleep(0.5)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
        if '07' in codigos[1][4:]:
            pyautogui.press('tab', presses=9)
            sleep(0.5)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)
        if '08' in codigos[1][4:]:
            pyautogui.press('tab', presses=10)
            sleep(0.5)
            pyautogui.hotkey('enter')
            pyautogui.press('home', presses=2)


# MINIMIZANDO E CLICANDO EM TAB - METODOLOGIAS
sleep(1)
# MINIMIZANDO 
pyautogui.click(90,375, duration=1)
#CLICK TAB
pyautogui.hotkey('tab')
sleep(0.2)
pyautogui.hotkey('tab')
sleep(0.2)
pyautogui.hotkey('enter')
sleep(1)
pyautogui.hotkey('tab')
sleep(0.2)
pyautogui.hotkey('enter')
sleep(0.2)
pyautogui.press('end', presses=2)
sleep(0.8)
# CLICANDO EM +
if len(codigos) == 2:
    click1 = 1142,467
else:
    click1 = 1145,407
processo_colar(click1)
# COPIANDO E COLANDO TEXTO
pyperclip.copy(txt1)
pyautogui.hotkey('ctrl', 'shft', 'v')

# CLICANDO EM AVALIAÇÕES
pyautogui.click(91,523, duration=1.5)
sleep(1)
pyautogui.hotkey('tab')
sleep(0.2)
pyautogui.hotkey('enter')
sleep(1)
pyautogui.click(865,413, duration=1)
pyautogui.press('end', presses=2)
sleep(0.8)
click2 = 1146,432
processo_colar(click2)
pyperclip.copy(txt2)
pyautogui.hotkey('ctrl', 'shft', 'v')

# MINIMIZANDO E CLICANDO EM RECURSOS
pyautogui.click(89,263, duration=1)
pyautogui.hotkey('tab')
sleep(0.3)
pyautogui.hotkey('enter')
sleep(1)
pyautogui.hotkey('tab')
sleep(0.3)
pyautogui.hotkey('enter')
sleep(1)
pyautogui.hotkey('tab')
sleep(0.3)
pyautogui.hotkey('enter')
sleep(0.3)
pyautogui.click(865,413, duration=1)
pyautogui.press('end', presses=2)
sleep(0.8)
click3 = 1146,513
processo_colar(click3)
pyperclip.copy(txt3)
pyautogui.hotkey('ctrl', 'shft', 'v')

# [['(22/04/2024)Segunda-feira'] 0
# ['(EI02EO04)'], ['(EI02EO05)'] 1
# ['As crianças assistirão o vídeo “Papai Dedo” depois faremos comentários sobre a família de cada uma. Quantos membros moram em sua casa, se possui irmãos, o nome de sua mãe, pai e irmãos. Em seguida, desenharão as pessoas que moram com elas. A proposta para a família será o desenho da mão da mãe, pai ou outro membro que more com ela. '] 2
# ['TvLápis Borracha CelularCaixa de som'] 3
# ['Demonstrar afetividade pelos membros de sua família.']]