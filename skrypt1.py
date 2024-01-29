from pywinauto import application
import time
import sys
from pywinauto.timings import wait_until, TimeoutError
from PIL import ImageGrab, ImageEnhance
import pytesseract
import os
import math

ergo_app = application.Application().start(r"C:\Users\Filip\Desktop\asd\ERGO3\winfim.exe")



login = ergo_app.ListaOperatorow.child_window(class_name="TClientGrid")
login.send_keystrokes("{DOWN 8}")  
login.type_keys("{ENTER}")

# ergo_app.Sprawdzeniehasla.print_control_identifiers()
passwd_window = ergo_app.Sprawdzeniehasla.child_window(class_name="TLPanel")
passwd=passwd_window.window(class_name="TLMaskEdit")
passwd.set_text("111")
passwd=passwd_window.window(title="Akceptuję", class_name="TBitBtn")
passwd.click()

identifiers_output = sys.stdout

with open('identifiers_output.txt', 'w') as file:    
    sys.stdout = file        
    ergo_app.ErgoFakturyiMagazyn.print_control_identifiers()
sys.stdout = identifiers_output

menu_towary = ergo_app.ErgoFakturyiMagazyn.menu_select("Towary -> [3]")

towary=ergo_app.ListaTowarow.window(class_name="TClientGrid")
#preview=ergo_app.Podglad.print_control_identifiers()
########################
towary.send_keystrokes("{DOWN 170}") 
# towary.send_keystrokes("{DOWN 67}")
############################################################
# left, top, right, bottom = 1365, 1258, 1476, 1280
left, top, right, bottom = 1420, 1258, 1476, 1280
region = (left, top, right, bottom)
current_directory = os.getcwd()
print("Bieżący katalog roboczy:", current_directory)

t1 = math.floor(time.time())
count = 0

# Ścieżka do pliku
file_path = os.path.join(current_directory, "towary.txt")
with open(file_path, "a") as file:
    if file.tell() == 0:
        file.write("Kod;Nazwa;Cena;Stan\n")

    while True:
        count+=1
        screenshot = ImageGrab.grab(bbox=region)
        enhancer = ImageEnhance.Contrast(screenshot)
        screenshot = enhancer.enhance(2.0)
        text = pytesseract.image_to_string(screenshot).replace(',', '').replace('\n', '').strip()

        if not text:
            text = 0
            #print("Odczytany stan:", text)
        else:
            try:
                text = int(float(text))
                text/=100
            except ValueError:
                text = 0
                print("Nie można przekształcić odczytanego stanu na liczbę całkowitą. Ustawiam stan na 0.")

        if text > 0:
            _t = math.floor(time.time()-t1)
            print("czas=%d, srednio=%f, ilosc=%d" % (_t, _t/count, count))
            towary.type_keys("{ENTER}")
            
            edit_kod= ergo_app.Podglad.child_window(class_name="TLEditStr", found_index=12).window_text()
            edit_nazwa= ergo_app.Podglad.child_window(class_name="TLEditStr", found_index=13).window_text()
            edit_cena= ergo_app.Podglad.child_window(class_name="TLEditNum", found_index=8).window_text()

            new_row = f"{edit_kod};{edit_nazwa};{edit_cena};{text}\n"
            file.write(new_row)
            print(edit_cena, edit_nazwa, edit_kod, int(text))

            ergo_app.Podglad.child_window(title="Rezygnuję", class_name="TBitBtn").click()

        towary.send_keystrokes("{DOWN}")
        time.sleep(0.2)
