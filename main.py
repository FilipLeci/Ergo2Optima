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

# with open('identifiers_output.txt', 'w') as file:    
#     sys.stdout = file        
#     ergo_app.ErgoFakturyiMagazyn.print_control_identifiers()
# sys.stdout = identifiers_output

menu_towary = ergo_app.ErgoFakturyiMagazyn.menu_select("Towary -> [3]")

towary=ergo_app.ListaTowarow.window(class_name="TClientGrid")

with open('identifiers_output.txt', 'w') as file:    
    sys.stdout = file        
    ergo_app.ListaTowarow.print_control_identifiers()
sys.stdout = identifiers_output
#preview=ergo_app.Podglad.print_control_identifiers()
########################
towary.send_keystrokes("{DOWN 170}") 
# towary.send_keystrokes("{DOWN 67}")


# 3340x1440
# left_stan, top_stan, right_stan, bottom_stan = 1420, 1258, 1476, 1280
#1920X1080
left_stan, top_stan, right_stan, bottom_stan = 1016, 738, 1090, 764
region_stan = (left_stan, top_stan, right_stan, bottom_stan)

# print("Bieżący katalog roboczy:", current_directory)
# 3340x1440
# left_cena, top_cena, right_cena, bottom_cena = 1420, 1258, 1476, 1280
# 1920X1080
left_cena, top_cena, right_cena, bottom_cena = 214, 200, 320, 675
region_cena = (left_cena, top_cena, right_cena, bottom_cena)
current_directory = os.getcwd()


t1 = math.floor(time.time())
count = 0

# Ścieżka do pliku
file_path = os.path.join(current_directory, "towary.txt")
with open(file_path, "a") as file:
    if file.tell() == 0:
        file.write("Kod;Nazwa;Cena;Stan\n")

    while True:
        count+=1
        screenshot_stan = ImageGrab.grab(bbox=region_stan)
        enhancer_stan = ImageEnhance.Contrast(screenshot_stan)
        screenshot_stan = enhancer_stan.enhance(2.0)
        time.sleep(1.8)
        text_stan = pytesseract.image_to_string(screenshot_stan).replace(',', '').replace('\n', '').strip()
        time.sleep(1.8)

        if not text_stan:
            text_stan = 0
            #print("Odczytany stan:", text)
        else:
            try:
                text_stan = int(float(text_stan))
                text_stan/=100
            except ValueError:
                text_stan = 0
                print("Nie można przekształcić odczytanego stanu na liczbę całkowitą. Ustawiam stan na 0.")

        if text_stan > 0:
            _t = math.floor(time.time()-t1)
            print("czas=%d, srednio=%f, ilosc=%d" % (_t, _t/count, count))
            towary.type_keys("{ENTER}")
            
            edit_kod= ergo_app.Podglad.child_window(class_name="TLEditStr", found_index=12).window_text()
            edit_nazwa= ergo_app.Podglad.child_window(class_name="TLEditStr", found_index=13).window_text()
            edit_cena= ergo_app.Podglad.child_window(class_name="TLEditNum", found_index=8).window_text()

            new_row = f"{edit_kod};{edit_nazwa};{edit_cena};{text_stan}\n"
            file.write(new_row)
            print(edit_cena, edit_nazwa, edit_kod, int(text_stan))

            ergo_app.Podglad.child_window(title="Rezygnuję", class_name="TBitBtn").click()
            # sprawdzenie ceny zakupu/ocr 
            towary.window(title="Stan&y", class_name="TBitBtn").click()
            # screen ceny zakupu
            screenshot_cena = ImageGrab.grab(bbox=region_cena)
            enhancer_cena = ImageEnhance.Contrast(screenshot_cena)
            screenshot_cena = enhancer_cena.enhance(2.0)
            text_cena = pytesseract.image_to_string(screenshot_cena).replace(',', '').replace('\n', '').strip()
            towary.StanywgcenzakupuwmagazynieD.close()

        towary.send_keystrokes("{DOWN}")
        time.sleep(1.8)
