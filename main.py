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

towary_grid=ergo_app.ListaTowarow.window(class_name="TClientGrid")

with open('identifiers_output.txt', 'w') as file:    
    sys.stdout = file        
    ergo_app.ListaTowarow.print_control_identifiers()
sys.stdout = identifiers_output

towary_grid.send_keystrokes("{DOWN 170}") 
# towary.send_keystrokes("{DOWN 67}")


# 3340x1440
left_stan, top_stan, right_stan, bottom_stan = 1420, 1258, 1476, 1280
region_stan = (left_stan, top_stan, right_stan, bottom_stan)

current_directory = os.getcwd()


t1 = math.floor(time.time())
count = 0

# Ścieżka do pliku
file_path = os.path.join(current_directory, "towary.txt")
with open(file_path, "a") as file:
    if file.tell() == 0:
        file.write("Kod;Nazwa;Cena;Stan;Cena zakupu\n")

    while True:
        count+=1
        screenshot_stan = ImageGrab.grab(bbox=region_stan)
        enhancer_stan = ImageEnhance.Contrast(screenshot_stan)
        screenshot_stan = enhancer_stan.enhance(2.0)        
        text_stan = pytesseract.image_to_string(screenshot_stan).replace(',', '').replace('\n', '').strip()
       

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
            ergo_app.ListaTowarow.window(title="Pdg", class_name="TBitBtn").click()   
            
            edit_kod= ergo_app.Podglad.child_window(class_name="TLEditStr", found_index=12).window_text()
            edit_nazwa= ergo_app.Podglad.child_window(class_name="TLEditStr", found_index=13).window_text()
            edit_cena= ergo_app.Podglad.child_window(class_name="TLEditNum", found_index=8).window_text()
            ergo_app.Podglad.child_window(title="Rezygnuję", class_name="TBitBtn").click()
            #wejście w stany
            ergo_app.ListaTowarow.window(title="Stan&y", class_name="TBitBtn").click()   

            left, top = 783, 180
            right = 860
            # Wysokość jednej komórki plus linia oddzielająca
            cell_height = 19
            border_height = 1
            total_cell_height = cell_height + border_height

            text_cena = None

            # Iteracja przez komórki, zaczynając od pierwszej
            for cell_index in range(49):  # Zakładając, że mamy maksymalnie 49 komórek
                # Obliczanie pozycji górnej i dolnej krawędzi komórki
                cell_top = top + cell_index * total_cell_height
                cell_bottom = cell_top + cell_height

                # Definiowanie obszaru zrzutu ekranu dla komórki
                bbox = (left, cell_top, right, cell_bottom)
                screenshot = ImageGrab.grab(bbox=bbox)
                enhancer = ImageEnhance.Contrast(screenshot)
                enhanced_screenshot = enhancer.enhance(2.0)

                # Użycie pytesseract do ekstrakcji tekstu z komórki
                text = pytesseract.image_to_string(enhanced_screenshot, lang='eng').strip()
                if text:
                    # Aktualizacja `cena_zakupu`, jeśli w komórce znajduje się tekst
                    text_cena = text
                    time.sleep(0.1)
                else:
                    # Przerywanie pętli, jeśli komórka jest pusta
                    break

            ergo_app.StanywgcenzakupuwmagazynieD.close()
            #zapis do pliku
            new_row = f"{edit_kod};{edit_nazwa};{edit_cena};{text_stan};{text_cena}\n"
            file.write(new_row)
            print(edit_cena, edit_nazwa, edit_kod, int(text_stan))


        towary_grid.send_keystrokes("{DOWN}")
        time.sleep(0.4)
