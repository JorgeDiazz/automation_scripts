import pandas as pd
import pyautogui as pg
import keyboard as k
import pywhatkit
import time
from platform import system

def add_country_code_plus_sign(phone_number):
    if not phone_number.startswith('+'):
        phone_number = '+' + phone_number
    return phone_number

def create_invitation_message(row):
    return f'Hola {row["Nombre"]}. Por favor, ¿puedes confirmar si puedes asistir a nuestra sesión el día {row["Día"]} a las {row["Hora"]}?'

def send_invitations(invitations_dict):
    for phone_number, invitation in invitations_dict.items():
        if phone_number:
            phone_number_with_country_code = add_country_code_plus_sign(phone_number)
            
            print(f'Enviando mensaje a: Número de teléfono: {phone_number_with_country_code}, Mensaje: {invitation}')
            pywhatkit.sendwhatmsg_instantly(phone_number_with_country_code, invitation, wait_time=10)
            
            # Pressing enter key to double-check message is sent 
            time.sleep(2)
            pg.click()
            time.sleep(1)
            k.press_and_release('enter')

            close_tab()

def close_tab(wait_time: int = 2) -> None:
    time.sleep(wait_time)
    _system = system().lower()
    if _system in ("windows", "linux"):
        pg.hotkey("ctrl", "w")
    elif _system == "darwin":
        pg.hotkey("command", "w")
    else:
        raise Warning(f"{_system} not supported!")
    pg.press("enter")

def main():
    file_path = './Contactos.xlsx'
    df = pd.read_excel(file_path, dtype=str)

    # Drop unnecessary columns
    df = df.iloc[:, :4]

    # Drop rows with all NaN values
    df = df.dropna(how="all")

    # Create invitations dictionary
    invitations_dict = {row["Número de Teléfono"]: create_invitation_message(row) for _, row in df.iterrows()}

    # Send invitations via WhatsApp
    send_invitations(invitations_dict)

if __name__ == "__main__":
    main()