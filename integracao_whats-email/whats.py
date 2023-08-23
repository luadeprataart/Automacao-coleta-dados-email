import pywhatkit # https://github.com/Ankit404butfound/PyWhatKit/tree/master
import pyperclip
from pywhatkit.core import exceptions
from re import fullmatch
import win32clipboard

def _send_msg_whats( to_ , msg='Oi'):
    pyperclip.copy(msg)
    #manda a mensagem para o contato dado como parametro 
    pywhatkit.sendwhats_image(to_, "./docs/img/img.jpeg", pyperclip.paste(), tab_close = True)


if __name__ == "__main__":
    _send_msg_whats(to_)




