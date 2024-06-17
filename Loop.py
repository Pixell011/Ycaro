import pyautogui
import time

def right_click():
    # Simula um clique com o botão direito do mouse
    pyautogui.click(button='right')

def refresh_screen():
    # Pressiona a tecla F5 para atualizar a tela
    pyautogui.press('f5')

def double_click_desktop():
    # Move o mouse para uma posição na área de trabalho e faz um duplo clique
    # Você pode precisar ajustar as coordenadas (x, y) conforme necessário
    desktop_position = (563, 751)  # Coordenadas de exemplo
    pyautogui.moveTo(desktop_position)
    pyautogui.doubleClick()

while True:
    right_click()
    time.sleep(1)  # Espera 1 segundo entre as ações
    refresh_screen()
    time.sleep(1)  # Espera 1 segundo entre as ações
    double_click_desktop()
    time.sleep(5)  # Espera 5 segundos antes de repetir o loop
