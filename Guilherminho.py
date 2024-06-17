import pyautogui
import threading
import time
import tkinter as tk
from PIL import Image, ImageTk

# Variável global para controle de cancelamento
cancel_automation = False

# Função para executar a automação
def run_automation(status_label, ok_button):
    global cancel_automation
    try:
        status_label.config(text="Abrindo o Excel...")
        pyautogui.press("win")
        time.sleep(3)
        if cancel_automation: return
        pyautogui.write("Excel")
        pyautogui.press("enter")
        time.sleep(10)  # Esperar o Excel abrir
        if cancel_automation: return

        status_label.config(text="Executando a automação no Excel...")
        pyautogui.click(x=487, y=55)  # Dados
        time.sleep(2)
        if cancel_automation: return
        pyautogui.click(x=29, y=111)  # Obter dados
        time.sleep(2)
        if cancel_automation: return
        pyautogui.click(x=237, y=153)  # De um arquivo
        time.sleep(2)
        if cancel_automation: return
        pyautogui.click(x=345, y=397)  # De uma pasta
        time.sleep(10)
        if cancel_automation: return
        pyautogui.click(x=89, y=184)  # Download
        time.sleep(2)
        if cancel_automation: return
        pyautogui.click(x=338, y=164)  # Pasta do dia
        time.sleep(2)
        if cancel_automation: return
        pyautogui.click(x=505, y=439)  # Abrir
        time.sleep(8)
        if cancel_automation: return
        pyautogui.click(x=744, y=646)  # Combinar
        time.sleep(3)
        if cancel_automation: return
        pyautogui.click(x=916, y=672)  # Transformar e combinar
        time.sleep(25)
        if cancel_automation: return
        pyautogui.click(x=414, y=241)  # Ex relatório
        time.sleep(2)
        if cancel_automation: return
        pyautogui.click(x=992, y=684)  # Ok
        time.sleep(25)
        if cancel_automation: return
        pyautogui.click(x=31, y=72)  # Fechar e carregar
        time.sleep(20)
        if cancel_automation: return
        pyautogui.hotkey('ctrl', 'g')
        pyautogui.write('$G:$G')
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'shift', '+')
        time.sleep(1)
        if cancel_automation: return
        pyautogui.click(x=791, y=239)
        time.sleep(1)
        if cancel_automation: return
        pyautogui.click(x=590, y=472)
        time.sleep(1)
        if cancel_automation: return
        pyautogui.click(x=592, y=507)
        time.sleep(1)
        if cancel_automation: return
        pyautogui.click(x=670, y=656)
        time.sleep(1)
        if cancel_automation: return
        pyautogui.click(x=25, y=19)
        time.sleep(1)
        if cancel_automation: return
        pyautogui.click(x=772, y=483)
        time.sleep(2)
        if cancel_automation: return
        pyautogui.hotkey('alt', 'tab')
        time.sleep(2)
        if cancel_automation: return
        
        pyautogui.hotkey('alt', 'tab')
        time.sleep(2)
        if cancel_automation: return
        
        pyautogui.doubleClick(x=467, y=257)
        time.sleep(3)
        if cancel_automation: return
        
        pyautogui.write("=PROCV(")
        time.sleep(1)
        if cancel_automation: return
        
        pyautogui.click(x=276, y=259)
        pyautogui.write(";")
        if cancel_automation: return
        
        pyautogui.hotkey('alt', 'tab')
        time.sleep(2)
        if cancel_automation: return
        
        pyautogui.click(x=881, y=221)
        time.sleep(1)
        if cancel_automation: return
        
        pyautogui.write(";1;0)")
        if cancel_automation: return
        
        pyautogui.hotkey('alt', 'tab')
        time.sleep(2)
        if cancel_automation: return
        
        pyautogui.press('enter')
        pyautogui.doubleClick(x=545,y=265)
        time.sleep(1)
        if cancel_automation: return
        
        pyautogui.hotkey('ctrl', 'c')
        pyautogui.hotkey('ctrl', 'shift', 'down')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(3)
        pyautogui.press("enter")
        if cancel_automation: return
        
        pyautogui.click(x=536, y=238)
        time.sleep(2)
        if cancel_automation: return
        
        pyautogui.click(x=337, y=472)
        time.sleep(2)
        if cancel_automation: return
        
        pyautogui.scroll(-300)
        time.sleep(2)
        if cancel_automation: return
        
        pyautogui.click(x=338, y=616)
        time.sleep(2)
        if cancel_automation: return
        
        pyautogui.click(x=415,y=654)
        time.sleep(3)
        if cancel_automation: return

        status_label.config(text="Automação concluída.")
        ok_button.config(state="normal")  # Ativar o botão OK
    except Exception as e:
        status_label.config(text=f"Erro na automação: {e}")
        ok_button.config(state="normal")  # Ativar o botão OK

# Função para cancelar a automação
def cancel():
    global cancel_automation
    cancel_automation = True
    root.quit()  # Fechar a janela de status

# Função para confirmar o fim da automação
def confirm():
    root.quit()  # Fechar a janela de status

# Função para mostrar a mensagem de status
def show_status_message():
    # Criar a janela
    global root
    root = tk.Tk()
    root.overrideredirect(True)  # Remover borda da janela
    root.attributes("-topmost", True)  # Manter a janela no topo

    # Definir o tamanho da janela
    window_width = 300
    window_height = 150

    # Calcular a posição para centralizar a janela na tela
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int(screen_height / 2 - window_height / 2)
    position_right = int(screen_width / 2 - window_width / 2)

    root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    
    # Frame para a label e botões
    frame = tk.Frame(root, bg="black", bd=3)
    frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    # Label para exibir a mensagem
    label = tk.Label(frame, text="Robô está trabalhando...", fg="white", bg="black", font=("Helvetica", 10))
    label.pack(padx=20, pady=10)

    # Frame para os botões
    button_frame = tk.Frame(frame, bg="black")
    button_frame.pack(pady=10)

    # Botão para cancelar a automação
    cancel_button = tk.Button(button_frame, text="Cancelar", command=cancel, bg="#d32f2f", fg="white", font=("Helvetica", 10))
    cancel_button.pack(side="left", padx=10)

    # Botão para confirmar o fim da automação
    ok_button = tk.Button(button_frame, text="OK", command=confirm, state="disabled", bg="#4CAF50", fg="white", font=("Helvetica", 10))
    ok_button.pack(side="right", padx=10)

    return root, label, ok_button

# Função principal
def main():
    root, status_label, ok_button = show_status_message()
    
    # Executar a automação em uma thread separada
    automation_thread = threading.Thread(target=run_automation, args=(status_label, ok_button))
    automation_thread.start()

    # Executar a janela de status
    root.mainloop()

if __name__ == "__main__":
    main()
