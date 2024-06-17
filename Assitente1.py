import tkinter as tk
import threading
import pyautogui
import time
import pyttsx3

# Variável global para controle de cancelamento
cancel_automation = False

# Configuração do sintetizador de voz
engine = pyttsx3.init()
engine.setProperty('rate', 230)

# Função para executar a automação
def run_automation():
    global cancel_automation
    try:
        engine.say("Iniciando automação Senhor Stark.    Aproveita e toma um bom café!  Tenha um excelente dia ")
        engine.runAndWait()
        print("Abrindo o Excel...")
        pyautogui.press("win")
        time.sleep(4)
        if cancel_automation: return
        pyautogui.write("Excel")
        pyautogui.press("enter")
        time.sleep(5)  # Esperar o Excel abrir
        if cancel_automation: return

        print("Executando a automação no Excel...")
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
        time.sleep(7)
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
        time.sleep(4)
        if cancel_automation: return
        pyautogui.click(x=916, y=672)  # Transformar e combinar
        time.sleep(20)
        if cancel_automation: return
        pyautogui.click(x=414, y=241)  # Ex relatório
        time.sleep(2)
        if cancel_automation: return
        pyautogui.click(x=992, y=684)  # Ok
        time.sleep(25)
        if cancel_automation: return
        pyautogui.click(x=31, y=72)  # Fechar e carregar
        time.sleep(17)
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

        print("Automação concluída.")
    except Exception as e:
        print(f"Erro na automação: {e}")

# Função para cancelar a automação
def cancel():
    global cancel_automation
    cancel_automation = True
    root.quit()  # Fechar a janela de status

# Função para confirmar o fim da automação
def confirm():
    root.quit()  # Fechar a janela de status

# Função para iniciar a automação
def start_automation():
    global cancel_automation
    cancel_automation = False
    automation_thread = threading.Thread(target=run_automation)
    automation_thread.start()

# Função para ação 1
def action_1():
    print("Fornecedor2")

# Função para ação 2
def action_2():
    print("E-mail")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Sexta-feira - Assistente de Automação")
root.geometry("400x200")

# Frame para a interface gráfica
frame = tk.Frame(root)
frame.pack(pady=20)

# Label para mensagem de voz
label_voice = tk.Label(frame, text="Bem-vindo Senhor Stark")
label_voice.pack()

# Botões de ação
btn_start = tk.Button(frame, text="Iniciar Automação", command=start_automation)
btn_start.pack(side=tk.LEFT, padx=10)

btn_cancel = tk.Button(frame, text="Cancelar", command=cancel)
btn_cancel.pack(side=tk.LEFT, padx=10)

btn_confirm = tk.Button(frame, text="Confirmar", command=confirm)
btn_confirm.pack(side=tk.LEFT, padx=10)

# Botões para outras ações
action1_button = tk.Button(root, text="Ação 1", command=action_1, bg="#2196F3", fg="white", font=("Helvetica", 10))
action1_button.pack(pady=5)

action2_button = tk.Button(root, text="Ação 2", command=action_2, bg="#FF9800", fg="white", font=("Helvetica", 10))
action2_button.pack(pady=5)

root.mainloop()
