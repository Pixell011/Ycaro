import pyautogui
import threading
import time
import tkinter as tk
from tkinter import messagebox

# Função para executar a automação
def run_automation():
    try:
        # Simulação do trabalho do robô
        pyautogui.hotkey('win', 'd')  # Minimizar todas as janelas
        time.sleep(2)

        pyautogui.press("win")
        time.sleep(3)
        pyautogui.write("Excel")
        pyautogui.press("enter")
        time.sleep(5)  # Esperar o Excel abrir
        pyautogui.click(x=271, y=217)  # Clicar em novo
        time.sleep(5)
        pyautogui.click(x=487, y=55)  # Dados
        time.sleep(2)
        pyautogui.click(x=29, y=111)  # Obter dados
        time.sleep(2)
        pyautogui.click(x=237, y=153)  # De um arquivo
        time.sleep(2)
        pyautogui.click(x=345, y=397)  # De uma pasta
        time.sleep(10)
        pyautogui.click(x=89, y=184)  # Download
        time.sleep(2)
        pyautogui.click(x=338, y=164)  # Pasta do dia
        time.sleep(2)
        pyautogui.click(x=505, y=439)  # Abrir
        time.sleep(8)
        pyautogui.click(x=744, y=646)  # Combinar
        time.sleep(3)
        pyautogui.click(x=916, y=672)  # Transformar e combinar
        time.sleep(15)
        pyautogui.click(x=414, y=241)  # Ex relatório
        time.sleep(2)
        pyautogui.click(x=992, y=684)  # Ok
        time.sleep(25)
        pyautogui.click(x=31, y=72)  # Fechar e carregar
        time.sleep(20)
        pyautogui.hotkey('ctrl', 'g')
        pyautogui.write('$G:$G')
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'shift', '+')
        time.sleep(1)
        pyautogui.click(x=791, y=239)
        time.sleep(1)
        pyautogui.click(x=590, y=472)
        time.sleep(1)
        pyautogui.click(x=592, y=507)
        time.sleep(1)
        pyautogui.click(x=670, y=656)
        time.sleep(1)
        pyautogui.click(x=25, y=19)
        time.sleep(1)
        pyautogui.click(x=772, y=483)
        # Automação concluída
        return "Automação concluída."
    except Exception as e:
        return f"Erro na automação: {e}"

# Função principal
def main():
    # Criar a janela de status
    root = tk.Tk()
    root.attributes('-fullscreen', True)  # Tela cheia
    root.config(bg="black")

    # Manter a janela sempre no topo
    root.attributes('-topmost', True)

    # Tornar a janela transparente para cliques do mouse
    root.wm_attributes('-transparentcolor', 'black')
    root.wm_attributes('-alpha', 0.8)

    # Frame para a label
    frame = tk.Frame(root, bg="black")
    frame.pack(expand=True)

    # Label para exibir a mensagem
    label = tk.Label(frame, text="O robô está trabalhando... Por favor, aguarde.", fg="white", bg="black", font=("Helvetica", 24))
    label.pack()

    # Função para executar a automação e mostrar o resultado na GUI principal
    def automation_task():
        result = run_automation()
        root.after(0, lambda: (messagebox.showinfo("Status da Automação", result), root.destroy()))

    # Mostrar a janela de status imediatamente
    root.update()

    # Executar a automação em uma thread separada
    automation_thread = threading.Thread(target=automation_task)
    automation_thread.start()

    # Executar o loop principal do Tkinter
    root.mainloop()

if __name__ == "__main__":
    main()
