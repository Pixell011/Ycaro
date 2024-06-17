import openai
import time
import tkinter as tk
import threading
import pyautogui
import pyttsx3

# Defina sua chave de API
openai.api_key = 'sk-proj-hQyJHtNwfZUKhGFFrTibT3BlbkFJSNRNa1RQMHl9lOo1Fj7X'

# Variáveis globais para controle de cancelamento e pausa
cancel_automation = False
pause_automation = False

# Configuração do sintetizador de voz
engine = pyttsx3.init()
engine.setProperty('rate', 230)

def run_automation():
    global cancel_automation, pause_automation

    def check_pause():
        while pause_automation:
            time.sleep(0.1)
        if cancel_automation:
            return True
        return False

    try:
        engine.say("Iniciando automação Senhor Stark. Assim que finalizar, irei avisar. Á aproveita e toma um bom café! Tenha um excelente dia ")
        engine.runAndWait()
        print("Abrindo o Excel...")
        pyautogui.press("win")
        time.sleep(4)
        if check_pause(): return
        pyautogui.write("Excel")
        pyautogui.press("enter")
        time.sleep(5)
        if check_pause(): return

        print("Executando a automação no Excel...")
        # Adicione suas ações de automação aqui...

        print("Automação concluída.")
    except Exception as e:
        print(f"Erro na automação: {e}")

def gerar_resposta(prompt, max_retries=5):
    retries = 0
    while retries < max_retries:
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Você é um assistente útil."},
                    {"role": "user", "content": prompt}
                ]
            )
            return response.choices[0].message['content']
        except openai.error.RateLimitError:
            print(f"Limite de taxa excedido. Retentativa {retries + 1} de {max_retries} em 5 segundos...")
            retries += 1
            time.sleep(5)
        except Exception as e:
            print(f"Erro ao gerar resposta: {e}")
            return None
    return "Não foi possível obter a resposta da API após várias tentativas."

# Funções para controle de automação
def cancel():
    global cancel_automation
    cancel_automation = True
    root.quit()

def confirm():
    root.quit()

def pause():
    global pause_automation
    pause_automation = True

def resume():
    global pause_automation
    pause_automation = False

def start_automation():
    global cancel_automation
    cancel_automation = False
    automation_thread = threading.Thread(target=run_automation)
    automation_thread.start()

def action_1():
    print("Fornecedor2")

def action_2():
    print("E-mail")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Sexta-feira - Assistente de Automação")
root.geometry("400x200")

frame = tk.Frame(root)
frame.pack(pady=20)

label_voice = tk.Label(frame, text="Bem-vindo Senhor Stark")
label_voice.pack()

btn_start = tk.Button(frame, text="Iniciar Automação", command=start_automation)
btn_start.pack(side=tk.LEFT, padx=10)

btn_pause = tk.Button(frame, text="Pausar", command=pause)
btn_pause.pack(side=tk.LEFT, padx=10)

btn_resume = tk.Button(frame, text="Retomar", command=resume)
btn_resume.pack(side=tk.LEFT, padx=10)

btn_cancel = tk.Button(frame, text="Cancelar", command=cancel)
btn_cancel.pack(side=tk.LEFT, padx=10)

btn_confirm = tk.Button(frame, text="Confirmar", command=confirm)
btn_confirm.pack(side=tk.LEFT, padx=10)

action1_button = tk.Button(root, text="Ação 1", command=action_1, bg="#2196F3", fg="white", font=("Helvetica", 10))
action1_button.pack(pady=5)

action2_button = tk.Button(root, text="Ação 2", command=action_2, bg="#FF9800", fg="white", font=("Helvetica", 10))
action2_button.pack(pady=5)

root.mainloop()
