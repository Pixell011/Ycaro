from langchain import Assistant
import langchain

# Inicialize a assistente virtual
assistente = Assistant()

# Adicione automações e respostas a comandos de voz
assistente.add_command("qual é a previsão do tempo", "A previsão do tempo para hoje é ensolarado.")
assistente.add_command("toca uma música", "Aqui está uma música para você!")

# Defina uma função para lidar com comandos de voz
def handle_voice_command(command):
    response = assistente.execute(command)
    print("Assistente:", response)

# Exemplo de interação com o usuário
while True:
    user_input = input("Você: ")
    handle_voice_command(user_input)
