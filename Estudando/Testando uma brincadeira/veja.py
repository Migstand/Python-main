import win32com.client
import random

import threading

# Variável global para armazenar a resposta do usuário
resultado_input = None

def pedir_input():
    global resultado_input
    resultado_input = input()

def input_com_timeout(tempo_limite, voz):
    global resultado_input
    resultado_input = None
    
    # Cria uma thread para a função de input
    t = threading.Thread(target=pedir_input)
    t.daemon = True  # Permite que a thread encerre se o programa fechar
    t.start()
    
    # Espera o tempo limite na thread principal
    t.join(timeout=tempo_limite)
    
    if t.is_alive():
        print("Deixa de enrolar")
        voz.Speak("Deixa de enrolar")
        return None
    else:
        return resultado_input

# Exemplo de uso (com 5 segundos de limite)


voz = win32com.client.Dispatch("SAPI.SpVoice")

frases = [
    "Cadê o trailer",
    "Já passou da hora",
    "Lança logo isso",
]

print("inclusive me lembrei de uma coisa")
pergunta = str(input())

if pergunta:
    while True:
        frase = random.choice(frases)
        print(frase)
        voz.Speak(frase)

        justificativa = input_com_timeout(5, voz)

        if justificativa is not None:
            #print(f"Você digitou: {justificativa}")

            if justificativa.lower() == "fazer depois":
                print("mentiroso")
                voz.Speak("Mentiroso")

            if justificativa.lower() == "esqueci":
                print("aiiinn esqueci")
                voz.Speak("aiinn esqueci")

            if justificativa.lower() == "ja ja pronto":
                print("Não tenho mais fé")
                voz.Speak("Não tenho mais fé")



