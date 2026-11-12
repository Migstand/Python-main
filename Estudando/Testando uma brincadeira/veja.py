import win32com.client
import random

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

        justificativa = input()
        if justificativa.lower() == "fazer depois":
            print("mentiroso")
            voz.Speak("Mentiroso")

        if justificativa.lower() == "esqueci":
            print("aiiinn esqueci")
            voz.Speak("aiinn esqueci")

        if justificativa.lower() == "ja ja pronto":
            print("Não tenho mais fé")
            voz.Speak("Não tenho mais fé")



