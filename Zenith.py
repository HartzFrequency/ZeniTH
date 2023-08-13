import speech_recognition as speechR 
import os
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    speaker.Speak(text)
    # os.system(f"say {text}")

def hear():
    r = speechR.Recognizer()
    with speechR.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio,language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Ocuured"            


if __name__ == '__main__':
    say("Welcome Zenith Here!") #What to say!? 
    
    loop_Enteringf = 2 #loop enetring frequency
    loop_EnteringC = True #helps to get out of the loop once the frequency overs
    def KeepLoopInCheck():
        global loop_Enteringf
        loop_Enteringf = loop_Enteringf - 1
        if loop_Enteringf == 0:
            global loop_EnteringC
            loop_EnteringC = False
    
    while loop_EnteringC:
        print("yeah! I'm listening..")
        query = hear()
        say(query)
        KeepLoopInCheck()