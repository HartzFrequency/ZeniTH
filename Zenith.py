import speech_recognition as speechR
import os
import win32com.client

# Initialize the Windows Speech API speaker
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# Function to make the system speak
def say(text):
    speaker.Speak(text)

# Function to listen to user's speech input
def hear():
    r = speechR.Recognizer()
    with speechR.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred"

if __name__ == '__main__':
    # Say a welcome message
    say("Welcome Zenith Here!")

    loop_Enteringf = 2  # loop entering frequency
    loop_EnteringC = True  # helps to get out of the loop once the frequency is over

    # Function to decrement the loop frequency and check if it's time to exit the loop
    def KeepLoopInCheck():
        global loop_Enteringf
        loop_Enteringf = loop_Enteringf - 1
        if loop_Enteringf == 0:
            global loop_EnteringC
            loop_EnteringC = False

    # Loop to listen to user input a certain number of times
    while loop_EnteringC:
        print("Yeah! I'm listening..")
        query = hear()
        say(query)
        KeepLoopInCheck()
