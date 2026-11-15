import speech_recognition as sr

recognizer = sr.Recognizer()

print("Available microphones:")

for index, name in enumerate(sr.Microphone.list_microphone_names()):
    print(f"{index}: {name}")

mic_index = None

print("\nEnter the microphone number you want to test.")
print("Choose the microphone corresponding to your Airdew 112 earbuds.")

try:
    mic_index = int(input("Microphone number: "))

    with sr.Microphone(device_index=mic_index) as source:
        print("\nAdjusting for background noise... Please wait.")
        recognizer.adjust_for_ambient_noise(source, duration=2)

        print("Speak clearly for a few seconds!")

        audio = recognizer.listen(
            source,
            timeout=10,
            phrase_time_limit=15
        )

        print("\nConverting speech to text...")

        text = recognizer.recognize_google(audio)

        print("\nRecognized text:")
        print(text)

except sr.WaitTimeoutError:
    print("No speech detected. Check your microphone.")

except sr.UnknownValueError:
    print("Audio was captured, but speech was not understood.")

except sr.RequestError as error:
    print("Google speech recognition error:", error)

except (ValueError, IndexError):
    print("Invalid microphone number.")

except OSError as error:
    print("Microphone error:", error)