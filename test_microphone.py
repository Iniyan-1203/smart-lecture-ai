
import speech_recognition as sr

recognizer = sr.Recognizer()

print("Available microphones:")

for index, name in enumerate(sr.Microphone.list_microphone_names()):
    print(f"{index}: {name}")

print("\nTesting microphone...")

try:
    with sr.Microphone() as source:
        print("Adjusting for background noise...")
        recognizer.adjust_for_ambient_noise(source, duration=2)

        print("Speak something now!")
        audio = recognizer.listen(source, timeout=10, phrase_time_limit=10)

    print("\nConverting speech to text...")

    text = recognizer.recognize_google(audio, language="en-IN")

    print("\nRecognized speech:")
    print(text)

except sr.WaitTimeoutError:
    print("No speech detected before the timeout.")

except sr.UnknownValueError:
    print("Speech was recorded, but the words could not be understood.")

except sr.RequestError as error:
    print(f"Speech recognition service error: {error}")

except Exception as error:
    print(f"Error: {error}")
