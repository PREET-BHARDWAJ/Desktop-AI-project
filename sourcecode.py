import tkinter as tk
from tkinter import scrolledtext
import webbrowser
import os
import datetime
import openai
import speech_recognition as sr
import win32com.client as w
from API import apikey

# Set up text-to-speech engine
speaker = w.Dispatch("SAPI.Spvoice")
openai.api_key = apikey

def speak(text):
    speaker.speak(text)

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        speak("Listening...")
        status_label.config(text="Listening...")
        root.update()
        audio = r.listen(source, 10, 3)
        try:
            status_label.config(text="Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            return query
        except Exception:
            status_label.config(text="Error recognizing speech.")
            return "Sorry, some error occurred."

def ai(prompt):
    response_text = f"AI Response to: {prompt}\n\n"
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    reply = response['choices'][0]['message']['content']
    response_text += reply

    if not os.path.exists("Openai"):
        os.mkdir("Openai")
    filename = prompt[:30].replace(" ", "_") + ".txt"
    with open(f"Openai/{filename}", "w", encoding="utf-8") as f:
        f.write(response_text)

    return reply

def handle_command():
    query = takecommand()
    output_text.insert(tk.END, f"\n🗣️You said: {query}\n\n")

    query_lower = query.lower()

    if "open youtube" in query_lower:
        speak("Opening YouTube")
        webbrowser.open("https://youtube.com")
    elif "open google" in query_lower:
        speak("Opening Google")
        webbrowser.open("https://google.com")
    elif "open wikipedia" in query_lower:
        speak("Opening Wikipedia")
        webbrowser.open("https://wikipedia.com")
    elif "play music" in query_lower:
        musicpath = "C:/Users/Dell/Music/magenta riddhim.mp3"
        os.startfile(musicpath)
    elif "the time" in query_lower:
        now = datetime.datetime.now()
        time_string = f"{now.strftime('%H')} hours and {now.strftime('%M')} minutes"
        speak(f"Sir, time is {time_string}")
        output_text.insert(tk.END, f"⏰ Time: {time_string}\n\n")
    elif "open code block" in query_lower:
        path = "C:/Program Files/CodeBlocks/codeblocks.exe"
        os.startfile(path)
    elif "open store" in query_lower:
        webbrowser.open("https://apps.microsoft.com/store")
    elif "using ai" in query_lower:
        answer = ai(query)
        speak("Your answer is ready.")
        output_text.insert(tk.END, f"🤖 AI Response:\n{answer}\n\n")
    elif "thanks" in query_lower or "exit" in query_lower:
        speak("You're welcome. Goodbye!")
        root.quit()
    else:
        speak("Sorry, I didn't understand that.")
        output_text.insert(tk.END, f"⚠️ Unknown command: {query}\n\n")

# ---- UI Setup ----
root = tk.Tk()
root.title("🎙️ Sam - Your Voice Assistant")
root.geometry("700x550")
root.configure(bg="#1e1e2f")
root.resizable(False, False)

# Fonts and colors
TITLE_FONT = ("Helvetica", 20, "bold")
BUTTON_FONT = ("Helvetica", 12, "bold")
TEXT_FONT = ("Consolas", 12)
ENTRY_BG = "#2b2b3c"
TEXT_COLOR = "#ffffff"
BUTTON_BG = "#4a90e2"
HOVER_BG = "#357ABD"
STATUS_COLOR = "#aaaaaa"

title_label = tk.Label(root, text="🎙️ Sam - AI Voice Assistant", font=TITLE_FONT, fg=TEXT_COLOR, bg="#1e1e2f")
title_label.pack(pady=15)
start_button = tk.Button(root, text="🎤 Speak", font=BUTTON_FONT,bg=BUTTON_BG, fg="white", activebackground=HOVER_BG,width=15, pady=5, bd=0, command=handle_command, cursor="hand2")
start_button.pack(pady=5)
output_text = scrolledtext.ScrolledText(root, font=TEXT_FONT, wrap=tk.WORD, height=17, width=80,bg=ENTRY_BG, fg=TEXT_COLOR, insertbackground=TEXT_COLOR,borderwidth=2, relief="groove")
output_text.pack(padx=15, pady=10)
status_label = tk.Label(root, text="Status: Idle", font=("Arial", 10),fg=STATUS_COLOR, bg="#1e1e2f")
status_label.pack(pady=5)
exit_button = tk.Button(root, text="❌ Exit", font=BUTTON_FONT, bg="#e74c3c", fg="white", activebackground="#c0392b",width=10, pady=5, bd=0, command=root.quit, cursor="hand2")
exit_button.pack(pady=10)

speak("Hello! I am Sam AI. What can I do for you today?")

root.mainloop()
