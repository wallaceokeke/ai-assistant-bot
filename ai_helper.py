import os
import platform
import datetime

def main():
    print("🔐 Welcome to your Personal AI Assistant")
    print(f"🖥️  OS Detected: {platform.system()} {platform.release()}")
    print("⌛", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    print("Type 'help' to see commands. Type 'exit' to quit.")

    while True:
        cmd = input(">> ").strip().lower()

        if cmd == 'exit':
            print("👋 Goodbye, Master!")
            break
        elif cmd == 'help':
            print("""
Available commands:
 - open folder     → Opens C:\\AI_Assistant
 - save note       → Saves your text to secret_notes.txt
 - show notes      → Displays your stored notes
 - system info     → Shows system details
 - clear           → Clears the screen
""")
        elif cmd == 'open folder':
            os.system('start C:\\AI_Assistant')
        elif cmd == 'save note':
            note = input("Type your note: ")
            with open("secret_notes.txt", "a") as f:
                f.write(note + "\n")
            print("📝 Note saved.")
        elif cmd == 'show notes':
            if os.path.exists("secret_notes.txt"):
                with open("secret_notes.txt", "r") as f:
                    print(f.read())
            else:
                print("No notes found.")
        elif cmd == 'system info':
            os.system("systeminfo")
        elif cmd == 'clear':
            os.system('cls' if os.name == 'nt' else 'clear')
        else:
            print("❌ Unknown command. Type 'help' to see options.")

if __name__ == '__main__':
    main()
