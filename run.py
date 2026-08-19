"""
PowerPoint Voice & Chat Assistant Launcher.
Starts the FastAPI server and opens the Assistant UI in your default web browser.
"""

import os
import sys
import webbrowser
import threading
import time
import uvicorn

def open_browser():
    time.sleep(1.2)
    webbrowser.open("http://127.0.0.1:8000")

def main():
    print("=" * 65)
    print("  PowerPoint AI Presenter (Voice & Chat Assistant)")
    print("=" * 65)
    
    # Ensure presentations folder exists
    pres_folder = os.path.abspath("./presentations")
    os.makedirs(pres_folder, exist_ok=True)
    
    # Check if sample presentation exists, if not generate it
    sample_path = os.path.join(pres_folder, "Sample_AI_Presentation.pptx")
    if not os.path.exists(sample_path):
        print("Generating sample demo presentation...")
        try:
            from create_sample_deck import create_sample_presentation
            create_sample_presentation(sample_path)
        except Exception as e:
            print(f"Sample generation note: {e}")

    print(f"[*] Presentations folder: {pres_folder}")
    print("[*] Starting Assistant Server at http://127.0.0.1:8000 ...")
    print("[*] Opening Web Assistant in your default browser...")
    print("=" * 65)

    threading.Thread(target=open_browser, daemon=True).start()
    
    uvicorn.run("backend.app:app", host="127.0.0.1", port=8000, reload=False)

if __name__ == "__main__":
    main()
