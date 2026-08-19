# PowerPoint AI Presenter (Node.js Edition)

Control Microsoft PowerPoint presentations on Windows using **Voice commands** or interactive **Chat messages** powered by Node.js.

---

## Quick Start

### Start the Node.js Server:
```bash
node server.js
```
The server will start at **`http://127.0.0.1:3000`** and automatically open your web browser.

---

## How It Works

1. **Presentations Folder**:
   - Place your `.pptx` or `.ppt` files inside the `presentations/` folder.
2. **Open PPT by Voice or Chat**:
   - Click the **Microphone** icon in your browser to enable continuous listening.
   - Say: **`"Open the PPT"`** or type **`open ppt`** in chat &rarr; Microsoft PowerPoint will launch on your screen.
3. **Control Slides**:
   - Say: **`"Next slide"`** &rarr; Advances slide in PowerPoint.
   - Say: **`"Previous slide"`** &rarr; Goes back.
   - Say: **`"Go to slide 3"`** &rarr; Jumps directly to slide 3.
   - Say: **`"First slide"`** / **`"Last slide"`** &rarr; Jumps to start or end.
   - Say: **`"Fullscreen"`** &rarr; Starts presentation mode.
   - Say: **`"Exit slideshow"`** &rarr; Exits presentation mode.
   - Say: **`"Black screen"`** &rarr; Blanks the screen.

---

## Project Structure

```
PPTControl/
├── server.js                 # Complete Node.js Backend Server
├── presentations/            # Put your PowerPoint files (.pptx / .ppt) here
├── static/
│   ├── index.html            # Web Assistant UI & Slide Stage
│   ├── style.css             # Glassmorphism dark-mode stylesheet
│   └── app.js                # Browser speech recognition & chat client
└── package.json
```
