"""
FastAPI Backend Server for PowerPoint Voice and Chat Assistant.
"""

import os
import asyncio
from typing import Optional
from fastapi import FastAPI, WebSocket, WebSocketDisconnect, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

from backend.ppt_controller import PPTController
from backend.command_parser import parse_command

app = FastAPI(title="PowerPoint Voice & Chat Assistant API")

# Initialize controller
controller = PPTController(default_folder=os.path.abspath("./presentations"))

# Active WebSocket connections
connected_clients = set()


class CommandRequest(BaseModel):
    text: str
    source: Optional[str] = "chat"  # "voice" or "chat"


class ActionRequest(BaseModel):
    action: str
    params: Optional[dict] = None


class FolderRequest(BaseModel):
    folder_path: str


async def broadcast_status():
    """Broadcast current PPT status to all connected web clients."""
    if not connected_clients:
        return
    status = controller.get_status()
    payload = {"type": "status_update", "data": status}
    disconnected = set()
    for ws in connected_clients:
        try:
            await ws.send_json(payload)
        except Exception:
            disconnected.add(ws)
    for ws in disconnected:
        connected_clients.remove(ws)


@app.get("/api/presentations")
async def get_presentations(folder: Optional[str] = None):
    """List PowerPoint presentations in the designated folder."""
    try:
        files = controller.list_presentations(folder)
        return {
            "folder": folder or controller.default_folder,
            "count": len(files),
            "presentations": files
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/folder")
async def set_folder(req: FolderRequest):
    """Set default presentations folder."""
    if not os.path.exists(req.folder_path):
        raise HTTPException(status_code=400, detail=f"Folder '{req.folder_path}' does not exist.")
    controller.default_folder = os.path.abspath(req.folder_path)
    files = controller.list_presentations()
    await broadcast_status()
    return {
        "status": "folder_updated",
        "folder": controller.default_folder,
        "files_found": len(files)
    }


@app.get("/api/status")
async def get_status():
    """Get current presentation and slide status."""
    return controller.get_status()


@app.get("/api/deck/slides")
async def get_deck_slides(path: Optional[str] = None):
    """Get extracted slides for web presentation stage."""
    slides = controller.parse_deck_slides(path)
    return {
        "presentation": os.path.basename(controller.active_presentation_path or "Presentation"),
        "total_slides": len(slides),
        "slides": slides
    }



@app.post("/api/action")
async def execute_action(req: ActionRequest):
    """Execute direct control actions."""
    action = req.action.lower()
    params = req.params or {}
    result = {}

    try:
        if action == "open":
            target = params.get("target") or params.get("file")
            result = controller.open_presentation(target, start_show=params.get("start_show", True))
        elif action == "start_show":
            result = controller.start_slideshow()
        elif action == "stop_show":
            result = controller.stop_slideshow()
        elif action == "next":
            result = controller.next_slide()
        elif action == "prev":
            result = controller.prev_slide()
        elif action == "goto":
            slide = params.get("slide", 1)
            result = controller.goto_slide(slide)
        elif action == "first":
            result = controller.first_slide()
        elif action == "last":
            result = controller.last_slide()
        elif action == "blank":
            color = params.get("color", "black")
            result = controller.blank_screen(color)
        elif action == "unblank":
            result = controller.blank_screen("unblank")
        elif action == "close":
            result = controller.close_presentation()
        elif action == "status":
            result = controller.get_status()
        else:
            raise HTTPException(status_code=400, detail=f"Unknown action '{action}'")

        await broadcast_status()
        return {"success": True, "action": action, "result": result, "status": controller.get_status()}
    except Exception as e:
        return JSONResponse(status_code=500, content={"success": False, "error": str(e)})


@app.post("/api/command")
async def handle_command(req: CommandRequest):
    """Process natural language voice or chat command."""
    parsed = parse_command(req.text)
    action = parsed.get("action")
    params = parsed.get("params", {})
    feedback = parsed.get("feedback", "")
    result = None

    if action == "unknown":
        return {
            "success": False,
            "parsed": parsed,
            "feedback": feedback,
            "status": controller.get_status()
        }

    # Execute matched action
    try:
        if action == "open":
            target = params.get("target")
            result = controller.open_presentation(target, start_show=True)
            if "file" in result:
                feedback = f"Opened presentation: {result['file']}"
        elif action == "start_show":
            result = controller.start_slideshow()
        elif action == "stop_show":
            result = controller.stop_slideshow()
        elif action == "next":
            result = controller.next_slide()
            if result.get("current_slide"):
                feedback = f"Slide {result['current_slide']} of {result.get('total_slides', '?')}"
        elif action == "prev":
            result = controller.prev_slide()
            if result.get("current_slide"):
                feedback = f"Slide {result['current_slide']} of {result.get('total_slides', '?')}"
        elif action == "goto":
            slide = params.get("slide", 1)
            result = controller.goto_slide(slide)
            if result.get("current_slide"):
                feedback = f"Jumped to slide {result['current_slide']}"
        elif action == "first":
            result = controller.first_slide()
            feedback = "At the first slide."
        elif action == "last":
            result = controller.last_slide()
            feedback = f"At the last slide ({result.get('current_slide', '')})."
        elif action == "blank":
            result = controller.blank_screen(params.get("color", "black"))
        elif action == "unblank":
            result = controller.blank_screen("unblank")
        elif action == "close":
            result = controller.close_presentation()
            feedback = "Closed presentation."
        elif action == "status":
            result = controller.get_status()
            if result.get("presentation_name"):
                feedback = f"Currently on slide {result.get('current_slide')} of {result.get('total_slides')} in {result.get('presentation_name')}."
            else:
                feedback = "No presentation currently active."

        await broadcast_status()
        return {
            "success": True,
            "parsed": parsed,
            "feedback": feedback,
            "result": result,
            "status": controller.get_status()
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={
            "success": False,
            "parsed": parsed,
            "feedback": f"Error executing command: {e}",
            "error": str(e)
        })


@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    """Real-time status broadcast websocket."""
    await websocket.accept()
    connected_clients.add(websocket)
    # Send immediate initial status
    try:
        await websocket.send_json({"type": "status_update", "data": controller.get_status()})
        while True:
            data = await websocket.receive_text()
            # If client sends ping, respond with status
            await websocket.send_json({"type": "status_update", "data": controller.get_status()})
    except WebSocketDisconnect:
        connected_clients.discard(websocket)
    except Exception:
        connected_clients.discard(websocket)


# Mount static directory for modern UI
static_dir = os.path.abspath("./static")
os.makedirs(static_dir, exist_ok=True)
app.mount("/static", StaticFiles(directory=static_dir), name="static")


@app.get("/")
async def root():
    index_file = os.path.join(static_dir, "index.html")
    if os.path.exists(index_file):
        return FileResponse(index_file)
    return {"message": "PowerPoint Assistant API running. Frontend static/index.html not found."}
