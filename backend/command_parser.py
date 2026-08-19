"""
Natural Language Voice and Chat Command Parser for PowerPoint Control.
Parses conversational user inputs into concrete actions and parameters.
"""

import re

WORD_TO_NUMBER = {
    "one": 1, "first": 1, "1st": 1,
    "two": 2, "second": 2, "2nd": 2,
    "three": 3, "third": 3, "3rd": 3,
    "four": 4, "fourth": 4, "4th": 4,
    "five": 5, "fifth": 5, "5th": 5,
    "six": 6, "sixth": 6, "6th": 6,
    "seven": 7, "seventh": 7, "7th": 7,
    "eight": 8, "eighth": 8, "8th": 8,
    "nine": 9, "ninth": 9, "9th": 9,
    "ten": 10, "tenth": 10, "10th": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
    "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20
}

def parse_command(raw_text: str) -> dict:
    """
    Parse natural text / speech transcripts into structured actions.
    """
    if not raw_text or not raw_text.strip():
        return {"action": "unknown", "feedback": "I didn't catch that. Please speak or type a command."}

    text = raw_text.lower().strip()
    text = re.sub(r"[^\w\s\.]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    # 1. Open Presentation (Check early for open/launch/load commands)
    # Examples: "open", "open ppt", "open the ppt", "open presentation", "launch ppt", "load ppt", "open sample"
    if re.search(r"\b(open|load|launch)\b", text) and not re.search(r"\b(slide|close|exit|stop)\b", text):
        open_match = re.search(r"\b(?:open|load|launch)\s+(?:the\s+)?(?:ppt|powerpoint|presentation|deck|file)?\s*(.*)", text)
        target_spec = None
        if open_match:
            spec = open_match.group(1).strip()
            spec = re.sub(r"\b(please|now|file|presentation|ppt|powerpoint|deck)\b", "", spec).strip()
            if spec:
                target_spec = spec
        return {
            "action": "open",
            "params": {"target": target_spec},
            "feedback": f"Opening presentation {target_spec or ''}..."
        }

    # 2. Next Slide
    if re.search(r"\b(next|advance|forward|move on|next page|next slide)\b", text):
        return {
            "action": "next",
            "params": {},
            "feedback": "Moving to next slide."
        }

    # 3. Previous Slide / Back
    if re.search(r"\b(previous|prev|back|go back|backwards|prior|return)\b", text) and not re.search(r"\blast slide\b", text):
        return {
            "action": "prev",
            "params": {},
            "feedback": "Going back to previous slide."
        }

    # 4. First Slide / Beginning
    if re.search(r"\b(first slide|beginning|start from beginning|go to start|jump to start|first page)\b", text):
        return {
            "action": "first",
            "params": {"slide": 1},
            "feedback": "Navigating to the first slide."
        }

    # 5. Last Slide / End
    if re.search(r"\b(last slide|end of presentation|final slide|jump to end|go to end)\b", text):
        return {
            "action": "last",
            "params": {},
            "feedback": "Navigating to the final slide."
        }

    # 6. Go To / Jump To specific slide number
    goto_match = re.search(r"\b(?:go to|jump to|switch to|move to|open|show)?\s*slide\s*(?:number|no\.?)?\s*(\w+)\b", text)
    if goto_match:
        target_str = goto_match.group(1).lower()
        num = None
        if target_str.isdigit():
            num = int(target_str)
        elif target_str in WORD_TO_NUMBER:
            num = WORD_TO_NUMBER[target_str]
        
        if num is not None:
            return {
                "action": "goto",
                "params": {"slide": num},
                "feedback": f"Jumping to slide {num}."
            }

    direct_slide = re.search(r"\bslide\s+(\d+)\b", text)
    if direct_slide:
        num = int(direct_slide.group(1))
        return {
            "action": "goto",
            "params": {"slide": num},
            "feedback": f"Jumping to slide {num}."
        }

    # 7. Start / Present Slideshow
    if re.search(r"\b(start presentation|start slideshow|present|presentation mode|full screen|fullscreen|play slideshow|start show)\b", text):
        return {
            "action": "start_show",
            "params": {},
            "feedback": "Starting slideshow presentation."
        }

    # 8. Stop / Exit Slideshow
    if re.search(r"\b(stop slideshow|exit slideshow|end slideshow|quit slideshow|stop presentation|exit presentation|end presentation)\b", text):
        return {
            "action": "stop_show",
            "params": {},
            "feedback": "Exiting slideshow mode."
        }

    # 9. Close PowerPoint / Close Deck
    if re.search(r"\b(close ppt|close presentation|close powerpoint|exit powerpoint|close deck|close)\b", text):
        return {
            "action": "close",
            "params": {},
            "feedback": "Closing the presentation."
        }

    # 10. Blank Screen / Blackout / Whiteout
    if re.search(r"\b(black screen|black out|blackout|blank screen|black)\b", text):
        return {
            "action": "blank",
            "params": {"color": "black"},
            "feedback": "Blacking out screen."
        }
    if re.search(r"\b(white screen|white out|whiteout|white)\b", text):
        return {
            "action": "blank",
            "params": {"color": "white"},
            "feedback": "White screen mode."
        }
    if re.search(r"\b(unblank|resume presentation|show screen|clear blackout)\b", text):
        return {
            "action": "unblank",
            "params": {"color": "unblank"},
            "feedback": "Resuming screen display."
        }

    # 11. Status / Info
    if re.search(r"\b(status|current slide|what slide|which slide|info|where are we)\b", text):
        return {
            "action": "status",
            "params": {},
            "feedback": "Checking presentation status."
        }

    return {
        "action": "unknown",
        "raw_text": raw_text,
        "feedback": f"Command not recognized: \"{raw_text}\". Try saying 'next slide', 'previous slide', or 'open the ppt'."
    }
