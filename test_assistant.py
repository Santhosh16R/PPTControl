"""
Comprehensive Automated Test Suite for PowerPoint Voice & Chat Assistant.
"""

import os
import sys
from backend.command_parser import parse_command
from backend.ppt_controller import PPTController

def test_command_parsing():
    print("Testing Natural Language & Voice Command Parser...")
    test_cases = [
        ("next slide", "next"),
        ("next", "next"),
        ("please move to next slide", "next"),
        ("advance slide", "next"),
        ("previous slide", "prev"),
        ("go back", "prev"),
        ("back", "prev"),
        ("jump to slide 4", "goto"),
        ("go to slide number 2", "goto"),
        ("slide three", "goto"),
        ("first slide", "first"),
        ("last slide", "last"),
        ("open ppt", "open"),
        ("open the presentation", "open"),
        ("start slideshow", "start_show"),
        ("present", "start_show"),
        ("exit slideshow", "stop_show"),
        ("stop presentation", "stop_show"),
        ("black screen", "blank"),
        ("white screen", "blank"),
        ("resume presentation", "unblank"),
        ("close presentation", "close"),
        ("status", "status")
    ]

    passed = 0
    for phrase, expected_action in test_cases:
        res = parse_command(phrase)
        action = res.get("action")
        if action == expected_action:
            passed += 1
            print(f"  [PASS] '{phrase}' -> {action}")
        else:
            print(f"  [FAIL] '{phrase}' -> got {action}, expected {expected_action}")

    print(f"Command Parser Test Results: {passed}/{len(test_cases)} Passed.\n")
    assert passed == len(test_cases), "Some command parser tests failed"

def test_ppt_controller():
    print("Testing PowerPoint COM Controller...")
    controller = PPTController(default_folder=os.path.abspath("./presentations"))
    
    # 1. List presentations
    decks = controller.list_presentations()
    print(f"  Found {len(decks)} presentations in {controller.default_folder}")
    assert len(decks) > 0, "Expected at least 1 demo presentation"

    # 2. Open presentation & start slideshow
    open_res = controller.open_presentation(start_show=True)
    print(f"  Opened deck: {open_res}")
    assert open_res.get("status") == "opened"

    # 3. Check status
    status = controller.get_status()
    print(f"  Current status: slide {status.get('current_slide')}/{status.get('total_slides')}")
    assert status.get("running") is True

    # 4. Next slide
    next_res = controller.next_slide()
    print(f"  Next slide result: {next_res}")

    # 5. Jump to slide 3
    goto_res = controller.goto_slide(3)
    print(f"  Goto slide 3 result: {goto_res}")
    assert goto_res.get("current_slide") == 3

    # 6. Previous slide
    prev_res = controller.prev_slide()
    print(f"  Prev slide result: {prev_res}")
    assert prev_res.get("current_slide") == 2

    # 7. Stop slideshow
    stop_res = controller.stop_slideshow()
    print(f"  Stop slideshow result: {stop_res}")

    # 8. Close presentation
    close_res = controller.close_presentation()
    print(f"  Close presentation result: {close_res}")
    print("PowerPoint COM Controller tests passed successfully!\n")

if __name__ == "__main__":
    test_command_parsing()
    test_ppt_controller()
