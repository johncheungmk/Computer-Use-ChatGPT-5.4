# Windows Computer Use Agent (GPT-5.4)

A Windows desktop automation app that uses **OpenAI GPT-5.4** with the **Responses API computer tool** to operate the UI like a human.

This project provides a simple **Tkinter GUI** where you can enter a task such as:

- `open edge browser and go to apple.com`
- `open Outlook and summarize unread emails`
- `open notepad and type a short note`

The app captures the screen, asks GPT-5.4 for the **next UI action**, executes **one action at a time**, captures a new screenshot, and repeats until the task is complete.

## Features

- OpenAI **GPT-5.4** + Responses API `computer` tool
- Windows UI automation using:
  - `pyautogui` for mouse/keyboard control
  - `mss` for screenshots
  - `Pillow` for image handling
- **One-action-at-a-time execution**
  - avoids blindly running a large batch of actions
  - re-checks the UI after every action
- **Final verification**
  - before reporting success, the app checks whether the goal is really complete
- **Optional minimize during run**
  - helps the model see the real desktop instead of the controller window
- **Optional one-time confirmation at start**
  - avoids step-by-step popup approvals that break focus
- **Status bar + progress indicator**
  - shows current stage such as:
    - Calling API
    - Executing action
    - Capturing screenshot
    - Final verification
- **Stall warning**
  - warns if the process appears stuck
- **Stop button**
  - requests a graceful stop
- **Emergency stop**
  - move the mouse to the **top-left corner** to trigger PyAutoGUI fail-safe

## How it works

1. You enter a task in the GUI.
2. The app sends the task to GPT-5.4 using the OpenAI Responses API with the `computer` tool.
3. GPT-5.4 returns the next UI action.
4. The app executes **only one action**.
5. The app captures a fresh screenshot.
6. GPT-5.4 re-plans based on the updated screen.
7. When the model appears done, the app performs a **final verification** screenshot before reporting success.

This design is more reliable than executing a whole batch at once, especially when window focus changes unexpectedly.

## Why this project exists

Simple UI automation often fails when the screen changes in unexpected ways.

Example:

Task:
`open edge browser and go to apple.com`

Possible problem:
- Edge is already open on `chatgpt.com`
- typing `apple.com` may go into the chat input box instead of the address bar

This project solves that by:

- checking the UI after **every action**
- adjusting the next action if the screen is not in the expected state
- verifying the final result before claiming success

## Requirements

- Windows 11
- Python 3.10+
- OpenAI API key with access to GPT-5.4 computer-use workflow
- A desktop session that is visible and interactive

## Install

Clone the repo and install dependencies:

```bash
pip install openai pyautogui mss pillow