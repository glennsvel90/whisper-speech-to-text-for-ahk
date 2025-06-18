# Whisper AHK Transcriber

![AutoHotkey](https://img.shields.io/badge/AutoHotkey-v1.1-green.svg)

A simple yet powerful AutoHotkey (v1.1) script for Windows that allows you to instantly transcribe your voice to text using OpenAI's Whisper API. Press a hotkey to start recording, press it again to stop, and have the transcribed text automatically pasted at your cursor and copied to your clipboard.

## Features

-   **Hotkey Activated:** Simply press `Q` + `Z` to start and stop recording. No need to open any windows.
-   **Seamless Integration:** Transcribed text is automatically typed out at your current cursor location, whether you're in a text editor, a browser, or any other application.
-   **Clipboard Access:** The transcribed text is also sent to your clipboard for easy pasting elsewhere.
-   **Visual Indicator:** A small, unobtrusive "REC" window appears on screen so you always know when the script is recording.
-   **System Notifications:** Uses standard Windows tray notifications to keep you informed of the script's status (e.g., "Recording Started," "Transcription Complete," "API Error").
-   **No External Dependencies:** Besides AutoHotkey itself, the script uses built-in Windows components (`MCI`, `ADODB.Stream`, `WinHttpRequest`) and does not require you to install anything else.

## Prerequisites

Before you can use this script, you will need:

1.  **[AutoHotkey v1.1](https://www.autohotkey.com/download/v1.1/)**: This script is specifically written for AutoHotkey v1.1 and is **not** compatible with v2.0.
2.  **An OpenAI API Key**: You need an account with OpenAI and a valid API key. You must also have a payment method set up on your account for API usage. You can get your key from the [OpenAI API keys page](https://platform.openai.com/account/api-keys).

## Installation

1.  **Install AutoHotkey:** If you don't have it already, download and install **AutoHotkey v1.1** from the official website.
2.  **Download the Script:** Download the `Whisper_Transcribe.ahk` file from this repository.
3.  **Save the Script:** Place the `.ahk` file anywhere on your computer, for example, on your Desktop or in your Documents folder.

## Configuration

You must add your OpenAI API key to the script before it will work.

1.  **Open the script file:** Right-click on `Whisper_Transcribe.ahk` and choose **Edit Script** (or open it with Notepad).
2.  **Find the configuration line:** Near the top of the script, you will see the following line:
    ```autohotkey
    Global YOUR_OPENAI_API_KEY := "PASTE_YOUR_OPENAI_API_KEY_HERE"
    ```
3.  **Paste your API key:** Replace `"PASTE_YOUR_OPENAI_API_KEY_HERE"` with your actual OpenAI secret key. Make sure your key remains inside the quotation marks.
    * **Example:** `Global YOUR_OPENAI_API_KEY := "sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"`
4.  **Save** and close the file.

## Usage

1.  **Run the script:** Double-click the `Whisper_Transcribe.ahk` file. You should see a new green "H" icon appear in your Windows system tray (near the clock).
2.  **Start Recording:** Click into any text field and press the **`Q`** and **`Z`** keys at the same time. A red "REC" indicator will appear in the bottom-right of your screen.
3.  **Speak:** Say whatever you want to transcribe.
4.  **Stop Recording:** Press **`Q`** and **`Z`** again. The "REC" indicator will disappear.
5.  **Get Text:** After a brief moment, the script will type out the transcribed text and copy it to your clipboard.
6.  **Get Text:** If you would like to type a capital letter Q, first press and hold down the Shift key, then press Q while shift key is still held down. 

To close the script completely, right-click the "H" icon in the system tray and select **Exit**.


## Contributing

Contributions are welcome! If you have ideas for improvements or have found a bug, please feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License. 
