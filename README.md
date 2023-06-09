
# winGPT (under development)
**winGPT** allows you to directly use chatGPT in the editors of Outlook, Notepad, MS Word, web browsers etc. on Windows. The idea behind winGPT was inspired by macGPT for Mac users.  


**Usage:**  
In Outlook, Notepad, MS Word, or other text editors,  
Type **/plz** to start a conversation, press Enter key twice (x2) or Shift+Enter for chatGPT to respond.

**Shortcuts:**  
Define your prompts as shortcuts in `shortcuts.json` file, and they will be converted to their full version before being sent to ChatGPT.  
Examples:   
* Use `/plz.sum` to summerize text.  
* Use `/plz.revise` to revise text.  
* Use `[[p]]` or `[[p` as substitutes for clipboard content.  
* `/plz.sum [[p]]` will create a summary of the clipboard's content.  
See shortcuts.json for more examples, and you can add your own shortcuts there.  

**Config file:**  
You can change the default settings such as the **trigger word** in **config.json** file.
See config.json for more details.

## Requirements
- OpenAI API key. You need to first get an OpenAI API key from https://platform.openai.com/account/api-keys
- python3

## Installation  
To install winGPT, simply clone this GitHub repository to your local machine and follow the instructions in the README file.
```
pip install -r requirements.txt
```

To create a standalone .exe file for wingpt.py, PyInstaller can be used. Follow these steps:
1. Create a virtual environment:
```
python -m venv wingpt_venv
```
2. Activate the virtual environment:
```
wingpt_venv\Scripts\activate
```
3. Install the required packages:
```
pip install -r requirements.txt
pip install pyinstaller
```
4. Use PyInstaller to create a single executable file:
```
pyinstaller --onefile wingpt.py
```
After the file has been created, remember to move wingpt.exe, config.json, and shortcuts.json to the same folder.



## License

winGPT is licensed under the MIT License. See the LICENSE file for more information.

## Contributions

Contributions to winGPT are welcome. If you would like to contribute, please open a pull request or issue on this GitHub repository.
