
# winGPT

**winGPT** allows you to directly use chatGPT in the editors of Outlook, Notepad, MS Word, web browsers etc. on Windows. The idea behind winGPT was inspired by macGPT for Mac users.

**Usage:**  
Type **/plz** to start a conversation, press 'Enter' key twice (x2) or Shift+Enter for chatGPT to respond.

**Shortcuts:**  
You can define shortcuts in **shortcuts.json**, the shortcuts will be replaced with its long form before sending to chatGPT.  
For example,
* Use **/plz.sum** to summarize text.  
* Use **/plz.revise** to revise text.  
* **[[p]]** or **[[p** will be replaced with the content in the clipboard.  
* **/plz.sum [[p]]** will summarize the content in the clipboard.  

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
## License

winGPT is licensed under the MIT License. See the LICENSE file for more information.

## Contributions

Contributions to winGPT are welcome. If you would like to contribute, please open a pull request or issue on this GitHub repository.
