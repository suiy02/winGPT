import time
import pynput.keyboard
import requests
import json
import pyperclip
import re
import sys
from colorama import init, Fore, Style
import win32gui
import win32process
import win32api
from datetime import datetime, timedelta
from pprint import pprint
from collections import OrderedDict

API_URL = 'https://api.openai.com/v1/engines/davinci-codex/completions'
# default trigger word is /gpt
trigger_word = '/gpt'

class ColorPrinter:
    def __init__(self, color=''):
        init() # initialize colorama
        self.color_map = {
            'black': Fore.BLACK,
            'red': Fore.RED,
            'green': Fore.GREEN,
            'yellow': Fore.YELLOW,
            'blue': Fore.BLUE,
            'magenta': Fore.MAGENTA,
            'cyan': Fore.CYAN,
            'white': Fore.WHITE,
            'b': Style.BRIGHT,
            'u': Style.DIM 
        }
        if not color:
            self.color = Fore.WHITE # default color is white
        elif color in self.color_map:
            self.color = self.color_map[color]
        else:
            print(f"Error: {color} is not a valid color.")
            self.color = Fore.WHITE

    def set_color(self, color:str):
        if color in self.color_map:
            self.color = self.color_map[color]
        else:
            print(f"Error: {color} is not a valid color.")

    def __call__(self, text:str):
        colored_text = text
        for color in self.color_map:
            start_tag = f"<{color}>"
            end_tag = f"</{color}>"
            colored_text = colored_text.replace(start_tag, f"{self.color_map[color]}")
            colored_text = colored_text.replace(end_tag, f"{Style.RESET_ALL}{self.color}")
        print(f"{self.color}{colored_text}{Style.RESET_ALL}")

print_error = ColorPrinter('red')

def get_input() -> str:
    input_str = ""
    previous_key = None
    recording = False
    myprint = lambda x: None
    # define a callback function to handle keystrokes
    def on_press(key):
        nonlocal input_str, previous_key, recording,myprint
        try:
            if key.char == trigger_word[0] and not recording:
                recording = True
                # if the active window is wingpt, then print the input   
                window_name = get_active_window_name().lower()
                window_name = window_name.split("\\")[-1]
                if window_name == "wingpt.exe":
                    myprint = lambda x: print(x, end='', flush=True)
                    input_str += trigger_word
                    myprint(input_str)
                    key.char = ''
                else:
                    myprint = lambda x: None
        except AttributeError:
            pass            # ignore non-character keys

        if recording:
            try:
                if key.char == '\x16':
                    # if the user types ctrl+v, then paste the clipboard
                    input_str += pyperclip.paste()
                    myprint(pyperclip.paste())
                elif key.char < '\x20': # ignore control characters
                    pass
                else:
                    input_str += key.char
                    myprint(key.char)
            except AttributeError:
                # handle special keys
                if key == pynput.keyboard.Key.enter:
                    if previous_key == pynput.keyboard.Key.enter or previous_key == pynput.keyboard.Key.shift:
                        # if the user presses 2 enter in a row, or shift+enter stop monitoring keystrokes
                        return False
                    else:
                        # append a new line to the input string
                        input_str += '\n'
                        myprint('\n')
                elif key == pynput.keyboard.Key.space:
                    input_str += ' '
                    myprint(' ')
                elif key == pynput.keyboard.Key.backspace:
                    # remove the last character from the input string
                    input_str = input_str[:-1]
                    myprint('\b \b')
                else:
                    # ignore other special keys
                    pass            # ignore non-character keys
            except TypeError:
                pass            # ignore non-character keys

            if len(input_str) > len(trigger_word) and not input_str.startswith(trigger_word):
                return False
        previous_key = key

    # start monitoring keystrokes
    with pynput.keyboard.Listener(on_press=on_press) as listener:
        listener.join()
    # return the recorded input
    return input_str


def query_gpt(question:str, config:dict, chat_history=[]):
    # Set API endpoint and parameters
    url = "https://api.openai.com/v1/chat/completions"
    API_KEY = config["API_KEY"]
    temperature = config["temperature"]
    system_prompt = config["system_prompt"]
    time_out = config["time_out"]

    headers = {
        "Content-Type": "application/json",
        'Authorization': f'Bearer {API_KEY}'
    }

    if not chat_history:
        chat_history = [{"role": "system", "content": system_prompt}]
    
    chat_history.append({"role": "user", "content": question})

    # print(chat_history)
    data = {
        "model": "gpt-3.5-turbo",
        "temperature": temperature,
        "messages": chat_history
    }

    # Send API request and get response
    print_color = ColorPrinter('blue')
    print_color("waiting for chatGPT... ")
    try:
        response = requests.post(url, headers=headers, data=json.dumps(data), timeout=time_out)
    except requests.exceptions.Timeout:
        print_error("timeout")
        return "chatGPT server timed out. Please try again."
    response_json = json.loads(response.text)

    # Extract response message from API response
    message = response_json["choices"][0]["message"]["content"]
    
    # Update chat history with the model's response
    chat_history.append({"role": "assistant", "content": message})

    return message, chat_history


def query_gpt_old(question,config):
    # Set API endpoint and parameters
    url = "https://api.openai.com/v1/chat/completions"
    API_KEY = config["API_KEY"]
    temperature = config["temperature"]
    # max_tokens = config["max_tokens"]
    system_prompt = config["system_prompt"]
    time_out = config["time_out"]

    headers = {
        "Content-Type": "application/json",
        'Authorization': f'Bearer {API_KEY}'
    }
    data = {
        "model": "gpt-3.5-turbo",
        "temperature": temperature,
        "messages": [{"role": "system", "content": system_prompt}, 
                     {"role": "user", "content": question}],
    }

    # Send API request and get response
    print = ColorPrinter('blue')
    print("waiting for chatGPT... ")
    try:
        response = requests.post(url, headers=headers, data=json.dumps(data),timeout=time_out)
    except requests.exceptions.Timeout:
        print_error("timeout")
        return "chatGPT server timed out. Please try again."
    response_json = json.loads(response.text)

    # Extract response message from API response
    message = response_json["choices"][0]["message"]["content"]
    
    return message

# get config from config.json
def get_config(fname='config.json') -> dict:
    default_config = {
        "API_KEY": "",
        "trigger_word": "/gpt",
        "temperature": 0.7,
        "max_tokens": 1024,
        "time_out": 60,
        "system_prompt": "You are an AI assistant. Answer the questions in a concise and accurate way",
        "history_length": 4,
        "history_timeout_in_seconds": 60    
    }
    
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
            for key, value in default_config.items():
                if key not in config or config[key] == "":
                    config[key] = value
        if config["API_KEY"] == "":
            print_error("No API key found. Please put your <b>API_KEY</b> in <b>config.json</b> file")
            print_error("API_KEY can be found at https://beta.openai.com/account/api-keys")
            pause = input("Press any key to continue...")
            sys.exit(1)
    except FileNotFoundError:
        print_error("config.json not found")
        print_error("Please create a file named <b>config.json</b> and put your API key in it")
        print_error("API_KEY can be found at https://beta.openai.com/account/api-keys")   
        pause = input("Press any key to continue...")
        sys.exit(1)
    except json.decoder.JSONDecodeError:
        print_error("config.json is not a valid json file")
        pause = input("Press any key to continue...")
        sys.exit(1)
    return config

    
def get_shortcuts(fname='shortcuts.json') -> dict[str, str]:
    # read shortcuts from a json file
    divider = "\n------\n"
    shortcuts = {
            # default shortcuts as an example
            "<trigger_word>.sum"    : "[[sum]]",
            "[[sum]]"               : "Please summarize the following text using bullet points:" + divider,
            "<trigger_word>.revise" : "[[revise]]",
            "[[revise]]"            : "Please revise the text below to correct any spelling or grammatical errors, enhance clarity, and make it sound more natural. Also, ensure that it is concise: " + divider,
        }
 
   # if 'shortcuts.json' exists, read shortcuts from a json file
    try:
        with open(fname, 'r') as f:
            # shortcuts = json.load(f, object_pairs_hook=OrderedDict)
            shortcuts = json.load(f)
    except FileNotFoundError as e:
        print_error(fname + " not found. Using default shortcuts.")
        e.printed = True
    except json.decoder.JSONDecodeError as e:
        print_error("<b>shortcuts.json</b> format error:")
        print_error(str(e))
    except Exception as e:
        print_error(str(e))  
    return shortcuts

def replace_shortcuts(text: str,shortcuts: dict[str,str]) -> str:
    content = pyperclip.paste()
    text = text.replace(trigger_word, '<trigger_word>')

    for shortcut, long_text in shortcuts.items():
        text = text.replace(shortcut, long_text)
    
    text = text.replace('<trigger_word>', '')  
    text = text.replace('[[p]]', content)
    text = text.replace('[[p', content)
    return text

def get_active_window_name():
    window = win32gui.GetForegroundWindow()
    class_name = win32gui.GetClassName(window)
    # Check if the active window's class name matches one of the known command prompt terminal names.
    # return class_name in ["ConsoleWindowClass", "VirtualConsoleClass", "XfceTerminal", "MobaXTerm", "TMobaXtermForm"]

    _, pid = win32process.GetWindowThreadProcessId(window)
    process_name = win32process.GetModuleFileNameEx(win32api.OpenProcess(0x0400 | 0x0010, False, pid), 0)
    if __debug__:
        print("window class name:", class_name, "process name:", process_name)
    return process_name

# decorate the chatGPT response according to the active window name
def decorate_response(text: str) -> str:
    window_name = get_active_window_name().lower()
    window_name = window_name.split("\\")[-1]
    decorate_method = {
        "cmd.exe"        : lambda text: add_comment_symbol(text, ":"),
        "powershell.exe" : lambda text: add_comment_symbol(text, "#"),
        "mobaxterm.exe"  : lambda text: add_comment_symbol(text, "#"),
        "mintty.exe"     : lambda text: add_comment_symbol(text, "#"),
        "putty.exe"      : lambda text: add_comment_symbol(text, "#"),
        "wsl.exe"        : lambda text: add_comment_symbol(text, "#"),
        "wingpt.exe"     : lambda text: "",
    }
    if window_name in decorate_method:
        text = decorate_method[window_name](text)
    return text

def _typing(text:str, controler:pynput.keyboard.Controller)->None:
    if not text:
        return
    # split text by lines \n , and sleep 0.1 second after typing each line.
    lines = text.split("\n")
    if len(lines) == 1:
        controler.type(text)
    else:
        for line in lines:
            if line:
                controler.type(line + "\n")
            time.sleep(0.1)

# a function to comment out the response if it is in a command window or terminal
def add_comment_symbol(text:str, comment_symbol=""):
    if comment_symbol:
        # add comment_symbol to the beginning of each line , if not already commented
        text = "\n".join([comment_symbol + " " + line if not line.startswith(comment_symbol) else line for line in text.split("\n")])
    return text 

def usage():
    print = ColorPrinter('yellow') 

    wordart = """
       _       _______ ______ _______ 
      (_)     (_______|_____ (_______)
 _ _ _ _ ____  _   ___ _____) )  _    
| | | | |  _ \| | (_  |  ____/  | |   
| | | | | | | | |___) | |       | |   
 \___/|_|_| |_|\_____/|_|       |_|
     ChatGPT at your fingertips!
 """
    
    print(wordart)
    # print("")

    print.set_color('white')
    tag = "<cyan><b>"
    tag_end = "</b></cyan>"
    print("<b>Usage:</b>")
    print("In Outlook, Notepad, MS Word, or other text editors, ")
    print(f"Type {tag}{trigger_word}{tag_end} to start a conversation, press {tag}Enter{tag_end} key twice (x2) or {tag}Shift+Enter{tag_end} for chatGPT to respond.")
    print("")
  
    print("<b>Shortcuts:</b>")
    print("Define your prompts as shortcuts in <b>shortcuts.json</b> file, and they will be converted to their full version before being sent to ChatGPT.")
    print("Example:")
    print(f"Type {tag}{trigger_word}.sum{tag_end} to summerize text.")
    print(f"Type {tag}{trigger_word}.revise{tag_end} to revise text.")
    print(f"Use {tag}[[p]] or [[p{tag_end} as substitutes for clipboard content.")
    print(f"{tag}{trigger_word}.sum [[p]]{tag_end} will create a summary of the clipboard's content.")
    print("See <b>shortcuts.json</b> for more example prompts, and you can add your own shortcuts there.")
    print("")

    print("<b>Config file:</b>")
    print("You can change the default settings in <b>config.json</b> file.")
    print("See <b>config.json</b> for more details.")
    print("")

    print("<b>Commands:</b> ")
    print(f"{tag}{trigger_word}.clear{tag_end} to clear the chat history.")
    print(f"{tag}{trigger_word}.config{tag_end} to reload the config.json file.")
    print(f"{tag}{trigger_word}.shortcuts{tag_end} to reload the shortcuts.json file.")
    
if __name__ == '__main__':

    # load config file
    config = get_config('config.json')
    trigger_word = config["trigger_word"]
    history_length = config["history_length"]
    history_timeout_in_seconds = config["history_timeout_in_seconds"]

    # print usage
    usage()

    print_question = ColorPrinter('yellow')
    print_answer = ColorPrinter('white')
    print_other = ColorPrinter('cyan')
    controller = pynput.keyboard.Controller()
    typing = lambda text: _typing(text, controller)
    chat_history = []

    # init a datetime object with January 1, 2023 at 12:00 PM
    time_last_question = datetime(2023, 1, 1, 12, 0, 0)
    shortcuts = get_shortcuts('shortcuts.json')
    
    while True:
        try:
            # wait for user input
            input_str = get_input()
            # print("get_input:  ", input_str)
            # if the input starts with trigger_word, remove the command and send the rest to the API
            if input_str.startswith(trigger_word):
                # clear history if input_str is trigger_word+.clear
                if input_str.strip() == trigger_word + ".clear":
                    chat_history = []
                    print_other("<b>Chat history cleared</b>")
                    print_other("")
                    continue
                if input_str.strip() == trigger_word + ".config":
                    print_other("<b>Reload Config:</b> ")
                    config = get_config('config.json')
                    trigger_word = config["trigger_word"]
                    history_length = config["history_length"]
                    history_timeout_in_seconds = config["history_timeout_in_seconds"]
                    pprint(config)
                    continue
                if input_str.strip() == trigger_word + ".shortcuts":
                    print_other("<b>Reload shortcuts in shortcuts.json:</b> ")
                    shortcuts = get_shortcuts('shortcuts.json')
                    pprint(shortcuts,sort_dicts=False, indent=4, width=120, compact=True)
                    continue

                time_now = datetime.now()
                if time_now - time_last_question > timedelta(seconds=history_timeout_in_seconds):
                    chat_history = []
                    print_other("<b>Chat history cleared</b>")
                    print_other("")
                time_last_question = time_now                
 
                if len(chat_history) > history_length:
                    chat_history = chat_history[-history_length:]
               
                input_str = replace_shortcuts(input_str,shortcuts)
                if input_str.isspace():
                    print_error("input_str is empty, skipping")
                    continue
                input_str = input_str.strip()
                print_question ("Q: " + input_str)
                typing(decorate_response("# chatGPT response: # ") +" \n")
               
                output_str, chat_history = query_gpt(input_str,config,chat_history)

                # comment out the response if it is in terminal or cmd
                # output_str = decorate_response(output_str)
                # print the output to the console
                print_answer(output_str)
                # send the output to the active window (e.g. Notepad)
                typing(decorate_response(output_str))
        except Exception as e:
            print_error(str(e))
            pass

