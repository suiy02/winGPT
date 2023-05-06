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
    def __init__(self, color=None):
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
            'b': Style.BRIGHT
        }
        if color is None:
            self.color = Fore.WHITE # default color is white
        elif color in self.color_map:
            self.color = self.color_map[color]
        else:
            print(f"Error: {color} is not a valid color.")
            self.color = Fore.WHITE

    def set_color(self, color):
        if color in self.color_map:
            self.color = self.color_map[color]
        else:
            print(f"Error: {color} is not a valid color.")

    def __call__(self, text):
        colored_text = text
        for color in self.color_map:
            start_tag = f"<{color}>"
            end_tag = f"</{color}>"
            colored_text = colored_text.replace(start_tag, f"{self.color_map[color]}")
            colored_text = colored_text.replace(end_tag, f"{Style.RESET_ALL}{self.color}")
        print(f"{self.color}{colored_text}{Style.RESET_ALL}")

print_error = ColorPrinter('red')

def get_input() -> str:
    # initialize empty string
    input_str = ""
    previous_key = None
    recording = False
    # define a callback function to handle keystrokes
    def on_press(key):
        # print(key)
        nonlocal input_str, previous_key, recording
        # print(key)
        try:
            if key.char == trigger_word[0]:
                recording = True
        except AttributeError:
            pass            # ignore non-character keys

        if recording:
            try:
                if key.char == '\x16':
                    # if the user types ctrl+v, then paste the clipboard
                    input_str += pyperclip.paste()
                elif key.char < '\x20': # ignore control characters
                    pass
                else:
                    input_str += key.char
            except AttributeError:
                # handle special keys
                if key == pynput.keyboard.Key.enter:
                    if previous_key == pynput.keyboard.Key.enter or previous_key == pynput.keyboard.Key.shift:
                        # if the user presses 2 enter in a row, or shift+enter stop monitoring keystrokes
                        return False
                    else:
                        # append a new line to the input string
                        input_str += '\n'
                elif key == pynput.keyboard.Key.space:
                    input_str += ' '
                elif key == pynput.keyboard.Key.backspace:
                    # remove the last character from the input string
                    input_str = input_str[:-1]
                else:
                    # ignore other special keys
                    pass            # ignore non-character keys
            except TypeError:
                pass            # ignore non-character keys

            # 
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
            shortcuts = json.load(f, object_pairs_hook=OrderedDict)
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
    }
    if window_name in decorate_method:
        text = decorate_method[window_name](text)
    return text

def _typing(text:str, controler:pynput.keyboard.Controller)->None:
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

    print("winGPT allows you to use chatGPT in the editors of Outlook, Notepad, MS Word, etc.")
    print("")

    print.set_color('white')
    tag = "<cyan><b>"
    tag_end = "</b></cyan>"

    print(f"Type {tag}{trigger_word}{tag_end} to start a conversation, press 'Enter' key twice (x2) or Shift+Enter for chatGPT to respond.")
    print("")
    print("<b>Shortcuts:</b>")
    print(f"Type {tag}{trigger_word}.sum{tag_end} to summarize text.")
    print(f"Type {tag}{trigger_word}.revise{tag_end} to revise text.")
    print(f"{tag}[[p]]{tag_end} will be replaced with the content in the clipboard.")
    print(f"eg. {tag}{trigger_word}.sum [[p]]{tag_end} will summarize the content in the clipboard.")
    print("See <b>shortcuts.json</b> for more shortcuts examples, and add your own shortcuts there.")
    # print config file
    print("")
    print("<b>Config file:</b>")
    print("You can change the default settings in <b>config.json</b> file.")
    print("See <b>config.json</b> for more details.")
    print("")
    
if __name__ == '__main__':

    # read API key from file
    default_config = {
        "API_KEY": "",
        "trigger_word": "/gpt",
        "temperature": 0.7,
        "max_tokens": 1024,
        "time_out": 40,
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

    # print(config)
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
                    print_other("<b>Config:</b> ")
                    pprint(config)
                    continue
                if input_str.strip() == trigger_word + ".shortcuts":
                    print_other("<b>Read in shortcuts.json:</b> ")
                    shortcuts = get_shortcuts('shortcuts.json')
                    pprint(shortcuts)
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
                output_str = decorate_response(output_str)
                # print the output to the console
                print_answer(output_str)
                # send the output to the active window (e.g. Notepad)
                typing(output_str)
        except Exception as e:
            print_error(str(e))
            pass

