from win32com.client import Dispatch
from colorama import Fore, Back, Style, init

init()

speaker = Dispatch("SAPI.SpVoice")
for o in speaker.GetAudioOutputs():
    if o.GetDescription() == 'Line 1 (Virtual Audio Cable)':
        speaker.AudioOutput = o
        print(f'Initialised {Style.DIM}{Fore.YELLOW}{o.GetDescription()}{Style.RESET_ALL} as the output device.')

# for i in speaker.GetVoices():
#     print(i.GetDescription())
#     if 'Hazel' in i.GetDescription():
#         speaker.Voice = i
#         print(f'Language set to {Style.DIM}{Fore.YELLOW}{i.GetDescription()}{Style.RESET_ALL}.')

while True:
    print(f'{Style.BRIGHT}{Fore.CYAN}>>> {Style.RESET_ALL}', end='')
    speech = input()
    speaker.Speak(str(speech))