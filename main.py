import ctypes
import json
import os
import shlex
import subprocess
import threading
import time
import winsound
import zipfile
from decimal import Decimal, InvalidOperation
from io import BytesIO
from json import JSONDecodeError
from pathlib import Path
from platform import release
from re import compile
from sys import exit
from tempfile import NamedTemporaryFile

import traceback

import PySimpleGUIQt as sg
import requests
from PIL import Image, ImageDraw

ZERO = Decimal("0")

"""
frame=953
fps=0.00
stream_0_0_q=-0.0
bitrate=N/A
total_size=N/A
out_time_us=31911000
out_time_ms=31911000
out_time=00:00:31.911000
dup_frames=0
drop_frames=0
speed=63.8x
progress=continue
"""


def create_initial_files():
    global config
    config = {
        "Font": "Calibri",
        "FontSize": 11,
        "Grab": False,
        "OnTop": False,
        "Theme": "SystemDefaultForReal",
        "RememberOptions": True,
        "TitleBar": True,
        "Splitter": "Normal",
        "Re-encode": False,
        "Encoder": "x264",
        "CRF": 23,
        "Preset": "Medium",
        "OnFinish": "Play Sound",
        "Format": ".mkv"
    }
    save_settings()


def load_settings():
    global config
    try:
        with open('config.json', 'r') as f:
            config = json.load(f)
    except (JSONDecodeError, FileNotFoundError):
        create_initial_files()

    save = False

    force_delete = ['input', 'input_filename', 'output', 'output_filename']
    for x in force_delete:
        if x in config:
            del config[x]
            save = True

    check_list = {
        "Theme": sg.theme_list(),
        "Font": ["Calibri", "Verdana", "Arial", "Courier", "Comic", "Fixedsys", "Times", "Helvetica"],
        "FontSize": list(range(5, 26)),
        "TitleBar": [True, False],
        "Grab": [False, True],
        "OnTop": [False, True],
        "RememberOptions": [True, False],
        "Splitter": ["Normal", "Fast", "Slow"],
        "Re-encode": [False, True],
        "Encoder": ["x264", "x265"],
        "CRF": [23] + list(range(17, 36)),
        "Preset": ["Medium", "Ultra Fast", "Super Fast", "Very Fast", "Faster", "Fast", "Slow", "Slower", "Very Slow"],
        "OnFinish": ["Play Sound", "Close Window", "Do Nothing"],
        "Format": [".mkv", ".mp4"],
    }
    for x in check_list:
        try:
            if config[x] not in check_list[x]:
                save = True
                config[x] = check_list[x][0]
        except KeyError:
            save = True
            config[x] = check_list[x][0]
    if save:
        save_settings()


def update_config(values):
    for key in (
            'Format',
            'OnFinish',
            'Preset',
            'CRF',
            'Encoder',
            'Re-encode',
            'Splitter',
            'TitleBar',
            'RememberOptions',
            'Theme',
            'OnTop',
            'Grab',
            'FontSize',
            'Font',

            'input_filename',
            'input',
            'output_filename',
            'output',
            'silence_threshold',
            'min_silence',
            'excess_duration',
            'minimum_interval_gap',
            'warning',
    ):
        if key in values:
            config[key] = values[key]
    save_settings()


def save_settings():
    global config
    with open('config.json', 'w') as f:
        json.dump(config, f, indent=4, sort_keys=True)


def update_window(values):
    for key, value in values.items():
        if key in (
                'input',
                'input_filename',
                'output',
                'output_filename',
                'silence_threshold',
                'min_silence',
                'excess_duration',
                'minimum_interval_gap',
                'warning'):
            window[key].update(value)


def ffmpeg_exists():
    # return False
    global p
    p = subprocess.Popen(
        ["ffmpeg", "-version"],
        shell=True,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=False
    )
    p.communicate()
    return p.returncode == 0


def check_ffmpeg():
    if not ffmpeg_exists():
        e, v = sg.Window(
            "Error",
            [[sg.Text(f'Your system does not have FFmpeg installed. \nPlease install FFmpeg.',
                      font=f'{config["Font"]} {int(config["FontSize"]) + 2}')],
             [sg.Stretch(),
              sg.B("Official Website", size=(16, 1), tooltip="https://www.ffmpeg.org/"),
              sg.B("Direct Download", size=(16, 1),
                   # tooltip="https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-full.7z"
                   tooltip="Direct Download FFmpeg binary file from ffbinaries.com"
                   ),
              sg.B("Quit", size=(6, 1))]],
            keep_on_top=config["OnTop"],
            no_titlebar=not config["TitleBar"],
            grab_anywhere=config["Grab"]
        ).read(close=True)
        if e == "Direct Download":
            if not download_ffmpeg():
                exit()
            else:
                return
        elif e == "Official Website":
            os.startfile("https://www.ffmpeg.org/")
        else:
            pass
        exit()


def download_ffmpeg():
    global window, downloading_ffmpeg
    window = sg.Window("Downloading...", [
        [sg.Text("FFmpeg", size=(50, 1), key="title")],
        [sg.T('Loading latest version info..', key='bar_text')],
        [sg.ProgressBar(100, orientation='h', key='bar')],
        [sg.Stretch(), sg.Cancel(key='Cancel', size=(10, 1.1))]
    ], disable_close=False)
    window.finalize()
    t = threading.Thread(target=background_download_ffmpeg, daemon=True)
    t.start()
    while t.is_alive():
        e, v = window.read(timeout=100)
        if e in (sg.WINDOW_CLOSED, "Cancel"):
            downloading_ffmpeg = False
        elif downloading_ffmpeg:
            window['bar'].update_bar(100 * ffmpeg_downloaded // ffmpeg_size)
            window['bar_text'].update(f"Downloaded {ffmpeg_downloaded / 1048576:.2f}/{ffmpeg_size / 1048576:.2f} MB")
    window.close()
    return download_completed


def background_download_ffmpeg():
    global downloading_ffmpeg, download_completed, window, ffmpeg_downloaded, ffmpeg_size
    try:
        r = requests.get("https://ffbinaries.com/api/v1/version/latest").json()
        window['title'].update(f"FFmpeg {r['version']}")
        link = r['bin']['windows-64']['ffmpeg']
        r = requests.get(link, stream=True)
    except requests.exceptions.ConnectionError:
        window['title'].update("Downloading Failed")
        window['bar_text'].update("Connection Error")
        return
    ffmpeg_size = int(r.headers.get('content-length'))
    temp_zip_file = NamedTemporaryFile(delete=False).name
    download_completed = False
    downloading_ffmpeg = True
    with open(temp_zip_file, "wb") as f:
        ffmpeg_downloaded = 0
        for downloaded_chunk in r.iter_content(chunk_size=4096):
            ffmpeg_downloaded += len(downloaded_chunk)
            f.write(downloaded_chunk)
            if not downloading_ffmpeg:
                break
    if downloading_ffmpeg and ffmpeg_downloaded == ffmpeg_size:
        window['bar_text'].update(f"Extracting File..")
        with zipfile.ZipFile(temp_zip_file, 'r') as fz:
            for file in fz.infolist():
                if file.filename == 'ffmpeg.exe':
                    fz.extract(file, os.getcwd())
                    window['bar_text'].update(f"Done!")
                    download_completed = True
                    break
    os.remove(temp_zip_file)


def main_window():
    menu_def = [['Misc', ['Settings', 'Reset Everything']],
                ['Help', 'About...']]
    file_types = (("Video", ".mp4"), ("Video", ".mkv"), ("Video", ".webm"), ("Video", ".avi"), ("ALL Files", "*.*"))
    tooltip = {
        'silence_threshold': 'Minimum volume level to accept as non-silence, in dB\n'
                             'The higher the decibel level, the louder the noise cutoff.',
        'min_silence': 'Minimum duration for each silence interval detection, in seconds',
        'excess_duration': 'Time to add before and after each cut interval, in seconds',
        'minimum_interval_gap': 'If less than this, will merge intervals, in seconds',
    }
    layout = [[sg.Menu(menu_def)],
              [sg.Text('Input', size=(6, 1)),
               sg.Input(key='input_filename', disabled=True, text_color='black'),
               sg.Input(key='input', visible=False, enable_events=True),
               sg.FilesBrowse(key='FileBrowse', size=(12, 1.1))],
              [sg.Text('Output', size=(6, 1)),
               sg.Input(key='output_filename', disabled=True, text_color='black'),
               sg.Input(key='output', visible=False, enable_events=True),
               sg.SaveAs(key='SaveAs', size=(12, 1.1), file_types=file_types, disabled=True)],
              [sg.Text('Silence Threshold (dB)', tooltip=tooltip['silence_threshold'], size=text_size),
               # sg.Stretch(),
               sg.Input(key='silence_threshold', default_text='-35', size=input_size, justification="r",
                        enable_events=True),
               sg.Stretch(), sg.VerticalSeparator(pad=(0, 0)), sg.Stretch(),
               sg.Text('Minimum Silence Duration (s)', tooltip=tooltip['min_silence'], size=text_size),
               # sg.Stretch(),
               sg.Input(key='min_silence', default_text='10', size=input_size, justification="r",
                        enable_events=True)],
              [sg.Text('Excess Duration (s)', tooltip=tooltip['excess_duration'], size=text_size),
               # sg.Stretch(),
               sg.Input(key='excess_duration', default_text='0.5', size=input_size, justification="r",
                        enable_events=True),
               sg.Stretch(), sg.VerticalSeparator(pad=(0, 0)), sg.Stretch(),
               sg.Text('Minimum Interval Gap (s)', tooltip=tooltip['minimum_interval_gap'], size=text_size),
               # sg.Stretch(),
               sg.Input(key='minimum_interval_gap', default_text='0.3', size=input_size, justification="r",
                        enable_events=True)],
              # [sg.Text('Video Splitter'),
              #  sg.Radio('Fast but imprecise', key='splitter', group_id=1, default=True),
              #  sg.Radio('Slow but precise', key='splitter_inv', group_id=1)],
              [sg.Text(key='warning', text_color='red', visible=False)],
              [sg.Stretch(), sg.Button('Start', disabled=True, size=(20, 1.1)), sg.Stretch(),
               sg.Button('Cancel', size=(20, 1.1)), sg.Stretch()]]
    return sg.Window(
        'No Silence',
        layout,
        keep_on_top=config["OnTop"],
        no_titlebar=not config["TitleBar"],
        grab_anywhere=config["Grab"],
        resizable=False,
        finalize=True
    )


def run_main_window():
    def update_output_filename():
        global input_format, output_format
        input_format = original.suffix
        output_format = config["Format"] if config["Re-encode"] else input_format
        output = original.parent / original.name.replace(original.suffix, "-nosilence" + output_format)
        # window['SaveAs'].update(disabled=False)  # todo: Qt bug
        window['SaveAs'].QT_QPushButton.setDisabled(False)
        if output.exists():
            # Todo: Main window can be closed while popup is visible, closing main window causes program to run forever
            output = overwrite_popup(original, output)
        window['output_filename'].update(output.name)
        window['output'].update(output.as_posix())

    warning = ''
    previously_selected_input = ''
    show_not_supported_warning = False
    global window, run, total_length
    while True:
        event, values = window.read()
        if event == 'Start':
            run = True
            break
        elif event in (sg.WIN_CLOSED, 'Cancel'):
            break
        elif event == 'input':
            if values['input'] != '':
                previously_selected_input = values['input']
                original = Path(values['input'])
                window['input_filename'].update(original.name)
                window['output_filename'].update('')
                window['output'].update('')
                window.disable()
                window.finalize()
                total_length = video_length(values['input'])
                if total_length is None:
                    everything_ok[0] = False
                    show_not_supported_warning = True
                    # window['SaveAs'].update(disabled=False)  # todo: Qt bug
                    window['SaveAs'].QT_QPushButton.setDisabled(True)
                else:
                    everything_ok[0] = True
                    show_not_supported_warning = False
                    update_output_filename()

                window.enable()
            else:
                window['input'].update(previously_selected_input)
        elif event == 'output':
            if values['output'] != '':
                window['output_filename'].update(Path(values['output']).name)
                everything_ok[0] = True
        elif event in ('silence_threshold', 'min_silence', 'excess_duration', 'minimum_interval_gap'):
            warning = ''
            for e in ('silence_threshold', 'min_silence', 'excess_duration', 'minimum_interval_gap'):
                name = e.replace("_", " ").title()
                if values[e] != '':
                    try:
                        current_value = Decimal(values[e])
                    except InvalidOperation:
                        warning = f'{name}: Invalid Duration'
                    else:
                        if not limit[e][0] <= current_value <= limit[e][1]:
                            warning = f'{name}: Enter value between {limit[e][0]} and {limit[e][1]}'
                else:
                    warning = f'{name}: Duration Can not be empty'
            everything_ok[1] = warning == ''
        elif event == 'About...':
            sg.popup_ok('Created By Rajibul Hassen.\nDate: 22 Feb 2021\nrajibridoy884@gmail.com', title='About',
                        no_titlebar=True, background_color='lightblue', grab_anywhere=True)
        elif event == 'Settings':
            update_config(values)
            window.close()
            advanced_settings_window()
            window = main_window()
            update_window(config)
            window.disable()
            original = Path(values['input'])
            update_output_filename()
            window.enable()
        elif event == 'Reset Everything':
            window_position = window.CurrentLocation()
            create_initial_files()
            load_settings()
            warning = ''
            window.close()  # TODO: Needs to change with adv settings (?? what ??)
            sg.change_look_and_feel(config["Theme"])
            window = main_window()
            window.finalize()
            window.move(window_position[0], window_position[1] - 38)
            del window_position

        # Todo: window does not shrink when set visible false
        if show_not_supported_warning:
            window['warning'].update('File Not Supported!', visible=True)
        else:
            window['warning'].update(warning, visible=warning != '')
        window['Start'].QT_QPushButton.setDisabled(not all(everything_ok))
    if run:
        update_config(values)

    window.close()
    return values


def overwrite_popup(original, output):
    e, v = sg.Window(
        "OverWrite?",
        [[sg.Text(f'Predefined output file "{output.name}" already exists, overwrite?')],
         [sg.B("Yes", size=input_size), sg.B("No", size=input_size)]],
        keep_on_top=True,
        no_titlebar=not config["TitleBar"],
        grab_anywhere=config["Grab"],
    ).read(close=True)
    if e != 'Yes':
        t = 1
        while output.exists():
            output = original.parent / original.name.replace(original.suffix,
                                                             f"-nosilence-v{t}" + output_format)
            t += 1
    return output


def advanced_settings_window():
    global window
    d = not config["Re-encode"]
    layout = [
        # [sg.Stretch(), sg.B("Theme Previewer", size=(20, 1.1))],
        # [sg.T("Font", ),
        #  sg.Combo(["Arial", "Calibri", "Courier", "Comic", "Fixedsys", "Times", "Verdana", "Helvetica"], readonly=True,
        #           default_value=config['Font'], key="Font")],
        # [
        #     sg.T("Font Size"),
        #     sg.Slider(default_value=config["FontSize"], range=(5, 25), orientation="h", size=(4, 0.4), border_width=1,
        #               key="FontSize")
        # ],
        [
            sg.T("Video Splitter: "),
            sg.Combo(["Fast", "Normal", "Slow"], key="Splitter", readonly=True, default_value=config["Splitter"])
        ],
        [
            sg.Checkbox("Re-encode Video", key="Re-encode", default=config["Re-encode"], enable_events=True)
        ],
        [sg.Frame(
            "",
            [
                [
                    sg.T("Format"),
                    sg.Combo([".mp4", ".mkv"], key="Format", readonly=True, default_value=config["Format"], disabled=d)
                ],
                [
                    sg.T("Encoder"),
                    sg.Combo(["x264", "x265"], key="Encoder", readonly=True, default_value=config["Encoder"],
                             disabled=d)
                ],
                [
                    sg.T("Constant Rate Factor (CRF): "),
                    sg.Combo(list(range(17, 36)), key="CRF", default_value=config["CRF"], readonly=True, disabled=d)
                ],
                [
                    sg.T("Preset: "),
                    sg.Combo(
                        ["Ultra Fast", "Super Fast", "Very Fast", "Faster", "Fast", "Medium", "Slow", "Slower",
                         "Very Slow"],
                        key="Preset", default_value=config["Preset"], readonly=True, disabled=d)
                ],
            ],
            pad=(0, 0),
            border_width=0,
        )],
        [sg.T("Theme"), sg.Combo(sg.theme_list(), readonly=True, default_value=sg.theme(), key="Theme")],
        [
            sg.T("After Finishing: "),
            sg.Combo(["Close Window", "Play Sound", "Do Nothing"], default_value=config["OnFinish"], key="OnFinish",
                     readonly=True)
        ],
        [
            sg.Checkbox("Show Title Bar", key="TitleBar", default=config["TitleBar"]),
            sg.Checkbox("Grab Anywhere", key="Grab", default=config["Grab"]),
            sg.Checkbox("Keep on Top", key="OnTop", default=config["OnTop"]),
        ],
        [
            sg.Checkbox("Remember Last Used Values", key="RememberOptions", default=config["RememberOptions"]),
        ],
        [sg.OK(size=(23, 1)), sg.Cancel(size=(23, 1))]
    ]
    window = sg.Window(
        "Settings",
        layout,
        keep_on_top=config["OnTop"],
        no_titlebar=not config["TitleBar"],
        grab_anywhere=config["Grab"],
        finalize=True,
        resizable=False
    )
    while True:
        e, v = window.read()
        if e in (sg.WINDOW_CLOSED, "Cancel"):
            break
        elif e == "OK":
            update_config(v)
            sg.theme(config["Theme"])
            break
        elif e == "Re-encode":
            d = not v["Re-encode"]
            window["Format"].update(disabled=d)
            window["Encoder"].update(disabled=d)
            window["Preset"].update(disabled=d)
            window["CRF"].update(disabled=d)
    window.close()


def second_window(filename):
    layout = [[sg.Text(filename)],
              [sg.Text(f"Video length: {total_length // 60}m {round(total_length % 60)} s", key='in_dur')],
              [sg.Text('(1/3) Detecting silence intervals...', key='info')],
              [sg.Image(data=progress(ZERO), key='bar', size_px=(w, h), pad=(0, 0))],
              [sg.T('Size: N/A', key='size'), sg.T('BitRate: 0 kbits/s', key='bitrate'), sg.T('Speed: 0x', key='speed'),
               sg.T('Elapsed: 0m 0s', key='time')],
              [sg.Stretch(), sg.Button("Stop", key="Cancel", size_px=(150, 40), pad=(10, 0))]
              ]
    return sg.Window(
        'No Silence',
        layout,
        disable_close=True,
        keep_on_top=config["OnTop"],
        no_titlebar=not config["TitleBar"],
        grab_anywhere=config["Grab"],
        resizable=False,
        finalize=True
    )


def detect_silence(values):
    cmd = (rf'ffmpeg -progress - -nostats -i "{values["input"]}" '
           rf'-vf select="eq(pict_type\,I)",showinfo -af silencedetect=noise={values["silence_threshold"]}dB:d={values["min_silence"]} -vsync 0 '
           rf"-f null -")

    window['size'].update(f"Size: N/A")
    window['bitrate'].update(f"BitRate: N/A")
    t = threading.Thread(target=background_silence_detect, args=(cmd,), daemon=True)
    t.start()
    canceled = False
    while t.is_alive():
        e, v = window.read(timeout=100)
        if e in (sg.WIN_CLOSED, "Cancel") and not canceled:
            canceled = True
            try:
                p.communicate(input="q".encode("utf8"), timeout=0)  # todo: Do I need to send 'q'?
            except subprocess.TimeoutExpired:
                p.terminate()
        time_diff = int(time.time()) - start_time
        window['bar'].update(data=progress_bar)
        window['speed'].update(f"Speed: {speed}x")
        window['time'].update(f"Elapsed: {time_diff // 60:2d}m {time_diff % 60:2d}s")
    if canceled:
        window['info'].update('Canceled!')
        close(status="Canceled")


def write_error(cmd=""):
    if p.returncode != 0:  # todo: gui version of this
        with open(f"error {time.strftime('%Y-%m-%d %H%M%S')}.log", "wb") as f:
            f.write(f"{''.join(traceback.format_stack())}\n"
                    f"Error running command {cmd}:\n{''.join(stderr)}\n"
                    f"Exit Code {p.returncode}".encode('utf8'))


def background_silence_detect(cmd):
    global p, stderr
    global speed, progress_bar
    stderr = []
    p = subprocess.Popen(
        cmd if type(cmd) == list else shlex.split(cmd),
        shell=True,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=False
    )
    while True:
        try:
            line = p.stdout.readline().decode("utf8", errors='replace')
        except ValueError:
            print("Value Error")
            break
        stderr.append(line)
        line = line.strip()
        if line == '' and p.poll() is not None:
            break
        result = keyframe_regex.search(line)
        if result:
            keyframes.append(Decimal(result.groupdict()['t']))
        else:
            result = silence_regex.search(line)
            if result:  # todo: gracefully
                add_silence(**result.groupdict())
            else:
                result = us_regex.search(line)
                if result:
                    # window['bar'].update(data=progress(Decimal(result.groupdict()['t'][:-3]) / 1000))
                    progress_bar = progress(Decimal(result.groupdict()['t'][:-3]) / 1000)
                else:
                    result = speed_regex.search(line)
                    if result:
                        # window['speed'].update(f"Speed: {result.groupdict()['speed']}x")
                        speed = result.groupdict()['speed']
    if ZERO not in keyframes:
        keyframes.insert(0, ZERO)
    if total_length not in keyframes:
        keyframes.append(total_length)
    write_error(cmd)


def add_silence(n, t):
    if n == "start" and (len(silence_intervals) == 0 or silence_intervals[-1]["end"] is not None):
        silence_intervals.append({"start": max(Decimal(t), ZERO), "end": None})
    elif n == "end" and silence_intervals[-1]["end"] is None:
        silence_intervals[-1]["end"] = min(Decimal(t), total_length)
    else:
        raise ValueError("Error!")  # TODO


def progress(current_time):  # todo: reuse image?
    img = Image.new('RGB', (w, h), color=colors[0])
    img1 = ImageDraw.Draw(img)

    t = []
    for silence in silence_intervals:
        t.append(silence["start"])
        if silence["end"] is not None:
            t.append(silence["end"])
    tc = False
    tn = ZERO
    for i in range(0, w):
        if i / w < (current_time / total_length):
            if total_length * i / w > tn:
                try:
                    tn = t.pop(0)
                except IndexError:
                    tn = total_length
                tc = not tc
            fil = colors[1] if tc else colors[4]
            img1.line([(i, 0), (i, h)], fill=fil, width=1)
        else:
            break
    with BytesIO() as output:
        img.save(output, format="PNG")
        im_data = output.getvalue()
    return im_data


def process_silence(values):
    global non_silence_intervals, non_silence_length
    if len(silence_intervals) == 0:
        window['info'].update("No Silence Detected!")
        close(status="Error")
    else:
        window['info'].update('Silence Detected')

    if silence_intervals[0]["start"] > 0:
        silence_intervals.insert(0, {"start": ZERO, "end": ZERO})
    if silence_intervals[-1]["end"] < total_length:
        silence_intervals.append({"start": total_length, "end": total_length})
    window['bar'].update(data=progress(total_length))

    window['info'].update('Calculating non-silence intervals...')
    pairs = zip(silence_intervals, silence_intervals[1:])
    excess_duration = Decimal(values['excess_duration'])
    total_added = ZERO
    for index, (silence_1, silence_2) in enumerate(pairs, start=1):
        if config['Splitter'] == "Fast":
            start = silence_1["end"]
            end = silence_2["start"]
            new_start = ZERO
            new_end = total_length
            for new_end in sorted(keyframes):
                if new_end < start:
                    new_start = new_end
                if new_end >= end:
                    break
            total_added += start - new_start + new_end - end
            start, end = new_start - excess_duration, new_end + excess_duration
            if start < 0:
                start = ZERO
            non_silence_intervals.append({
                "start": start,
                "end": end,
                "method": "-c:v copy"
            })
        elif config['Splitter'] == "Normal":
            start = silence_1["end"] - excess_duration
            end = silence_2["start"] + excess_duration
            copy_start = ZERO
            copy_end = total_length
            for copy_start in sorted(keyframes):
                if copy_start > start:
                    break
            for copy_end in sorted(keyframes, reverse=True):
                if copy_end < end:
                    break
            if start < 0:
                start = ZERO
            if start < copy_start < end and copy_start < copy_end and start < copy_end < end:
                non_silence_intervals.append({
                    "start": start,
                    "end": copy_start,
                    "method": "-c:v libx264 -crf 18 -tune zerolatency"
                })
                non_silence_intervals.append({
                    "start": copy_start,
                    "end": copy_end,
                    "method": "-c:v copy"
                })
                non_silence_intervals.append({
                    "start": copy_end,
                    "end": end,
                    "method": "-c:v libx264 -crf 18 -tune zerolatency"
                })
            else:
                non_silence_intervals.append({
                    "start": start,
                    "end": end,
                    "method": "-c:v libx264 -crf 18 -tune zerolatency"
                })
        else:
            start = silence_1["end"] - excess_duration
            end = silence_2["start"] + excess_duration
            if start < 0:
                start = ZERO
            non_silence_intervals.append({
                "start": start,
                "end": end,
                "method": "-c:v libx264 -crf 18 -tune zerolatency"
            })

    window['info'].update('Optimizing non-silence intervals...')
    intervals = non_silence_intervals
    minimum_gap = Decimal(values['minimum_interval_gap'])
    finished = False
    while not finished:
        finished = True
        result = []
        last = len(intervals) - 1
        index = 0
        while index <= last:
            interval = intervals[index]
            if index < last:
                next_interval = intervals[index + 1]
                if next_interval["start"] - interval["end"] <= minimum_gap and (
                        interval.get("method") != "-c:v copy" and next_interval.get("method") != "-c:v copy"):
                    # merge these two intervals
                    finished = False
                    result.append({"start": interval["start"], "end": next_interval["end"]})
                    index += 1  # jump the next one
                else:
                    result.append(interval)

            else:  # last interval, just append
                result.append(interval)

            index += 1
        intervals = result

    non_silence_intervals = [interval for interval in intervals if interval["end"] - interval["start"] > 0]

    non_silence_length = ZERO
    for x in non_silence_intervals:
        non_silence_length += x['end'] - x['start']
    silence_length = total_length - non_silence_length
    window['in_dur'].update(
        f"Video length: {total_length // 60}m {round(total_length % 60)} s, Silence {silence_length // 60}m {round(silence_length % 60)}s ({round(100 * silence_length / total_length, 2)}%)")


def video_length(filename):
    # cmd = f'ffprobe -v quiet -of csv=p=0 -show_entries format=duration "{filename}"'
    # return Decimal(p.stdout.read().decode().strip())
    global p
    cmd = f'ffmpeg -hide_banner -i "{filename}" -an -frames 1 -f null -'
    p = subprocess.Popen(
        cmd if type(cmd) == list else shlex.split(cmd),
        shell=True,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=False
    )
    output = p.communicate()[0]
    if p.returncode == 0:
        result = dur_regex.search(output.decode()).groupdict()
        hour = int(result.get('hour', 0))
        minute = int(result.get('min', 0))
        sec = Decimal(result.get('sec', 0))
        return (hour * 3600) + (minute * 60) + sec or None
    return None


def split_video(values):
    global temporary_filename, splitnames
    window['info'].update('(2/3) Splitting video...')
    temporary_filename = NamedTemporaryFile(delete=False).name
    print(temporary_filename)
    splitnames = []
    total_parts = len(non_silence_intervals)
    total_size = 0

    canceled = False
    for n, interval in enumerate(non_silence_intervals):
        if canceled:
            break
        splitname = f"{temporary_filename}-{n:05d}{input_format}"  # todo: codec not currently supported in container error using libx264
        splitnames.append(splitname)
        fast_seek = max(ZERO, interval["start"] - Decimal("2.0"))
        if interval["start"] == ZERO:
            cmd = f'ffmpeg -progress - -nostats -i "{values["input"]}" -to {interval["end"]} {interval["method"]} -c:a copy -copyts -avoid_negative_ts 1 "{splitname}"'
        else:
            cmd = f'ffmpeg -progress - -nostats -ss {fast_seek} -i "{values["input"]}" -ss {interval["start"]} -to {interval["end"]} {interval["method"]} -c:a copy -copyts -avoid_negative_ts 1 "{splitname}"'

        window['info'].update(f"(2/3) Splitting video...{n + 1}/{total_parts}")
        t = threading.Thread(target=background_split_video, args=(cmd, n, total_size), daemon=True)
        t.start()
        while t.is_alive():
            e, v = window.read(timeout=100)
            if e in (sg.WIN_CLOSED, "Cancel") and not canceled:
                canceled = True
                try:
                    p.communicate(input="q".encode("utf8"), timeout=0)  # todo: Do I need to send 'q'?
                except subprocess.TimeoutExpired:
                    p.terminate()
            time_diff = int(time.time()) - start_time
            window['bar'].update(data=progress_bar)
            window['speed'].update(f"Speed: {speed}x")
            window['size'].update(f"Size: {size}")
            window['bitrate'].update(f"BitRate: {bitrate}")
            window['time'].update(f"Elapsed: {time_diff // 60:02d}m {time_diff % 60:02d}s")
        total_size += os.path.getsize(splitname)
        if p.returncode != 0:
            break
    if canceled:
        remove_temp_files()
        window['info'].update("Canceled!")
        close(status="Canceled")
    elif p.returncode != 0:
        remove_temp_files()
        window['info'].update("Error!")
        close(status="Error")


def background_split_video(cmd, n, total_size):
    global p, stderr
    global speed, size, bitrate, progress_bar
    stderr = []
    p = subprocess.Popen(
        cmd if type(cmd) == list else shlex.split(cmd),
        shell=True,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=False
    )
    while True:
        try:
            line = p.stdout.readline().decode("utf8", errors='replace')
        except ValueError:
            print("Value Error")
            break
        stderr.append(line)
        line = line.strip()
        if line == '' and p.poll() is not None:
            break

        result = us_regex.search(line)
        if result:
            progress_bar = progress2(n, Decimal(result.groupdict()['t'][:-3] or '0') / 1000)
        else:
            result = speed_regex.search(line)
            if result:
                speed = result.groupdict()['speed']
            else:
                if size_regex.search(line):
                    try:
                        size = f"{(total_size + int(line[11:])) / 1048576:.2f} MB"
                    except ValueError:
                        size = line[11:]
                else:
                    result = bitrate_regex.search(line)
                    if result:
                        bitrate = result.groupdict()['bitrate']

    write_error(cmd)


def progress2(current_enumerator_number, current_time):  # todo: better?
    img = Image.new('RGB', (w, h), color=colors[4])
    img1 = ImageDraw.Draw(img)

    for i, st in enumerate(non_silence_intervals):
        if i <= current_enumerator_number:
            c = colors[2]
        else:
            c = colors[1]
        r_s = round(w * (st['start']) / total_length)
        r_e = round(w * (st['end']) / total_length)
        if i == current_enumerator_number:
            a = round(w * (st['start'] + current_time) / total_length)
            img1.rectangle([(r_s, 0), (a, h)], fill=colors[2], width=1)
            img1.rectangle([(a, 0), (r_e, h)], fill=colors[1], width=1)
        else:
            img1.rectangle([(r_s, 0), (r_e, h)], fill=c, width=1)

    with BytesIO() as output:
        img.save(output, format="PNG")
        im_data = output.getvalue()
    return im_data


def remove_temp_files():
    # return None
    global temporary_filename
    window['info'].update("Removing temp files..")
    if os.path.exists(temporary_filename):
        os.remove(temporary_filename)
        temporary_filename = None
        while len(splitnames) > 0:
            splitname = splitnames.pop(0)
            if os.path.exists(splitname):
                os.remove(splitname)


def merge_video(values):
    if temporary_filename is None or len(splitnames) == 0:
        return
    window['info'].update("(3/3) Merging video parts...")
    with open(temporary_filename, mode="w", encoding="utf8") as f:
        for splitname in splitnames:
            f.write(f"file '{splitname}'\n")
    if config["Re-encode"]:
        cmd = f'ffmpeg -progress - -nostats -y -safe 0 -f concat -i "{temporary_filename}" -c:v lib{config["Encoder"]} -crf {config["CRF"]} -preset {config["Preset"].replace(" ", "").lower()} -tune animation "{values["output"]}"'
    else:
        cmd = f'ffmpeg -progress - -nostats -y -safe 0 -f concat -i "{temporary_filename}" -c copy "{values["output"]}"'
    t = threading.Thread(target=background_merge_video, args=(cmd,), daemon=True)
    t.start()
    canceled = False
    while t.is_alive():
        e, v = window.read(timeout=100)
        if e in (sg.WIN_CLOSED, "Cancel") and not canceled:
            canceled = True
            try:
                p.communicate(input="q".encode("utf8"), timeout=0)  # todo: Do I need to send 'q'?
            except subprocess.TimeoutExpired:
                p.terminate()
        time_diff = int(time.time()) - start_time
        window['bar'].update(data=progress_bar)
        window['speed'].update(f"Speed: {speed}x")
        window['size'].update(f"Size: {size}")
        window['bitrate'].update(f"BitRate: {bitrate}")
        window['time'].update(f"Elapsed: {time_diff // 60:02d}m {time_diff % 60:02d}s")
    remove_temp_files()
    if canceled:
        window['info'].update("Canceled!")
    elif p.returncode != 0:
        window['info'].update("Error!")
    else:
        window['info'].update("Finished!")
    close(status="Finished")


def background_merge_video(cmd):
    global p, stderr
    global speed, size, bitrate, progress_bar
    stderr = []
    p = subprocess.Popen(
        cmd if type(cmd) == list else shlex.split(cmd),
        shell=True,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=False
    )
    while True:
        try:
            line = p.stdout.readline().decode("utf8", errors='replace')
        except ValueError:
            print("Value Error")
            break
        stderr.append(line)
        line = line.strip()
        if line == '' and p.poll() is not None:
            break

        result = us_regex.search(line)
        if result:
            progress_bar = progress3(Decimal(result.groupdict()['t'][:-3] or '0') / 1000)
        else:
            result = speed_regex.search(line)
            if result:
                speed = result.groupdict()['speed']
            else:
                if size_regex.search(line):
                    try:
                        size = f"{(int(line[11:])) / 1048576:.2f} MB"
                    except ValueError:
                        size = line[11:]
                else:
                    result = bitrate_regex.search(line)
                    if result:
                        bitrate = result.groupdict()['bitrate']
    write_error(cmd)


def progress3(current_time):
    img = Image.new('RGB', (w, h), color=colors[4])
    img1 = ImageDraw.Draw(img)
    a = round(w * current_time / non_silence_length)
    img1.rectangle([(0, 0), (a, h)], fill=colors[1], width=1)

    with BytesIO() as output:
        img.save(output, format="PNG")
        im_data = output.getvalue()
    return im_data


def close(status):
    if status == "Finished":
        if config["OnFinish"] == "Play Sound":
            winsound.MessageBeep(winsound.MB_OK)
        elif config["OnFinish"] == "Close Window":
            window.close()
            exit()
        elif config["OnFinish"] == "Do Nothing":
            pass
    elif status == "Error":
        winsound.MessageBeep(winsound.MB_ICONHAND)
    elif status == "Canceled":
        pass
    window['Cancel'].update("Close")
    window.read()
    window.close()
    exit()


p = None
run = False
config = {}
stderr = []
silence_intervals = []
non_silence_intervals = []
keyframes = []
total_length = ZERO
non_silence_length = ZERO
temporary_filename = None
splitnames = []
w, h = 600, 25
colors = ["#B4BCF7", "#4C8CF5", "#2B2BF7", "#0228B9", "#00059A"]
everything_ok = [False, True]
text_size = (25, 1)
input_size = (6, 1)
start_time = None

input_format = None
output_format = None

downloading_ffmpeg = False
download_completed = False
ffmpeg_downloaded = 0
ffmpeg_size = 0

# noinspection RegExpRedundantEscape
keyframe_regex = compile(r"\A\[Parsed_showinfo_1 @ \w+?\] n: *?\d+? pts: *?\d+? pts_time:(?P<t>[\d\.]+?) +?pos:")
# noinspection RegExpRedundantEscape
silence_regex = compile(r"\A\[silencedetect @ \w+?\] silence_(?P<n>start|end): (?P<t>[\d\.]+)+")
# noinspection RegExpRedundantEscape
bitrate_regex = compile(r"\Abitrate= +?(?P<bitrate>[a-zA-Z0-9\.]+?\/s)")
dur_regex = compile(r'Duration: (?P<hour>\d{2}):(?P<min>\d{2}):(?P<sec>\d{2})\.(?P<ms>\d{2})')
us_regex = compile(r"\Aout_time_us=(?P<t>[0-9]+)")
speed_regex = compile(r"\Aspeed=(?P<speed>.+?)x")
size_regex = compile(r"\Atotal_size=.+")

speed = ''
size = ''
bitrate = ''
with BytesIO() as temp_output:
    Image.new('RGB', (w, h), color=colors[0]).save(temp_output, format="PNG")
    progress_bar = temp_output.getvalue()

limit = {
    'silence_threshold': (Decimal(-45), Decimal(0)),
    'min_silence': (Decimal(0), Decimal(100)),
    'excess_duration': (Decimal(0), Decimal(2)),
    'minimum_interval_gap': (Decimal(0), Decimal(2)),
}
if int(release()) >= 8:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)

load_settings()
sg.theme(config['Theme'])
sg.set_options(font=f'{config["Font"]} {int(config["FontSize"])}')
check_ffmpeg()
window = main_window()
if config['RememberOptions']:
    update_window(config)

data = run_main_window()

if not run:
    exit()

start_time = int(time.time())
window = second_window(data['input_filename'])

detect_silence(data)
process_silence(data)
split_video(data)
merge_video(data)
