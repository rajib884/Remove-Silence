Removes silent parts from video using [ffmpeg](https://www.ffmpeg.org/)

GUI built with [PySimpleGuiQt](https://github.com/PySimpleGUI/PySimpleGUI)

Inspired by [no-silence](https://github.com/turicas/no-silence)

## Installation

Tested on Python 3.9.2.

```shell
pip install -r requirements.txt
```

## Usage

```shell
python main.py
```

### Standalone

Standalone executable file can be created using [PyInstaller](https://github.com/pyinstaller/pyinstaller)

```shell
pip install pyinstaller
```

```shell
pyinstalelr -wF main.py
```