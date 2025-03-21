# Merge pptx
Enjoy your presentations merging python script, meine freund.

## Prerequisites:
 - python 3.6+
 - pip

### Windows does have all the fonts already installed.

### for debian-based distros
```bash
sudo apt install ttf-mscorefonts-installer
```

### for arch-based distros
```bash
sudo pacman -S ttf-ms-win10-auto
```

## Install and run

Create venv and install dependencies:
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Run:
```bash
python merge_pptx.py
```

Result:
```console
projects/python/mpptx via 🐍 v3.13.2 (.venv) took 4s 
❯ python merge_pptx.py
Processing files ./presentations/:['1.pptx', '2.pptx', '3.pptx']...
Merged!
Removing free version Spire library 'watermarks/ads'...
Done!
Results saved to file: ./presentations/merged.pptx
```
