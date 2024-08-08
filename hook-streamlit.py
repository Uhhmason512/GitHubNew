# hook-streamlit.py
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = collect_all('streamlit')

# Explicitly include the cli module if not collected automatically
hiddenimports.append('streamlit.cli')