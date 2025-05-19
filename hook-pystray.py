from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files('pystray', subdir='_darwin')
hiddenimports = ['pystray._darwin']