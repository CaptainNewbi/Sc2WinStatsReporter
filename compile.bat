pyinstaller Sc2WinStatsReporter.py

pyinstaller --collect-data sc2reader --hiddenimport win32timezone -F --windowed --paths=C:\ProgramData\anaconda3\envs\SC2Programs\Lib\site-packages  Sc2WinStatsReporter.py
