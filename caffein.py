import time
import subprocess

while True:
    subprocess.run([
        "powershell",
        "-command",
        "$wshell = New-Object -ComObject wscript.shell; $wshell.SendKeys('{F13}')"
    ])
    print("Pressed F13")
    time.sleep(60)