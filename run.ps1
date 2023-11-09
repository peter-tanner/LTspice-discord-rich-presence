cd $(Split-Path $MyInvocation.MyCommand.Path)
.\.venv\Scripts\Activate.ps1
Start-Process "python" -ArgumentList "ltspice-discord.py" -WindowStyle Hidden
