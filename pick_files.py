# -*- coding: utf-8 -*-
"""系统原生文件/文件夹选择框（借 Windows PowerShell + WinForms 实现，不依赖 tkinter）"""
import json
import os
import subprocess
import sys
import tempfile

SCRIPT = r"""
param([string]$Mode, [string]$OutFile)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
$result = @()
if ($Mode -eq 'folder') {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = '选择包含 Excel 文件的文件夹'
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $result = @($dlg.SelectedPath) }
} else {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = '选择要打印的 Excel 文件'
    $dlg.Multiselect = $true
    $dlg.Filter = 'Excel 文件 (*.xlsx;*.xls;*.xlsm;*.xlsb)|*.xlsx;*.xls;*.xlsm;*.xlsb|所有文件 (*.*)|*.*'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $result = @($dlg.FileNames) }
}
$json = ConvertTo-Json -InputObject @($result) -Compress
[IO.File]::WriteAllText($OutFile, $json, (New-Object Text.UTF8Encoding $false))
"""


def pick(mode="files"):
    """返回被选中的路径列表；用户取消则返回空列表"""
    out_file = os.path.join(tempfile.gettempdir(), "excelprint_pick.json")
    if os.path.exists(out_file):
        try:
            os.remove(out_file)
        except Exception:
            pass
    tmp_ps = os.path.join(tempfile.gettempdir(), "excelprint_pick.ps1")
    with open(tmp_ps, "w", encoding="utf-8-sig") as fh:
        fh.write(SCRIPT)
    cmd = ["powershell.exe", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass",
           "-File", tmp_ps, "-Mode", mode, "-OutFile", out_file]
    try:
        subprocess.run(cmd, capture_output=True, timeout=300,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        pass
    paths = []
    if os.path.exists(out_file):
        try:
            with open(out_file, "r", encoding="utf-8-sig") as fh:
                data = json.load(fh)
            if isinstance(data, str):
                data = [data]
            paths = [str(p) for p in data if p]
        except Exception:
            paths = []
    return paths


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "files"
    print(json.dumps(pick(mode), ensure_ascii=False))


if __name__ == "__main__":
    main()
