"""Build and validate the Windows GUI executable on a native Windows runner."""
from __future__ import annotations
import ctypes
from ctypes import wintypes
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import platform
import shutil
import struct
import subprocess
import sys
import tempfile
import time
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[2]
OUTPUT = ROOT / 'windows-output'


def pe_info(path: Path) -> dict:
    data = path.read_bytes()
    if data[:2] != b'MZ':
        raise ValueError('Not a Windows PE executable')
    offset = struct.unpack_from('<I', data, 0x3c)[0]
    if data[offset:offset + 4] != b'PE\0\0':
        raise ValueError('Invalid PE signature')
    machine = struct.unpack_from('<H', data, offset + 4)[0]
    subsystem = struct.unpack_from('<H', data, offset + 24 + 68)[0]
    if machine != 0x8664 or subsystem != 2:
        raise ValueError(f'Expected x64 Windows GUI executable: machine={machine}, subsystem={subsystem}')
    return {'machine': 'AMD64', 'subsystem': 'Windows GUI'}


def native_gui_smoke(executable: Path) -> dict:
    """Start the frozen application and close its real Tk window with WM_CLOSE."""
    user = ctypes.WinDLL('user32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user.EnumWindows.argtypes = [callback_type, wintypes.LPARAM]
    user.EnumWindows.restype = wintypes.BOOL
    user.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user.GetWindowThreadProcessId.restype = wintypes.DWORD
    user.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user.IsWindowVisible.argtypes = [wintypes.HWND]
    user.IsWindowVisible.restype = wintypes.BOOL
    user.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user.PostMessageW.restype = wintypes.BOOL

    class ProcessEntry(ctypes.Structure):
        _fields_ = [
            ('dwSize', wintypes.DWORD), ('cntUsage', wintypes.DWORD),
            ('th32ProcessID', wintypes.DWORD), ('th32DefaultHeapID', ctypes.c_size_t),
            ('th32ModuleID', wintypes.DWORD), ('cntThreads', wintypes.DWORD),
            ('th32ParentProcessID', wintypes.DWORD), ('pcPriClassBase', wintypes.LONG),
            ('dwFlags', wintypes.DWORD), ('szExeFile', wintypes.WCHAR * 260),
        ]
    kernel.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
    kernel.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel.Process32FirstW.argtypes = [wintypes.HANDLE, ctypes.POINTER(ProcessEntry)]
    kernel.Process32NextW.argtypes = [wintypes.HANDLE, ctypes.POINTER(ProcessEntry)]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]

    def descendants(parent):
        snapshot = kernel.CreateToolhelp32Snapshot(2, 0)
        if snapshot == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        children = []
        try:
            entry = ProcessEntry()
            entry.dwSize = ctypes.sizeof(entry)
            more = kernel.Process32FirstW(snapshot, ctypes.byref(entry))
            while more:
                children.append((entry.th32ProcessID, entry.th32ParentProcessID))
                more = kernel.Process32NextW(snapshot, ctypes.byref(entry))
        finally:
            kernel.CloseHandle(snapshot)
        result = {parent}
        while True:
            expanded = result | {pid for pid, ppid in children if ppid in result}
            if expanded == result:
                return result
            result = expanded

    def windows(process_ids):
        found = []
        @callback_type
        def callback(hwnd, extra):
            pid = wintypes.DWORD()
            user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value in process_ids:
                title = ctypes.create_unicode_buffer(512)
                classname = ctypes.create_unicode_buffer(128)
                user.GetWindowTextW(hwnd, title, len(title))
                user.GetClassNameW(hwnd, classname, len(classname))
                found.append((hwnd, title.value, classname.value, bool(user.IsWindowVisible(hwnd))))
            return True
        user.EnumWindows(callback, 0)
        return found

    with tempfile.TemporaryDirectory(prefix='hotkeys-frozen-smoke-') as temporary:
        directory = Path(temporary)
        settings = directory / 'HotkeysTemplateEditor' / 'templates.json'
        settings.parent.mkdir()
        settings.write_text(json.dumps({'version': 1, 'templates': {'1': 'Сборка {clip}'}}, ensure_ascii=False), encoding='utf-8')
        before = settings.read_bytes()
        environment = os.environ.copy()
        environment['APPDATA'] = str(directory)
        process = subprocess.Popen([str(executable)], env=environment, cwd=directory)
        last_windows = []
        try:
            deadline = time.monotonic() + 90
            editor = None
            while time.monotonic() < deadline:
                if process.poll() is not None:
                    raise RuntimeError(f'Executable exited before its window appeared: {process.returncode}')
                last_windows = windows(descendants(process.pid))
                consoles = [window for window in last_windows if window[2] == 'ConsoleWindowClass' and window[3]]
                if consoles:
                    raise RuntimeError('Executable created a visible console window')
                editor = next((w for w in last_windows if w[1].startswith('Мои шаблоны') and w[3]), None)
                if editor:
                    break
                time.sleep(0.5)
            if not editor:
                raise RuntimeError(f'The application window did not appear. Owned windows: {last_windows!r}')
            if not user.PostMessageW(editor[0], 0x0010, 0, 0):
                raise ctypes.WinError(ctypes.get_last_error())
            result = process.wait(timeout=20)
            if result != 0:
                raise RuntimeError(f'Executable failed on window close: {result}')
            if settings.read_bytes() != before:
                raise RuntimeError('Opening and closing the executable changed saved templates')
            return {'window_opened': True, 'console_visible': False, 'window_close_exits': True,
                    'exit_code': result, 'saved_templates_preserved': True}
        finally:
            if process.poll() is None:
                subprocess.run(['taskkill', '/PID', str(process.pid), '/T', '/F'], capture_output=True)
                process.wait(timeout=20)


def main():
    if os.name != 'nt' or struct.calcsize('P') != 8:
        raise RuntimeError('Build requires native 64-bit Windows Python')
    os.chdir(ROOT)
    sys.path.insert(0, str(ROOT))
    OUTPUT.mkdir(exist_ok=True)
    source = ROOT / 'hotkeys_with_template_editor.pyw'
    if source.read_bytes() != (ROOT / 'hotkeys_with_template_editor.py').read_bytes():
        raise RuntimeError('The .pyw entrypoint differs from the tested .py program')
    suite = unittest.defaultTestLoader.loadTestsFromName('test_hotkeys_editor')
    with (OUTPUT / 'tests.log').open('w', encoding='utf-8') as stream:
        result = unittest.TextTestRunner(stream=stream, verbosity=2).run(suite)
    if not result.wasSuccessful() or result.testsRun < 38 or result.skipped:
        raise RuntimeError(f'Windows tests did not pass completely: run={result.testsRun}, skipped={len(result.skipped)}')
    print(f'All {result.testsRun} Windows tests passed; no skips.', flush=True)
    with (OUTPUT / 'pyinstaller.log').open('w', encoding='utf-8') as stream:
        subprocess.run([
            sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onefile', '--windowed',
            '--name', 'Hotkeys', '--hidden-import', 'win32timezone',
            '--hidden-import', 'pythoncom', '--hidden-import', 'pywintypes',
            '--distpath', str(OUTPUT), '--workpath', str(ROOT / 'build' / 'pyinstaller'),
            '--specpath', str(ROOT / 'build'), str(source),
        ], check=True, stdout=stream, stderr=subprocess.STDOUT, cwd=ROOT)
    executable = OUTPUT / 'Hotkeys.exe'
    metadata = pe_info(executable)
    print('Frozen x64 GUI executable built; checking its actual window.', flush=True)
    smoke = native_gui_smoke(executable)
    checksum = hashlib.sha256(executable.read_bytes()).hexdigest()
    (OUTPUT / 'SHA256SUMS.txt').write_text(f'{checksum}  Hotkeys.exe\n', encoding='ascii')
    readme = (
        'Hotkeys для Windows 10/11 x64\n\n'
        '1. Завершите старую версию через Ctrl+Alt+Q.\n'
        '2. Запустите Hotkeys.exe. Установка Python не требуется.\n'
        '3. Программа работает, пока окно открыто; крестик завершает её работу.\n'
        'Шаблоны: %APPDATA%\\HotkeysTemplateEditor\\templates.json\n'
        'Для функций Excel нужен установленный Microsoft Excel.\n'
        'Полная инструкция: HOTKEYS_README_RU.md в архиве.\n'
    )
    (OUTPUT / 'README.txt').write_text(readme, encoding='utf-8-sig')
    report = {
        'status': 'success', 'build_commit': os.environ.get('GITHUB_SHA'),
        'source_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
        'exe_sha256': checksum, 'exe_size': executable.stat().st_size,
        'platform': platform.platform(), 'python': platform.python_version(),
        'pyinstaller': importlib.metadata.version('pyinstaller'),
        'unit_tests_passed': result.testsRun, 'unit_tests_skipped': 0,
        'pe': metadata, 'native_gui_smoke': smoke,
    }
    (OUTPUT / 'build-report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    with zipfile.ZipFile(OUTPUT / 'Hotkeys-Windows-x64.zip', 'w', zipfile.ZIP_DEFLATED) as archive:
        for name in ['Hotkeys.exe', 'README.txt', 'SHA256SUMS.txt', 'build-report.json']:
            archive.write(OUTPUT / name, arcname=name)
        archive.write(ROOT / 'HOTKEYS_README_RU.md', arcname='HOTKEYS_README_RU.md')
    print(json.dumps(report, ensure_ascii=False), flush=True)


if __name__ == '__main__':
    try:
        main()
    except Exception:
        import traceback
        OUTPUT.mkdir(exist_ok=True)
        (OUTPUT / 'failure.log').write_text(traceback.format_exc(), encoding='utf-8')
        raise
