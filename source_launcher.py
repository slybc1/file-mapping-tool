"""Install or launch the source distribution using a private Python environment."""
from __future__ import annotations

import argparse
from datetime import datetime
import os
from pathlib import Path
import struct
import subprocess
import sys
import traceback
import venv

ROOT = Path(__file__).resolve().parent
ENV = ROOT / '.venv'
PYTHON = ENV / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')
LOG = ROOT / 'source-launch.log'


def record(message: str) -> None:
    print(message, flush=True)
    with LOG.open('a', encoding='utf-8') as stream:
        stream.write(f'{datetime.now().isoformat(timespec="seconds")} {message}\n')


def run(command: list[str]) -> int:
    env = os.environ.copy()
    env['PYTHONUTF8'] = '1'
    env['PYTHONUNBUFFERED'] = '1'
    record('Starting: ' + subprocess.list2cmdline(command))
    with subprocess.Popen(command, cwd=ROOT, env=env, stdout=subprocess.PIPE,
                          stderr=subprocess.STDOUT, text=True, encoding='utf-8',
                          errors='replace') as child:
        assert child.stdout is not None
        with LOG.open('a', encoding='utf-8') as log:
            for line in child.stdout:
                print(line, end='', flush=True)
                log.write(line)
                log.flush()
        return child.wait()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=['setup', 'run', 'check'])
    parser.add_argument('--offline', type=Path,
                        help='Install only from this directory of approved wheel files')
    args = parser.parse_args()
    if sys.version_info < (3, 11) or struct.calcsize('P') != 8:
        print('Use 64-bit Python 3.11 or newer. Python 3.12 is recommended.')
        return 2
    try:
        record(f'Action: {args.action}; Python: {sys.version.split()[0]}')
        if args.action == 'setup':
            if not PYTHON.is_file():
                record('Creating the private .venv environment...')
                venv.EnvBuilder(with_pip=True).create(ENV)
            command = [str(PYTHON), '-m', 'pip', 'install', '--disable-pip-version-check',
                       '-r', str(ROOT / 'requirements-runtime.txt')]
            if args.offline:
                if not args.offline.is_dir():
                    raise ValueError('The offline wheel directory does not exist')
                command += ['--no-index', '--find-links', str(args.offline.resolve())]
            result = run(command)
            if result:
                record('Setup failed. See source-launch.log; contact your IT team if access is blocked.')
                return result
        if not PYTHON.is_file():
            record('Environment is not installed. Run setup_source.bat first.')
            return 2
        result = run([str(PYTHON), '-c',
                      'import openpyxl; from PySide6 import QtCore, QtGui, QtWidgets; '
                      'print("Runtime dependencies loaded successfully")'])
        if result:
            record('Runtime check failed. Run setup_source.bat or consult source-launch.log.')
            return result
        if args.action == 'run':
            result = run([str(PYTHON), str(ROOT / 'file_mapping_tool.py')])
        else:
            result = run([str(PYTHON), str(ROOT / 'file_mapping_tool.py'), '--smoke-test'])
        record(f'Finished with exit code {result}. Log: {LOG}')
        return result
    except Exception:
        error = traceback.format_exc()
        print(error)
        try:
            record(error)
        except OSError:
            print('Cannot write to this folder. Extract the ZIP to a writable local folder.')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
