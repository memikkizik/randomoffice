"""Verify the frozen EXE in an isolated directory, including a hidden WebView."""

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--exe', type=Path, default=Path('dist/standalone/RandomOffice.exe'))
    parser.add_argument('--report', type=Path, default=Path('artifacts/desktop-isolated-self-test.json'))
    args = parser.parse_args()
    executable = args.exe.resolve(strict=True)
    report = args.report.resolve()
    report.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='standalone-check-', dir=report.parent) as directory:
        isolated = Path(directory)
        target = isolated / 'RandomOffice.exe'
        shutil.copy2(executable, target)
        environment = dict(os.environ)
        environment.pop('PYTHONPATH', None)
        environment.pop('PYTHONHOME', None)
        completed = subprocess.run([str(target), '--self-test', str(report)], cwd=isolated,
                                   env=environment, creationflags=subprocess.CREATE_NO_WINDOW,
                                   timeout=90)
        if completed.returncode:
            raise SystemExit(f'Desktop test failed with exit code {completed.returncode}')
        result = json.loads(report.read_text(encoding='utf-8'))
        if not result.get('ok') or not result.get('frozen') or 'native-webview-ui' not in result.get('checks', []):
            raise SystemExit(f'Incomplete desktop verification: {result}')
        print(json.dumps(result, indent=2))


if __name__ == '__main__':
    main()
