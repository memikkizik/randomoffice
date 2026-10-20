"""Package the self-contained Windows executable without an external DLL folder."""

import argparse
import hashlib
import json
from pathlib import Path
import zipfile


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--exe', type=Path, default=Path('dist/standalone/RandomOffice.exe'))
    parser.add_argument('--output', type=Path, default=Path('randomoffice/static/downloads/RandomOffice-Windows.zip'))
    args = parser.parse_args()
    from PyInstaller.archive.readers import CArchiveReader
    archive = CArchiveReader(str(args.exe))
    if not any(Path(name).name.lower() == 'python312.dll' for name in archive.toc):
        raise SystemExit('Refusing to package: executable does not contain python312.dll')
    instructions = (
        'RandomOffice — Windows\n\n'
        'Bu paket tek dosyalı uygulama içerir. Ayrı Python kurulumu veya _internal klasörü gerekmez.\n'
        '1. ZIP dosyasına sağ tıklayıp Tümünü Ayıkla seçeneğini kullanın.\n'
        '2. Masaüstü/Başlat kısayolu ve Windows dosya türleri için Install_RandomOffice.bat dosyasını çalıştırın.\n'
        '3. Kurulumdan sonra masaüstündeki RandomOffice simgesini kullanın. Kurulum istemiyorsanız RandomOffice.exe doğrudan açılabilir.\n'
        '4. Varsayılan seçimi: uygulama Ayarlar > Windows varsayılan uygulamaları. DOCX/XLSX/PPTX için RandomOffice seçin.\n'
        'İlk açılışta çalışma ortamı hazırlandığı için birkaç saniye bekleyin.\n'
        'Önceki ZIP içinden açılan eski EXE yerine bu yeni dosyayı kullanın.\n'
        'Windows 64 bit ve Microsoft Edge WebView2 çalışma ortamı gerekir.\n'
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    staging = args.output.with_suffix('.zip.tmp')
    with zipfile.ZipFile(staging, 'w', compression=zipfile.ZIP_DEFLATED) as package:
        package.write(args.exe, 'RandomOffice.exe')
        installer=Path('scripts/install_desktop.ps1')
        if installer.exists():
            package.write(installer,'install_desktop.ps1')
            package.writestr('Install_RandomOffice.bat',
                '@echo off\r\n'
                'powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -Path \'%~dp0*\' | ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }"\r\n'
                'powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_desktop.ps1" -OpenDefaultApps\r\n'
                'if errorlevel 1 pause\r\n'
            )
        package.writestr('ONCE_OKUYUN.txt', instructions.encode('utf-8-sig'))
    with zipfile.ZipFile(staging) as package:
        if package.testzip() is not None:
            raise SystemExit('Archive verification failed')
    staging.replace(args.output)
    metadata = {'executable_sha256': hashlib.sha256(args.exe.read_bytes()).hexdigest(),
                'archive_sha256': hashlib.sha256(args.output.read_bytes()).hexdigest(),
                'executable_bytes': args.exe.stat().st_size, 'self_contained': True}
    args.output.with_suffix('.json').write_text(json.dumps(metadata, indent=2), encoding='utf-8')
    print(json.dumps(metadata, indent=2))


if __name__ == '__main__':
    main()
