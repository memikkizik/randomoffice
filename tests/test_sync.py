from pathlib import Path
import time
import zipfile

from randomoffice.sync.watcher import DirectorySyncManager, SyncEvent


def _create_dummy_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("[Content_Types].xml", b"<Types/>")


def test_directory_sync_detects_creation_modification_and_deletion(tmp_path: Path) -> None:
    events = []

    def handle_change(change):
        events.append(change)

    manager = DirectorySyncManager(
        watch_dir=tmp_path,
        on_change=handle_change,
        poll_interval=0.1,
    )

    # Initial scan
    manager.scan_once()

    # 1. Create file
    doc_path = tmp_path / "test.docx"
    _create_dummy_docx(doc_path)

    changes = manager.scan_once()
    assert len(changes) == 1
    assert changes[0].event_type == SyncEvent.CREATED
    assert changes[0].path == doc_path

    # 2. Modify file
    time.sleep(0.05)
    with zipfile.ZipFile(doc_path, "a") as zf:
        zf.writestr("word/document.xml", b"<w:document/>")

    changes2 = manager.scan_once()
    assert len(changes2) == 1
    assert changes2[0].event_type == SyncEvent.MODIFIED

    # 3. Delete file
    doc_path.unlink()
    changes3 = manager.scan_once()
    assert len(changes3) == 1
    assert changes3[0].event_type == SyncEvent.DELETED

