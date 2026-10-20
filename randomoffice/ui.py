"""Load the shared UI with a single behavior source for web and desktop."""

from pathlib import Path
import base64
import re


def render_office_ui(html_path: Path) -> str:
    source = html_path.read_text(encoding="utf-8")
    for name,media in [('randomoffice-logo.svg','image/svg+xml'),('randomoffice-logo.png','image/png')]:
        asset=html_path.with_name(name)
        if asset.exists():
            source=source.replace('/static/'+name,'data:'+media+';base64,'+base64.b64encode(asset.read_bytes()).decode('ascii'))
    script = "\n".join(html_path.with_name(name).read_text(encoding="utf-8")
                       for name in ("formulas.js", "editor.js"))
    workspace = html_path.with_name("workspace.js")
    if workspace.exists():
        script += "\n" + workspace.read_text(encoding="utf-8")
    script_tag = '<script data-randomoffice-editor>\n' + script.replace('</script', '<\\/script') + '\n</script>'
    source, count = re.subn(r"<script>.*?</script>", lambda _: script_tag, source, count=1, flags=re.S)
    if not count:
        source = source.replace("</body>", script_tag + "</body>")
    stylesheet = html_path.with_name("editor.css")
    if stylesheet.exists():
        source = source.replace("</head>", "<style>" + stylesheet.read_text(encoding="utf-8") + "</style></head>")
    return source
