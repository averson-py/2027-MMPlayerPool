"""Build step for self-hosting: let the page draw under the phone's safe area (notch) and run full screen.

Streamlit's index.html has a viewport tag without viewport-fit=cover, so phones leave blank margins.
"""
from pathlib import Path

import streamlit

index = Path(streamlit.__file__).parent / "static" / "index.html"
html = index.read_text(encoding="utf-8")

if "viewport-fit=cover" in html:
    print("already patched:", index)
else:
    old_viewport = 'content="width=device-width, initial-scale=1, shrink-to-fit=no"'
    old_charset = '<meta charset="UTF-8" />'
    if old_viewport not in html or old_charset not in html:
        raise SystemExit("Streamlit's index.html changed; update patch_streamlit.py")
    html = html.replace(old_viewport, 'content="width=device-width, initial-scale=1, shrink-to-fit=no, viewport-fit=cover"', 1)
    html = html.replace(
        old_charset,
        old_charset
        + '\n    <meta name="apple-mobile-web-app-capable" content="yes" />'
        + '\n    <meta name="mobile-web-app-capable" content="yes" />'
        + '\n    <meta name="apple-mobile-web-app-status-bar-style" content="black" />',
        1,
    )
    index.write_text(html, encoding="utf-8")
    print("patched:", index)
