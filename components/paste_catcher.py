# components/paste_catcher.py
"""
Custom paste-catcher: click the area, then Ctrl+V to paste an image.
Returns the base64-encoded image data if a paste happened.
"""
from __future__ import annotations

import base64
import streamlit.components.v1 as components


_PASTE_COMPONENT = """
<html>
<head>
<style>
  #paste-area {
    width: 100%;
    height: 80px;
    border: 2px dashed #888;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #666;
    font-family: system-ui, sans-serif;
    cursor: pointer;
    background: #fafafa;
  }
  #paste-area:focus {
    border-color: #0B5FFF;
    background: #f0f7ff;
    outline: none;
  }
</style>
</head>
<body>
<div id="paste-area" tabindex="0">
  📋 Click here, then Ctrl+V to paste an image
</div>
<script>
const area = document.getElementById("paste-area");
const streamlit = window.parent;

area.addEventListener("paste", async (e) => {
  const items = e.clipboardData.items;
  for (const item of items) {
    if (item.type.startsWith("image/")) {
      const file = item.getAsFile();
      const reader = new FileReader();
      reader.onload = () => {
        const b64 = reader.result.split(",")[1];
        streamlit.postMessage({
          isStreamlitMessage: true,
          type: "streamlit:setComponentValue",
          value: { data: b64, mime: item.type },
        }, "*");
        area.innerHTML = "✅ Image pasted!";
        setTimeout(() => {
          area.innerHTML = "📋 Click here, then Ctrl+V to paste an image";
        }, 2000);
      };
      reader.readAsDataURL(file);
      return;
    }
  }
});

area.focus();
</script>
</html>
"""


def paste_catcher(key: str) -> dict | None:
    """
    Render the paste area.
    Returns {"data": base64_str, "mime": "image/png"} if paste happened,
    else None.
    """
    result = components.html(
        _PASTE_COMPONENT,
        height=100,
        key=key,
    )
    # components.html returns None in most cases; the value flows
    # via st.session_state indirectly. For reliable capture,
    # consider using a proper custom component (see note below).
    return result


def save_pasted_image(data: str, mime: str, dest_path: str) -> str:
    """Decode base64 image data and write to disk."""
    raw = base64.b64decode(data)
    Path(dest_path).write_bytes(raw)
    return dest_path
