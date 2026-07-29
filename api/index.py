import os
import json
import base64
import zipfile
import shutil
import urllib.parse
import xml.etree.ElementTree as ET
from flask import Flask, request, jsonify, Response

app = Flask(__name__)

TMP_DIR = "/tmp" if os.path.exists("/tmp") else os.path.join(os.getcwd(), "temp_export")
os.makedirs(TMP_DIR, exist_ok=True)

active_presentation = {
    "name": "presentation",
    "html": None
}

@app.route('/api/upload', methods=['POST'])
@app.route('/upload', methods=['POST'])
def handle_upload():
    try:
        filename_header = request.headers.get('X-File-Name', 'presentation.pptx')
        filename = urllib.parse.unquote(filename_header)
        pptx_name = os.path.splitext(filename)[0]
        
        active_presentation["name"] = pptx_name

        file_bytes = request.get_data()
        if not file_bytes:
            return Response("No file data received.", status=400)

        pptx_path = os.path.join(TMP_DIR, "uploaded.pptx")
        with open(pptx_path, "wb") as f:
            f.write(file_bytes)

        unzip_dir = os.path.join(TMP_DIR, "unzipped")
        if os.path.exists(unzip_dir):
            shutil.rmtree(unzip_dir, ignore_errors=True)
            
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(unzip_dir)

        pres_xml_path = os.path.join(unzip_dir, "ppt", "presentation.xml")
        pres_rels_path = os.path.join(unzip_dir, "ppt", "_rels", "presentation.xml.rels")

        ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
        }

        pres_rels_tree = ET.parse(pres_rels_path)
        pres_rels = {}
        for rel in pres_rels_tree.getroot():
            rel_id = rel.attrib.get('Id')
            target = rel.attrib.get('Target')
            if rel_id and target:
                pres_rels[rel_id] = target

        pres_tree = ET.parse(pres_xml_path)
        ordered_slide_paths = []
        for sld_id in pres_tree.getroot().findall('.//p:sldId', ns):
            r_id = sld_id.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if r_id and r_id in pres_rels:
                target = pres_rels[r_id].replace('/', os.sep)
                ordered_slide_paths.append(os.path.join("ppt", target))

        if not ordered_slide_paths:
            slides_dir = os.path.join(unzip_dir, "ppt", "slides")
            if os.path.exists(slides_dir):
                for f in sorted(os.listdir(slides_dir)):
                    if f.startswith("slide") and f.endswith(".xml"):
                        ordered_slide_paths.append(os.path.join("ppt", "slides", f))

        slide_meta_list = []
        video_data_obj = {}

        for idx, slide_path_rel in enumerate(ordered_slide_paths):
            slide_xml_path = os.path.join(unzip_dir, slide_path_rel)
            slide_dir = os.path.dirname(slide_xml_path)
            slide_filename = os.path.basename(slide_xml_path)
            rels_path = os.path.join(slide_dir, "_rels", f"{slide_filename}.rels")

            has_video = False
            video_key = ""

            if os.path.exists(slide_xml_path) and os.path.exists(rels_path):
                rels_tree = ET.parse(rels_path)
                rel_map = {}
                for rel in rels_tree.getroot():
                    r_id = rel.attrib.get('Id')
                    target = rel.attrib.get('Target')
                    r_type = rel.attrib.get('Type', '')
                    if r_id and target:
                        rel_map[r_id] = {'target': target, 'type': r_type}

                video_r_ids = []
                for r_id, r_info in rel_map.items():
                    t_lower = r_info['target'].lower()
                    type_lower = r_info['type'].lower()
                    is_video = ('/video' in type_lower) or ('/media' in type_lower) or t_lower.endswith(('.mp4', '.webm', '.mov', '.avi', '.wmv'))
                    if is_video:
                        video_r_ids.append(r_id)

                if video_r_ids:
                    first_r_id = video_r_ids[0]
                    r_info = rel_map[first_r_id]
                    clean_target = r_info['target'].replace('../', 'ppt/').replace('/', os.sep)
                    src_video_path = os.path.join(unzip_dir, clean_target)

                    if os.path.exists(src_video_path):
                        has_video = True
                        video_key = f"video_slide_{idx}"
                        with open(src_video_path, 'rb') as vf:
                            vbytes = vf.read()
                            video_data_obj[video_key] = base64.b64encode(vbytes).decode('utf-8')

            slide_meta_list.append({
                "index": idx,
                "hasVideo": has_video,
                "stepsCount": 1,
                "bg": "",
                "stepBgs": [""],
                "videoKey": video_key,
                "left": 0,
                "top": 0,
                "width": 100,
                "height": 100
            })

        html_content = generate_standalone_html(pptx_name, len(ordered_slide_paths), slide_meta_list, video_data_obj)
        active_presentation["html"] = html_content

        html_out_path = os.path.join(TMP_DIR, "index.html")
        with open(html_out_path, "w", encoding="utf-8") as f:
            f.write(html_content)

        return jsonify({
            "status": "success",
            "redirect": "/api/download"
        })

    except Exception as e:
        return Response(f"Vercel Serverless Conversion Error: {str(e)}", status=500, mimetype="text/plain")

@app.route('/api/download', methods=['GET'])
@app.route('/download', methods=['GET'])
def handle_download():
    html_content = active_presentation.get("html")
    if not html_content:
        html_path = os.path.join(TMP_DIR, "index.html")
        if os.path.exists(html_path):
            with open(html_path, "r", encoding="utf-8") as f:
                html_content = f.read()

    if not html_content:
        return Response("No converted presentation found. Please convert a presentation first.", status=404)

    filename = f"{active_presentation['name']}.html"
    safe_filename = urllib.parse.quote(filename)

    return Response(
        html_content,
        mimetype="text/html; charset=utf-8",
        headers={
            "Content-Disposition": f'attachment; filename="{safe_filename}"; filename*=UTF-8\'\'{safe_filename}'
        }
    )

def generate_standalone_html(pptx_name, num_slides, slide_meta_list, video_data_obj):
    thumbnails = []
    for idx in range(num_slides):
        slide_num = idx + 1
        thumb_html = f'''      <button class="thumb-btn" data-index="{idx}" onclick="goToSlide({idx}, false)">
        <span class="thumb-number">{slide_num}</span>
        <div class="thumb-preview" style="background:#111827;"></div>
      </button>'''
        thumbnails.append(thumb_html)
    thumbnails_joined = "\n".join(thumbnails)

    html_template = '''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>__PPTX_NAME__ - Presentation Deck</title>
  
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=Plus+Jakarta+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">

  <style>
    :root {
      --font-primary: 'Plus Jakarta Sans', system-ui, -apple-system, sans-serif;
      --font-display: 'Outfit', system-ui, -apple-system, sans-serif;
      --bg-app: #0b0f19;
      --bg-card: rgba(17, 24, 39, 0.7);
      --bg-sidebar: #111827;
      --border-color: rgba(255, 255, 255, 0.08);
      --text-main: #f3f4f6;
      --text-muted: #9ca3af;
      --accent: #3b82f6;
      --accent-hover: #60a5fa;
      --accent-gradient: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%);
      --active-thumb-bg: rgba(59, 130, 246, 0.15);
      --active-thumb-border: #3b82f6;
      --shadow-lg: 0 10px 25px -5px rgba(0, 0, 0, 0.3), 0 8px 10px -6px rgba(0, 0, 0, 0.3);
      --glass-bg: rgba(17, 24, 39, 0.6);
      --glass-blur: 16px;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: var(--font-primary);
      background-color: var(--bg-app);
      color: var(--text-main);
      height: 100vh;
      overflow: hidden;
      display: flex;
      flex-direction: column;
    }

    header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 0.75rem 1.5rem;
      background-color: var(--bg-sidebar);
      border-bottom: 1px solid var(--border-color);
      z-index: 100;
      height: 64px;
    }

    .brand { display: flex; align-items: center; gap: 0.75rem; }
    .brand-logo {
      width: 32px; height: 32px;
      background: var(--accent-gradient);
      border-radius: 8px;
      display: flex; align-items: center; justify-content: center;
      color: white; font-family: var(--font-display); font-weight: 700;
    }

    .brand h1 {
      font-family: var(--font-display);
      font-size: 1.1rem; font-weight: 600;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 300px;
    }

    .header-controls { display: flex; align-items: center; gap: 0.75rem; }

    .btn {
      background: transparent; border: 1px solid var(--border-color);
      color: var(--text-main); padding: 0.5rem 0.85rem; border-radius: 8px;
      cursor: pointer; font-family: var(--font-primary); font-size: 0.875rem;
      display: flex; align-items: center; gap: 0.5rem; text-decoration: none;
    }

    .btn-primary { background: var(--accent-gradient); color: white; border: none; }

    .app-container { display: flex; flex: 1; position: relative; height: calc(100vh - 64px); }

    .sidebar {
      width: 280px; background-color: var(--bg-sidebar);
      border-right: 1px solid var(--border-color);
      display: flex; flex-direction: column; z-index: 90; overflow-y: auto;
    }

    .sidebar-header { padding: 1rem 1.25rem; border-bottom: 1px solid var(--border-color); font-size: 0.9rem; color: var(--text-muted); font-weight: 600; }
    .thumb-list { padding: 0.75rem; display: flex; flex-direction: column; gap: 0.5rem; }

    .thumb-btn {
      display: flex; align-items: center; gap: 0.75rem; padding: 0.5rem;
      background: transparent; border: 1px solid transparent; border-radius: 8px;
      color: var(--text-muted); text-align: left; cursor: pointer;
    }

    .thumb-btn.active { background-color: var(--active-thumb-bg); border-color: var(--active-thumb-border); color: var(--accent); }
    .thumb-number { width: 24px; height: 24px; border-radius: 6px; background-color: var(--border-color); font-size: 0.75rem; display: flex; align-items: center; justify-content: center; }

    .viewer-panel {
      flex: 1; display: flex; flex-direction: column; position: relative;
      background-color: var(--bg-app); overflow: hidden; align-items: center; justify-content: center; padding: 2rem;
    }

    .slide-arena { width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; position: relative; }

    #interactive-slide-container {
      width: 960px; height: 540px; position: relative; overflow: hidden;
      box-shadow: 0 10px 25px rgba(0,0,0,0.3); border-radius: 8px; background-color: #000;
    }

    #interactive-slide-bg { width: 100%; height: 100%; position: relative; cursor: pointer; }

    .player-controls {
      position: absolute; bottom: 2rem; display: flex; align-items: center; gap: 1rem;
      background: var(--glass-bg); backdrop-filter: blur(16px);
      border: 1px solid var(--border-color); padding: 0.6rem 1.2rem; border-radius: 30px; z-index: 80;
    }

    .slide-indicator { font-family: var(--font-display); font-size: 0.9rem; color: var(--text-main); min-width: 90px; text-align: center; }
    .ctrl-btn { background: transparent; border: none; color: var(--text-main); width: 36px; height: 36px; border-radius: 50%; cursor: pointer; display: flex; align-items: center; justify-content: center; }
    .ctrl-btn svg { width: 20px; height: 20px; fill: currentColor; }
  </style>
</head>
<body>

  <header>
    <div class="brand">
      <div class="brand-logo">P</div>
      <h1 title="__PPTX_NAME__">__PPTX_NAME__</h1>
    </div>
    <div class="header-controls">
      <a href="/api/download" class="btn btn-primary" download>
        <span>Download HTML Deck</span>
      </a>
    </div>
  </header>

  <div class="app-container">
    <aside class="sidebar">
      <div class="sidebar-header">Slides</div>
      <div class="thumb-list">__THUMBNAILS__</div>
    </aside>

    <main class="viewer-panel">
      <div class="slide-arena">
        <div id="interactive-slide-container">
          <div id="interactive-slide-bg"></div>
        </div>
      </div>

      <div class="player-controls">
        <button class="ctrl-btn" onclick="prevSlideDirect()"><svg viewBox="0 0 24 24"><path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg></button>
        <span class="slide-indicator" id="slide-indicator">1 / 1</span>
        <button class="ctrl-btn" onclick="nextSlideDirect()"><svg viewBox="0 0 24 24"><path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/></svg></button>
      </div>
    </main>
  </div>

  <script>
    let currentSlide = 0;
    const totalSlides = __NUM_SLIDES__;
    const slideMeta = __SLIDE_META_JSON__;
    const videoData = __VIDEO_DATA_JSON__;
    const videoUrls = {};

    const slideIndicator = document.getElementById('slide-indicator');
    const interBg = document.getElementById('interactive-slide-bg');

    function base64ToBlob(base64, mimeType) {
      const byteCharacters = atob(base64);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) byteNumbers[i] = byteCharacters.charCodeAt(i);
      return new Blob([new Uint8Array(byteNumbers)], { type: mimeType });
    }

    function initVideoBlobUrls() {
      for (const key in videoData) {
        if (videoData[key]) {
          try {
            const blob = base64ToBlob(videoData[key], 'video/mp4');
            videoUrls[key] = URL.createObjectURL(blob);
          } catch (e) { console.error(e); }
        }
      }
    }

    function goToSlide(index) {
      if (index < 0 || index >= totalSlides) return;
      currentSlide = index;
      slideIndicator.textContent = `${currentSlide + 1} / ${totalSlides}`;

      document.querySelectorAll('.thumb-btn').forEach(btn => btn.classList.remove('active'));
      const activeThumb = document.querySelector(`.thumb-btn[data-index="${index}"]`);
      if (activeThumb) activeThumb.classList.add('active');

      renderSlide();
    }

    function renderSlide() {
      const meta = slideMeta[currentSlide];
      interBg.innerHTML = '';
      interBg.style.background = '#000';

      if (meta.hasVideo) {
        const assetUrl = videoUrls[meta.videoKey];
        interBg.innerHTML = `<video id="slide-interactive-video" src="${assetUrl}" controls style="position: absolute; left: 0; top: 0; width: 100%; height: 100%; object-fit: contain; background: #000; z-index: 10;"></video>`;
      } else {
        interBg.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;font-size:1.5rem;font-family:sans-serif;">Slide ${currentSlide + 1}</div>`;
      }
    }

    function nextSlideDirect() { goToSlide(currentSlide + 1); }
    function prevSlideDirect() { goToSlide(currentSlide - 1); }

    window.onload = function() {
      initVideoBlobUrls();
      goToSlide(0);
      document.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowRight' || e.key === ' ') { nextSlideDirect(); }
        else if (e.key === 'ArrowLeft') { prevSlideDirect(); }
      });
    };
  </script>
</body>
</html>'''

    html_content = html_template.replace("__PPTX_NAME__", pptx_name)
    html_content = html_content.replace("__NUM_SLIDES__", str(num_slides))
    html_content = html_content.replace("__THUMBNAILS__", thumbnails_joined)
    html_content = html_content.replace("__SLIDE_META_JSON__", json.dumps(slide_meta_list))
    html_content = html_content.replace("__VIDEO_DATA_JSON__", json.dumps(video_data_obj))

    return html_content

if __name__ == "__main__":
    app.run(port=8095)
