# PowerPoint to Standalone HTML Converter - Python 3.12 Server
# Operating System: Windows

import os
import sys
import json
import base64
import zipfile
import subprocess
import shutil
import urllib.parse
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
import xml.etree.ElementTree as ET

PORT = 8095
WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(WORKSPACE_DIR, "html_output")
SLIDES_DIR = os.path.join(OUTPUT_DIR, "slides")
MEDIA_DIR = os.path.join(OUTPUT_DIR, "media")
TEMP_UNZIP_DIR = os.path.join(WORKSPACE_DIR, "temp_pptx_unzip")

active_presentation_name = "presentation"

def log_info(msg):
    print(f"[INFO] {msg}", flush=True)

def log_success(msg):
    print(f"[SUCCESS] {msg}", flush=True)

def log_error(msg):
    print(f"[ERROR] {msg}", flush=True)

def cleanup_orphan_powerpoint():
    """Kill any hanging orphan POWERPNT processes on Windows."""
    try:
        subprocess.run(
            ["powershell", "-Command", "Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"],
            capture_output=True,
            text=True
        )
    except Exception as e:
        log_error(f"Failed to cleanup POWERPNT processes: {e}")

def convert_pptx_to_html(pptx_file_path, pptx_name):
    global active_presentation_name
    active_presentation_name = pptx_name
    log_info(f"Starting Python presentation conversion for: {pptx_name}...")

    # 1. Recreate output folders
    if os.path.exists(SLIDES_DIR):
        shutil.rmtree(SLIDES_DIR, ignore_errors=True)
    if os.path.exists(MEDIA_DIR):
        shutil.rmtree(MEDIA_DIR, ignore_errors=True)
    
    os.makedirs(SLIDES_DIR, exist_ok=True)
    os.makedirs(MEDIA_DIR, exist_ok=True)

    # 2. Invoke PowerPoint COM engine via PowerShell helper to export progressive step frames
    cleanup_orphan_powerpoint()

    # Inline PowerShell snippet to process PowerPoint slides
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$pptxFilePath = '{pptx_file_path.replace("\\", "\\\\")}'
$slidesFolder = '{SLIDES_DIR.replace("\\", "\\\\")}'

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = 1
$ppt.WindowState = 2

$pres = $ppt.Presentations.Open($pptxFilePath, [Microsoft.Office.Core.MsoTristate]::msoFalse, [Microsoft.Office.Core.MsoTristate]::msoFalse, [Microsoft.Office.Core.MsoTristate]::msoTrue)
$totalSlides = $pres.Slides.Count

$origWidth = $pres.PageSetup.SlideWidth
$origHeight = $pres.PageSetup.SlideHeight
$exportWidth = [int]($origWidth * 2.0)
$exportHeight = [int]($origHeight * 2.0)

$slideMetaList = @()

for ($i = 1; $i -le $totalSlides; $i++) {{
    $slide = $pres.Slides.Item($i)
    $effects = $slide.TimeLine.MainSequence
    $clickIndex = 0
    $effectMap = @{{}}
    
    if ($i -ne 3) {{
        foreach ($effect in $effects) {{
            if ($effect.Timing.TriggerType -eq 1) {{ $clickIndex++ }}
            if ($effect.Exit -eq $false) {{
                $shapeId = $effect.Shape.Id
                if (!$effectMap.ContainsKey($shapeId) -or $effectMap[$shapeId] -gt $clickIndex) {{
                    $effectMap[$shapeId] = $clickIndex
                }}
            }}
        }}
    }}

    $hasVideo = $false
    foreach ($shape in $slide.Shapes) {{
        if ($shape.Type -eq 16 -and ($shape.MediaType -eq 1 -or $shape.MediaType -eq 3)) {{
            $hasVideo = $true
            break
        }}
    }}

    $stepBgs = @()
    if ($clickIndex -gt 0) {{
        for ($k = 0; $k -le $clickIndex; $k++) {{
            $hiddenShapes = @()
            foreach ($shape in $slide.Shapes) {{
                $shapeId = $shape.Id
                if ($effectMap.ContainsKey($shapeId) -and $effectMap[$shapeId] -gt $k) {{
                    $originalVisibility = $shape.Visible
                    if ($originalVisibility -eq -1) {{
                        $shape.Visible = 0
                        $hiddenShapes += @{{ shape = $shape; original = $originalVisibility }}
                    }}
                }}
            }}
            
            $outputPath = Join-Path $slidesFolder "slide_${{i}}_step_${{k}}.png"
            $slide.Export($outputPath, "PNG", $exportWidth, $exportHeight)
            
            foreach ($item in $hiddenShapes) {{
                $item.shape.Visible = $item.original
            }}
            
            $pngBytes = [System.IO.File]::ReadAllBytes($outputPath)
            $stepBgs += [Convert]::ToBase64String($pngBytes)
            Remove-Item $outputPath -Force
        }}
    }} else {{
        $outputPath = Join-Path $slidesFolder "slide_${{i}}_step_0.png"
        $slide.Export($outputPath, "PNG", $exportWidth, $exportHeight)
        $pngBytes = [System.IO.File]::ReadAllBytes($outputPath)
        $stepBgs += [Convert]::ToBase64String($pngBytes)
        Remove-Item $outputPath -Force
    }}

    $slideMetaList += @{{
        index = $i - 1
        hasVideo = $hasVideo
        stepsCount = $stepBgs.Count
        bg = $stepBgs[0]
        stepBgs = $stepBgs
        videoKey = if ($hasVideo) {{ "video_slide_$($i - 1)" }} else {{ "" }}
        left = 0
        top = 0
        width = 100
        height = 100
    }}
}}

$pres.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()

$slideMetaList | ConvertTo-Json -Depth 4 -Compress
"""

    log_info("Exporting progressive slide layouts via PowerPoint COM bridge...")
    ps_cmd = ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script]
    res = subprocess.run(ps_cmd, capture_output=True, text=True)
    
    if res.returncode != 0:
        log_error(f"PowerPoint COM execution failed: {res.stderr}")
        raise RuntimeError(f"PowerPoint conversion failed: {res.stderr}")
    
    slide_meta_list = json.loads(res.stdout.strip())
    log_success("Slide layouts and animation frames extracted successfully.")

    # 3. Unzip PPTX and parse relationships / embedded media using Python zipfile & ElementTree
    log_info("Parsing presentation package and extracting video assets...")
    if os.path.exists(TEMP_UNZIP_DIR):
        shutil.rmtree(TEMP_UNZIP_DIR, ignore_errors=True)
        
    with zipfile.ZipFile(pptx_file_path, 'r') as zip_ref:
        zip_ref.extractall(TEMP_UNZIP_DIR)

    pres_xml_path = os.path.join(TEMP_UNZIP_DIR, "ppt", "presentation.xml")
    pres_rels_path = os.path.join(TEMP_UNZIP_DIR, "ppt", "_rels", "presentation.xml.rels")

    ns = {
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'pr': 'http://schemas.openxmlformats.org/package/2006/relationships'
    }

    # Parse presentation rels
    pres_rels_tree = ET.parse(pres_rels_path)
    pres_rels = {}
    for rel in pres_rels_tree.getroot():
        rel_id = rel.attrib.get('Id')
        target = rel.attrib.get('Target')
        if rel_id and target:
            pres_rels[rel_id] = target

    # Parse slide order
    pres_tree = ET.parse(pres_xml_path)
    ordered_slide_paths = []
    for sld_id in pres_tree.getroot().findall('.//p:sldId', ns):
        r_id = sld_id.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        if r_id and r_id in pres_rels:
            target = pres_rels[r_id].replace('/', os.sep)
            ordered_slide_paths.append(os.path.join("ppt", target))

    video_data_obj = {}

    # Scan slides for embedded video files
    for idx, slide_path_rel in enumerate(ordered_slide_paths):
        slide_xml_path = os.path.join(TEMP_UNZIP_DIR, slide_path_rel)
        slide_dir = os.path.dirname(slide_xml_path)
        slide_filename = os.path.basename(slide_xml_path)
        rels_path = os.path.join(slide_dir, "_rels", f"{slide_filename}.rels")

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
                log_info(f"Extracting embedded video file for Slide {idx + 1}...")
                first_r_id = video_r_ids[0]
                r_info = rel_map[first_r_id]
                clean_target = r_info['target'].replace('../', 'ppt/').replace('/', os.sep)
                src_video_path = os.path.join(TEMP_UNZIP_DIR, clean_target)

                if os.path.exists(src_video_path):
                    video_key = f"video_slide_{idx}"
                    with open(src_video_path, 'rb') as vf:
                        video_bytes = vf.read()
                        video_data_obj[video_key] = base64.b64encode(video_bytes).decode('utf-8')

                    # Update slide meta
                    slide_meta_list[idx]['hasVideo'] = True
                    slide_meta_list[idx]['videoKey'] = video_key

    # 4. Generate sidebar thumbnails
    thumbnails = []
    for idx in range(len(ordered_slide_paths)):
        slide_num = idx + 1
        slide_bg = slide_meta_list[idx]['bg']
        thumb_html = f'''      <button class="thumb-btn" data-index="{idx}" onclick="goToSlide({idx}, false)">
        <span class="thumb-number">{slide_num}</span>
        <div class="thumb-preview" style="background-image: url('data:image/png;base64,{slide_bg}');"></div>
      </button>'''
        thumbnails.append(thumb_html)
    thumbnails_joined = "\n".join(thumbnails)

    # 5. Build HTML Presentation Template
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

    [data-theme="light"] {
      --bg-app: #f8fafc;
      --bg-card: rgba(255, 255, 255, 0.7);
      --bg-sidebar: #ffffff;
      --border-color: rgba(0, 0, 0, 0.06);
      --text-main: #0f172a;
      --text-muted: #475569;
      --accent: #2563eb;
      --accent-hover: #1d4ed8;
      --accent-gradient: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
      --active-thumb-bg: rgba(37, 99, 235, 0.08);
      --active-thumb-border: #2563eb;
      --shadow-lg: 0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.05);
      --glass-bg: rgba(255, 255, 255, 0.6);
      --glass-blur: 16px;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      font-family: var(--font-primary);
      background-color: var(--bg-app);
      color: var(--text-main);
      height: 100vh;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      transition: background-color 0.3s ease, color 0.3s ease;
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
      transition: background-color 0.3s ease, border-color 0.3s ease;
    }

    .brand {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    .brand-logo {
      width: 32px;
      height: 32px;
      background: var(--accent-gradient);
      border-radius: 8px;
      display: flex;
      align-items: center;
      justify-content: center;
      color: white;
      font-family: var(--font-display);
      font-weight: 700;
      font-size: 1.1rem;
      box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
    }

    .brand h1 {
      font-family: var(--font-display);
      font-size: 1.1rem;
      font-weight: 600;
      letter-spacing: -0.02em;
      white-space: nowrap;
      text-overflow: ellipsis;
      overflow: hidden;
      max-width: 300px;
    }

    .header-controls {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }

    .btn {
      background: transparent;
      border: 1px solid var(--border-color);
      color: var(--text-main);
      padding: 0.5rem 0.85rem;
      border-radius: 8px;
      cursor: pointer;
      font-family: var(--font-primary);
      font-size: 0.875rem;
      font-weight: 500;
      display: flex;
      align-items: center;
      gap: 0.5rem;
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
      text-decoration: none;
    }

    .btn:hover {
      background-color: var(--border-color);
      border-color: var(--text-muted);
      transform: translateY(-1px);
    }

    .btn-primary {
      background: var(--accent-gradient);
      color: white;
      border: none;
    }

    .btn-primary:hover {
      box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
      filter: brightness(1.1);
    }

    .app-container {
      display: flex;
      flex: 1;
      position: relative;
      height: calc(100vh - 64px);
    }

    .sidebar {
      width: 280px;
      background-color: var(--bg-sidebar);
      border-right: 1px solid var(--border-color);
      display: flex;
      flex-direction: column;
      transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1), width 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      z-index: 90;
      overflow-y: auto;
    }

    .sidebar.collapsed {
      width: 0;
      transform: translateX(-100%);
      border-right-width: 0;
    }

    .sidebar-header {
      padding: 1rem 1.25rem;
      border-bottom: 1px solid var(--border-color);
      font-family: var(--font-display);
      font-size: 0.9rem;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      color: var(--text-muted);
      font-weight: 600;
    }

    .thumb-list {
      padding: 0.75rem;
      display: flex;
      flex-direction: column;
      gap: 0.5rem;
    }

    .thumb-btn {
      display: flex;
      align-items: center;
      gap: 0.75rem;
      padding: 0.5rem;
      background: transparent;
      border: 1px solid transparent;
      border-radius: 8px;
      color: var(--text-muted);
      text-align: left;
      cursor: pointer;
      transition: all 0.2s ease;
    }

    .thumb-btn:hover {
      background-color: var(--border-color);
      color: var(--text-main);
    }

    .thumb-btn.active {
      background-color: var(--active-thumb-bg);
      border-color: var(--active-thumb-border);
      color: var(--accent);
    }

    .thumb-number {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 24px;
      height: 24px;
      border-radius: 6px;
      background-color: var(--border-color);
      font-size: 0.75rem;
      font-weight: 600;
      color: var(--text-main);
    }

    .thumb-btn.active .thumb-number {
      background-color: var(--accent);
      color: white;
    }

    .thumb-preview {
      flex: 1;
      height: 50px;
      border-radius: 4px;
      background-size: cover;
      background-position: center;
      border: 1px solid var(--border-color);
    }

    .viewer-panel {
      flex: 1;
      display: flex;
      flex-direction: column;
      position: relative;
      background-color: var(--bg-app);
      overflow: hidden;
      align-items: center;
      justify-content: center;
      padding: 2rem;
    }

    .slide-arena {
      width: 100%;
      height: 100%;
      display: flex;
      align-items: center;
      justify-content: center;
      position: relative;
    }

    #interactive-slide-container {
      width: 960px;
      height: 540px;
      position: relative;
      overflow: hidden;
      box-shadow: 0 10px 25px rgba(0,0,0,0.3);
      border-radius: 8px;
      background-color: #000;
      transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    #interactive-slide-bg {
      width: 100%;
      height: 100%;
      background-size: cover;
      background-position: center;
      position: relative;
      cursor: pointer;
    }

    .player-controls {
      position: absolute;
      bottom: 2rem;
      display: flex;
      align-items: center;
      gap: 1rem;
      background: var(--glass-bg);
      backdrop-filter: blur(var(--glass-blur));
      -webkit-backdrop-filter: blur(var(--glass-blur));
      border: 1px solid var(--border-color);
      padding: 0.6rem 1.2rem;
      border-radius: 30px;
      box-shadow: var(--shadow-lg);
      z-index: 80;
    }

    .slide-indicator {
      font-family: var(--font-display);
      font-size: 0.9rem;
      font-weight: 500;
      color: var(--text-main);
      min-width: 90px;
      text-align: center;
    }

    .ctrl-btn {
      background: transparent;
      border: none;
      color: var(--text-main);
      width: 36px;
      height: 36px;
      border-radius: 50%;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      transition: all 0.2s ease;
    }

    .ctrl-btn:hover {
      background-color: var(--border-color);
      color: var(--accent);
    }

    .ctrl-btn:disabled {
      color: var(--text-muted);
      opacity: 0.4;
      cursor: not-allowed;
    }

    .ctrl-btn svg {
      width: 20px;
      height: 20px;
      fill: currentColor;
    }

    .progress-bar-container {
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      height: 4px;
      background-color: var(--border-color);
      z-index: 85;
    }

    .progress-bar {
      height: 100%;
      width: 0%;
      background: var(--accent-gradient);
      transition: width 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    .toast {
      position: absolute;
      bottom: 6rem;
      background-color: rgba(17, 24, 39, 0.9);
      color: white;
      padding: 0.5rem 1rem;
      border-radius: 8px;
      font-size: 0.875rem;
      pointer-events: none;
      opacity: 0;
      transform: translateY(10px);
      transition: all 0.2s ease;
      z-index: 100;
    }

    .toast.show {
      opacity: 1;
      transform: translateY(0);
    }

    :fullscreen .viewer-panel { padding: 0; background-color: #000; }
    :fullscreen #interactive-slide-container { border-radius: 0; box-shadow: none; }
  </style>
</head>
<body>

  <header>
    <div class="brand">
      <div class="brand-logo">P</div>
      <h1 title="__PPTX_NAME__">__PPTX_NAME__</h1>
    </div>
    <div class="header-controls">
      <a href="/download" class="btn btn-primary" download title="Download standalone HTML file">
        <svg style="width:18px; height:18px; fill:currentColor" viewBox="0 0 24 24"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
        <span>Download HTML File</span>
      </a>
      <button class="btn" id="theme-toggle" onclick="toggleTheme()">
        <svg id="theme-sun" style="display:none; width:18px; height:18px; fill:currentColor" viewBox="0 0 20 20"><path d="M10 2a1 1 0 011 1v1a1 1 0 11-2 0V3a1 1 0 011-1zm4 8a4 4 0 11-8 0 4 4 0 018 0zm-.464 4.95l.707.707a1 1 0 001.414-1.414l-.707-.707a1 1 0 00-1.414 1.414zm2.12-10.607a1 1 0 010 1.414l-.706.707a1 1 0 11-1.414-1.414l.707-.707a1 1 0 011.414 0zM17 11a1 1 0 100-2h-1a1 1 0 100 2h1zm-7 4a1 1 0 011 1v1a1 1 0 11-2 0v-1a1 1 0 011-1zM5.05 6.464A1 1 0 106.465 5.05l-.708-.707a1 1 0 00-1.414 1.414l.707.707zm1.414 8.486l-.707.707a1 1 0 01-1.414-1.414l.707-.707a1 1 0 011.414 1.414zM4 11a1 1 0 100-2H3a1 1 0 100 2h1z"/></svg>
        <svg id="theme-moon" style="width:18px; height:18px; fill:currentColor" viewBox="0 0 20 20"><path d="M17.293 13.293A8 8 0 016.707 2.707a8.001 8.001 0 1010.586 10.586z"/></svg>
        <span>Theme</span>
      </button>
      <button class="btn" onclick="toggleSidebar()">
        <span>Sidebar</span>
      </button>
      <button class="btn" onclick="toggleFullscreen()">
        <span>Fullscreen</span>
      </button>
    </div>
  </header>

  <div class="app-container">
    <div class="progress-bar-container">
      <div class="progress-bar" id="progress-bar"></div>
    </div>

    <aside class="sidebar" id="sidebar">
      <div class="sidebar-header">Slides</div>
      <div class="thumb-list">
        __THUMBNAILS__
      </div>
    </aside>

    <main class="viewer-panel" id="viewer-panel">
      <div class="slide-arena" id="slide-arena">
        <!-- Main Slide Container -->
        <div id="interactive-slide-container">
          <div id="interactive-slide-bg"></div>
        </div>
      </div>
      <div class="toast" id="toast"></div>

      <div class="player-controls">
        <button class="ctrl-btn" id="prev-btn" onclick="prevSlideDirect()">
          <svg viewBox="0 0 24 24"><path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg>
        </button>
        <span class="slide-indicator" id="slide-indicator">1 / 1</span>
        <button class="ctrl-btn" id="next-btn" onclick="nextSlideDirect()">
          <svg viewBox="0 0 24 24"><path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/></svg>
        </button>
      </div>
    </main>
  </div>

  <script>
    let currentSlide = 0;
    let currentStep = 0;
    const totalSlides = __NUM_SLIDES__;
    const slideWidth = 960;
    const slideHeight = 540;

    const slideMeta = __SLIDE_META_JSON__;
    const videoData = __VIDEO_DATA_JSON__;
    const videoUrls = {};

    const sidebar = document.getElementById('sidebar');
    const slideIndicator = document.getElementById('slide-indicator');
    const progressBar = document.getElementById('progress-bar');
    const prevBtn = document.getElementById('prev-btn');
    const nextBtn = document.getElementById('next-btn');
    const toast = document.getElementById('toast');
    const themeSun = document.getElementById('theme-sun');
    const themeMoon = document.getElementById('theme-moon');
    const interContainer = document.getElementById('interactive-slide-container');
    const interBg = document.getElementById('interactive-slide-bg');

    function base64ToBlob(base64, mimeType) {
      const byteCharacters = atob(base64);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      return new Blob([byteArray], { type: mimeType });
    }

    function initVideoBlobUrls() {
      for (const key in videoData) {
        if (videoData.hasOwnProperty(key) && videoData[key]) {
          try {
            const blob = base64ToBlob(videoData[key], 'video/mp4');
            videoUrls[key] = URL.createObjectURL(blob);
          } catch (e) {
            console.error("Failed to decode slide video asset: " + key, e);
          }
        }
      }
    }

    function initViewer() {
      initVideoBlobUrls();
      goToSlide(0, false);
      setupResizeHandler();
      setupKeyboardControls();
      
      interBg.addEventListener('click', (e) => {
        if (e.target.tagName.toLowerCase() === 'video' || e.target.tagName.toLowerCase() === 'button') {
          return;
        }
        triggerAdvance();
      });

      if (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) {
        setTheme('light');
      }
    }

    function goToSlide(index, startAtLastStep = false) {
      if (index < 0 || index >= totalSlides) return;
      
      const activeSlideVideo = document.getElementById('slide-interactive-video');
      if (activeSlideVideo) activeSlideVideo.pause();

      currentSlide = index;
      const meta = slideMeta[currentSlide];
      
      currentStep = startAtLastStep ? (meta.stepsCount - 1) : 0;

      slideIndicator.textContent = `${currentSlide + 1} / ${totalSlides}`;
      progressBar.style.width = (((currentSlide + 1) / totalSlides) * 100) + '%';
      
      prevBtn.disabled = currentSlide === 0;
      nextBtn.disabled = currentSlide === totalSlides - 1;

      document.querySelectorAll('.thumb-btn').forEach(btn => btn.classList.remove('active'));
      const activeThumb = document.querySelector(`.thumb-btn[data-index="${index}"]`);
      if (activeThumb) {
        activeThumb.classList.add('active');
        activeThumb.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      }

      renderCurrentStep();
    }

    function renderCurrentStep() {
      const meta = slideMeta[currentSlide];
      
      interBg.style.backgroundImage = `url("data:image/png;base64,${meta.stepBgs[currentStep]}")`;
      interBg.innerHTML = '';

      if (meta.hasVideo) {
        const assetUrl = videoUrls[meta.videoKey];
        interBg.style.backgroundImage = '';
        interBg.innerHTML = `
          <video id="slide-interactive-video" src="${assetUrl}" controls style="position: absolute; left: 0; top: 0; width: 100%; height: 100%; object-fit: contain; background: #000; z-index: 10;"></video>
        `;
      }
    }

    function nextStep() {
      const meta = slideMeta[currentSlide];
      if (currentStep < meta.stepsCount - 1) {
        currentStep++;
        renderCurrentStep();
      } else {
        goToSlide(currentSlide + 1, false);
      }
    }

    function prevStep() {
      if (currentStep > 0) {
        currentStep--;
        renderCurrentStep();
      } else if (currentSlide > 0) {
        goToSlide(currentSlide - 1, true);
      }
    }

    function nextSlideDirect() {
      goToSlide(currentSlide + 1, false);
    }

    function prevSlideDirect() {
      goToSlide(currentSlide - 1, false);
    }

    function triggerAdvance() {
      const meta = slideMeta[currentSlide];
      
      if (meta.hasVideo) {
        const slideVideo = document.getElementById('slide-interactive-video');
        if (slideVideo) {
          if (slideVideo.paused) {
            slideVideo.play();
            showToast("Playing video");
          } else {
            slideVideo.pause();
            showToast("Video paused");
          }
        }
      } else {
        nextStep();
      }
    }

    function scaleSlideToFit() {
      const arena = document.getElementById('slide-arena');
      if (!arena || !interContainer) return;

      const containerWidth = arena.clientWidth;
      const containerHeight = arena.clientHeight;
      const scale = Math.min(containerWidth / slideWidth, containerHeight / slideHeight) * 0.95;
      
      interContainer.style.transform = `scale(${scale})`;
    }

    function setupResizeHandler() {
      window.addEventListener('resize', scaleSlideToFit);
      if (window.ResizeObserver) {
        new ResizeObserver(scaleSlideToFit).observe(document.getElementById('slide-arena'));
      }
    }

    function toggleSidebar() {
      sidebar.classList.toggle('collapsed');
      setTimeout(scaleSlideToFit, 300);
    }

    function setupKeyboardControls() {
      document.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowRight' || e.key === 'PageDown') { e.preventDefault(); nextSlideDirect(); }
        else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { e.preventDefault(); prevSlideDirect(); }
        else if (e.key === ' ') { e.preventDefault(); triggerAdvance(); }
        else if (e.key === 'Backspace') { e.preventDefault(); prevStep(); }
        else if (e.key === 'Home') { e.preventDefault(); goToSlide(0, false); }
        else if (e.key === 'End') { e.preventDefault(); goToSlide(totalSlides - 1, false); }
        else if (e.key === 'f' || e.key === 'F') { toggleFullscreen(); }
        else if (e.key === 's' || e.key === 'S') { toggleSidebar(); }
      });
    }

    function toggleFullscreen() {
      if (!document.fullscreenElement) {
        interContainer.requestFullscreen().then(() => showToast("Entered Fullscreen")).catch(err => console.error(err));
      } else {
        document.exitFullscreen();
        showToast("Exited Fullscreen");
      }
      setTimeout(scaleSlideToFit, 100);
    }

    function toggleTheme() {
      const currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
      setTheme(currentTheme === 'dark' ? 'light' : 'dark');
    }

    function setTheme(theme) {
      document.documentElement.setAttribute('data-theme', theme);
      if (theme === 'light') {
        themeSun.style.display = 'none';
        themeMoon.style.display = 'block';
      } else {
        themeSun.style.display = 'block';
        themeMoon.style.display = 'none';
      }
    }

    function showToast(msg) {
      toast.textContent = msg;
      toast.classList.add('show');
      setTimeout(() => toast.classList.remove('show'), 2000);
    }

    window.onload = initViewer;
  </script>
</body>
</html>'''

    html_content = html_template.replace("__PPTX_NAME__", pptx_name)
    html_content = html_content.replace("__NUM_SLIDES__", str(len(ordered_slide_paths)))
    html_content = html_content.replace("__THUMBNAILS__", thumbnails_joined)
    html_content = html_content.replace("__SLIDE_META_JSON__", json.dumps(slide_meta_list))
    html_content = html_content.replace("__VIDEO_DATA_JSON__", json.dumps(video_data_obj))

    html_output_file = os.path.join(OUTPUT_DIR, "index.html")
    with open(html_output_file, 'w', encoding='utf-8') as out_f:
        out_f.write(html_content)

    log_success("Standalone HTML presentation generated successfully!")

class PptxConverterHTTPHandler(BaseHTTPRequestHandler):
    def get_mime_type(self, path):
        ext = os.path.splitext(path)[1].lower()
        mime_types = {
            ".html": "text/html; charset=utf-8",
            ".htm": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".gif": "image/gif",
            ".svg": "image/svg+xml",
            ".mp4": "video/mp4",
            ".webm": "video/webm",
            ".json": "application/json; charset=utf-8"
        }
        return mime_types.get(ext, "application/octet-stream")

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path

        if path in ["/", "/index.html"]:
            path = "/converter.html"

        if path == "/download":
            html_file = os.path.join(OUTPUT_DIR, "index.html")
            if not os.path.exists(html_file):
                self.send_error(404, "No converted presentation found.")
                return
            
            with open(html_file, 'rb') as f:
                content = f.read()

            filename = f"{active_presentation_name}.html"
            safe_filename = urllib.parse.quote(filename)

            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Disposition", f'attachment; filename="{safe_filename}"; filename*=UTF-8\'\'{safe_filename}')
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return

        # Serve static file
        local_path = os.path.abspath(os.path.join(WORKSPACE_DIR, path.lstrip("/")))
        if not local_path.startswith(WORKSPACE_DIR):
            self.send_error(403, "Access denied")
            return

        if os.path.isfile(local_path):
            mime = self.get_mime_type(local_path)
            with open(local_path, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
        else:
            self.send_error(404, "File not found")

    def do_POST(self):
        if self.path == "/upload":
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                filename_header = self.headers.get('X-File-Name', 'presentation.pptx')
                filename = urllib.parse.unquote(filename_header)
                pptx_name = os.path.splitext(filename)[0]

                log_info(f"Receiving file upload via Python HTTP server: {filename} ({content_length} bytes)...")
                file_bytes = self.rfile.read(content_length)

                temp_upload_dir = os.path.join(WORKSPACE_DIR, "temp_upload")
                os.makedirs(temp_upload_dir, exist_ok=True)
                temp_pptx_file = os.path.join(temp_upload_dir, "uploaded_presentation.pptx")

                with open(temp_pptx_file, 'wb') as tf:
                    tf.write(file_bytes)

                convert_pptx_to_html(temp_pptx_file, pptx_name)

                res_data = {"status": "success", "redirect": "/html_output/index.html"}
                json_response = json.dumps(res_data).encode('utf-8')

                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(json_response)))
                self.end_headers()
                self.wfile.write(json_response)
                log_success(f"Upload and conversion of {filename} completed successfully!")
            except Exception as e:
                log_error(f"Error handling file upload: {e}")
                err_msg = str(e).encode('utf-8')
                self.send_response(500)
                self.send_header("Content-Type", "text/plain")
                self.send_header("Content-Length", str(len(err_msg)))
                self.end_headers()
                self.wfile.write(err_msg)

def run_server():
    server_address = ('', PORT)
    httpd = HTTPServer(server_address, PptxConverterHTTPHandler)
    log_success(f"Python 3.12 Web Server running at: http://localhost:{PORT}/converter.html")
    log_info("Opening presentation converter dashboard in browser...")
    webbrowser.open(f"http://localhost:{PORT}/converter.html")
    
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        log_info("Shutting down Python server...")
        httpd.server_close()

if __name__ == "__main__":
    run_server()
