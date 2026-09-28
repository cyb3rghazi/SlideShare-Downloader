from flask import Flask, render_template, request, send_file, jsonify
import os
import re
import json
import time
import uuid
import shutil
import zipfile
import threading
import traceback
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from PIL import Image
from pptx import Presentation
import fitz  # PyMuPDF


# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------
app = Flask(__name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DOWNLOADS_DIR = os.path.join(BASE_DIR, 'downloads')
os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Progress store must exist at import time (not only under __main__)
app.config['PROGRESS'] = {}
PROGRESS_LOCK = threading.Lock()

UUID_RE = re.compile(
    r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
)
ALLOWED_HOSTS = {'www.slideshare.net', 'slideshare.net'}
REQUEST_HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/122.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
}
REQUEST_TIMEOUT = 30
CHUNK_SIZE = 8192
ALLOWED_FORMATS = {'zip', 'pdf', 'ppt'}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def is_valid_slideshare_url(url):
    if not url or not isinstance(url, str):
        return False
    try:
        parsed = urlparse(url)
    except Exception:
        return False
    if parsed.scheme not in ('http', 'https'):
        return False
    return parsed.netloc.lower() in ALLOWED_HOSTS


def session_dir_for(session_id):
    """Return the absolute, validated session directory."""
    if not session_id or not UUID_RE.fullmatch(session_id):
        raise ValueError('Invalid session id')
    path = os.path.abspath(os.path.join(DOWNLOADS_DIR, session_id))
    if not path.startswith(DOWNLOADS_DIR + os.sep):
        raise ValueError('Invalid session id')
    return path


def safe_filename(title, fallback='slides'):
    title = title or fallback
    cleaned = re.sub(r'[^A-Za-z0-9_.-]+', '_', title).strip('._')
    return (cleaned or fallback)[:80]


def normalize_host(host):
    """SlideShare sometimes returns the CDN host without a scheme."""
    host = (host or '').strip()
    if not host:
        return ''
    if host.startswith('//'):
        return 'https:' + host
    if not host.startswith(('http://', 'https://')):
        return 'https://' + host
    return host


def parse_slideshow(url):
    """Fetch the SlideShare page and pull the slideshow metadata out."""
    print(f"[parse] GET {url}")
    response = requests.get(url, headers=REQUEST_HEADERS, timeout=REQUEST_TIMEOUT)
    print(f"[parse] status={response.status_code} "
          f"final_url={response.url} len={len(response.content)}")
    response.raise_for_status()

    soup = BeautifulSoup(response.content, 'html.parser')
    tag = soup.select_one('#__NEXT_DATA__')
    if tag is None:
        raise ValueError('Could not find __NEXT_DATA__ in the page')

    data = json.loads(tag.text)
    slideshow = data['props']['pageProps']['slideshow']
    slides = slideshow['slides']

    image_sizes = slides['imageSizes']
    if not image_sizes:
        raise ValueError('No image sizes available')

    largest = image_sizes[-1]
    info = {
        'total_slides': int(slideshow['totalSlides']),
        'host': normalize_host(slides.get('host')),
        'image_location': slides.get('imageLocation', ''),
        'quality': largest.get('quality', ''),
        'width': largest.get('width', ''),
        'title': slides.get('title', 'slides'),
    }
    print(f"[parse] info={info}")
    return info


def build_image_url(info, index):
    return (
        f"{info['host']}/{info['image_location']}/{info['quality']}/"
        f"{info['title']}-{index}-{info['width']}.jpg"
    )


def set_progress(session_id, value):
    with PROGRESS_LOCK:
        app.config['PROGRESS'][session_id] = value


def get_progress(session_id):
    with PROGRESS_LOCK:
        return app.config['PROGRESS'].get(session_id, 0)


def clear_progress(session_id):
    with PROGRESS_LOCK:
        app.config['PROGRESS'].pop(session_id, None)


def clean_up_directory(directory, attempts=5, delay=1):
    if not directory or not os.path.isdir(directory):
        return
    for attempt in range(1, attempts + 1):
        try:
            shutil.rmtree(directory)
            print(f"[cleanup] removed {directory}")
            return
        except OSError as exc:
            print(f"[cleanup] error on {directory} (attempt {attempt}): {exc}")
            time.sleep(delay)
    print(f"[cleanup] failed to remove {directory} after {attempts} attempts")


def schedule_cleanup(path, session_id=None):
    """Delete `path` after the response has been fully sent."""
    def _cleanup():
        clean_up_directory(path)
        if session_id:
            clear_progress(session_id)
    threading.Thread(target=_cleanup, daemon=True).start()


# ---------------------------------------------------------------------------
# Image / PDF / PPTX builders
# ---------------------------------------------------------------------------
def download_images(info, selected_slides, slides_dir, session_id):
    os.makedirs(slides_dir, exist_ok=True)
    total = len(selected_slides)
    set_progress(session_id, 0.0)

    for done, index in enumerate(selected_slides, start=1):
        img_url = build_image_url(info, index)
        target = os.path.join(slides_dir, f"{index}.jpg")
        print(f"[download] GET {img_url}")

        with requests.get(
            img_url, headers=REQUEST_HEADERS,
            timeout=REQUEST_TIMEOUT, stream=True
        ) as r:
            print(f"[download] status={r.status_code} "
                  f"content-type={r.headers.get('content-type')}")
            r.raise_for_status()

            size = 0
            with open(target, 'wb') as fh:
                for chunk in r.iter_content(CHUNK_SIZE):
                    if chunk:
                        fh.write(chunk)
                        size += len(chunk)
            print(f"[download] saved {target} ({size} bytes)")

        set_progress(session_id, done / total)


def create_pdf_from_images(slides_dir, pdf_path):
    def sort_key(name):
        stem = os.path.splitext(name)[0]
        return int(stem) if stem.isdigit() else 0

    images = []
    for name in sorted(os.listdir(slides_dir), key=sort_key):
        if name.lower().endswith('.jpg'):
            with Image.open(os.path.join(slides_dir, name)) as im:
                images.append(im.convert('RGB'))

    if not images:
        return None

    images[0].save(pdf_path, save_all=True, append_images=images[1:])
    for im in images:
        im.close()
    print(f"[pdf] wrote {pdf_path} ({len(images)} pages)")
    return pdf_path


def create_ppt_from_pdf(pdf_path, ppt_path, temp_dir):
    os.makedirs(temp_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    ppt = Presentation()
    slide_w = ppt.slide_width
    slide_h = ppt.slide_height
    slide_ratio = slide_w / slide_h

    try:
        for page_number in range(len(doc)):
            page = doc.load_page(page_number)
            pix = page.get_pixmap()
            img_path = os.path.join(temp_dir, f"slide_{page_number + 1}.png")
            pix.save(img_path)

            # Blank layout (index 6) - avoids placeholder shapes on top of image
            slide = ppt.slides.add_slide(ppt.slide_layouts[6])

            with Image.open(img_path) as im:
                img_w, img_h = im.size

            img_ratio = img_w / img_h
            if img_ratio > slide_ratio:
                width = slide_w
                height = int(slide_w / img_ratio)
            else:
                height = slide_h
                width = int(slide_h * img_ratio)

            left = int((slide_w - width) / 2)
            top = int((slide_h - height) / 2)
            slide.shapes.add_picture(img_path, left, top,
                                     width=width, height=height)

            os.remove(img_path)
    finally:
        doc.close()

    ppt.save(ppt_path)
    print(f"[ppt] wrote {ppt_path}")
    return ppt_path


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/fetch_slides', methods=['POST'])
def fetch_slides():
    payload = request.get_json(silent=True) or {}
    url = payload.get('url', '')

    if not is_valid_slideshare_url(url):
        return jsonify({'error': 'Invalid SlideShare URL'}), 400

    try:
        info = parse_slideshow(url)
    except Exception as exc:
        traceback.print_exc()
        return jsonify({'error': f'Failed to fetch slides: {exc}'}), 500

    thumbnails = [
        build_image_url(info, i)
        for i in range(1, info['total_slides'] + 1)
    ]

    return jsonify({
        'slides': thumbnails,
        'session_id': str(uuid.uuid4()),
        'total_slides': info['total_slides'],
    })


@app.route('/download_images', methods=['POST'])
def download_images_route():
    url = request.form.get('url', '')
    format_type = (request.form.get('format') or '').lower()
    raw_slides = request.form.getlist('slides')
    session_id = request.form.get('session_id', '')

    print(f"[route] download_images url={url} format={format_type} "
          f"slides={raw_slides} session={session_id}")

    if not is_valid_slideshare_url(url):
        return "Invalid SlideShare URL.", 400
    if format_type not in ALLOWED_FORMATS:
        return "Invalid format specified.", 400

    try:
        session_dir = session_dir_for(session_id)
    except ValueError:
        return "Invalid session.", 400

    try:
        info = parse_slideshow(url)
    except Exception as exc:
        traceback.print_exc()
        return f"Failed to fetch slides: {exc}", 500

    # Validate and de-duplicate selected slides
    selected = set()
    for raw in raw_slides:
        # accept "1,2,3" or individual "1"
        for piece in str(raw).split(','):
            piece = piece.strip()
            if not piece:
                continue
            try:
                n = int(piece)
            except ValueError:
                continue
            if 1 <= n <= info['total_slides']:
                selected.add(n)
    selected = sorted(selected)

    if not selected:
        return "No valid slides selected.", 400

    slides_dir = os.path.join(session_dir, 'slides')
    temp_dir = os.path.join(session_dir, 'tmp')
    base_name = safe_filename(info['title'])

    try:
        download_images(info, selected, slides_dir, session_id)
    except Exception as exc:
        traceback.print_exc()
        clean_up_directory(session_dir)
        clear_progress(session_id)
        return f"Failed to download images: {exc}", 500

    try:
        if format_type == 'zip':
            out_path = os.path.join(session_dir, f"{base_name}.zip")
            with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for name in sorted(
                    os.listdir(slides_dir),
                    key=lambda n: int(os.path.splitext(n)[0])
                    if os.path.splitext(n)[0].isdigit() else 0
                ):
                    if name.lower().endswith('.jpg'):
                        zf.write(os.path.join(slides_dir, name), name)
            download_name = f"{base_name}.zip"

        elif format_type == 'pdf':
            out_path = os.path.join(session_dir, f"{base_name}.pdf")
            if not create_pdf_from_images(slides_dir, out_path):
                raise ValueError("No images to build PDF from")
            download_name = f"{base_name}.pdf"

        else:  # ppt
            pdf_path = os.path.join(session_dir, f"{base_name}.pdf")
            if not create_pdf_from_images(slides_dir, pdf_path):
                raise ValueError("No images to build PPT from")
            out_path = os.path.join(session_dir, f"{base_name}.pptx")
            create_ppt_from_pdf(pdf_path, out_path, temp_dir)
            os.remove(pdf_path)
            download_name = f"{base_name}.pptx"

    except Exception as exc:
        traceback.print_exc()
        clean_up_directory(session_dir)
        clear_progress(session_id)
        return f"Failed to build {format_type}: {exc}", 500

    response = send_file(out_path, as_attachment=True,
                         download_name=download_name)
    # Clean up only after the client has received the whole file
    response.call_on_close(lambda: (
        clean_up_directory(session_dir),
        clear_progress(session_id),
    ))
    return response


@app.route('/progress/<session_id>')
def progress(session_id):
    if not UUID_RE.fullmatch(session_id or ''):
        return jsonify({'progress': 0}), 400
    return jsonify({'progress': get_progress(session_id)})


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    # Never enable debug=True in production
    app.run(host='127.0.0.1', port=5000, debug=False, threaded=True)
