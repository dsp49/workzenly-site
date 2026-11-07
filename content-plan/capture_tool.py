#!/usr/bin/env python3
"""Fetch a software tool's logo and a screenshot of its home page into assets/tools/<slug>/.

Usage:
  python content-plan/capture_tool.py <tool-slug> <homepage-url> [--name="When I Work"] [--logo-url=URL] [--force]
  e.g. python content-plan/capture_tool.py when-i-work https://wheniwork.com/ --name="When I Work"
  --logo-url  use this image (an official press-kit PNG/JPG, or a Wikimedia Commons PNG thumbnail) as the logo
  --crop-logo=x,y,w,h  cut the logo out of home.jpg (1280x720 pixel coordinates) when no good icon exists;
              look at home.jpg first, pick a tight box around the brand mark in the header

Writes:
  assets/tools/<slug>/logo.png   square PNG, 160x160, on a white tile
  assets/tools/<slug>/home.jpg   1280x720 screenshot of the home page (first screen)
  assets/tools/<slug>/meta.json  name, url, capture date, sources, and quality flags

Re-running is cheap: existing files are kept unless --force is given. Always LOOK at home.jpg
afterwards. If it shows a captcha, an "access denied" page or only a cookie wall, delete it or
re-run with --force; never publish a broken screenshot.
"""
import html, io, json, os, re, shutil, subprocess, sys, tempfile, time, urllib.parse
from PIL import Image, ImageStat

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
]
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/130.0.0.0 Safari/537.36")


def curl(url, binary=False, timeout=30):
    r = subprocess.run(["curl", "-sL", "-m", str(timeout), "-A", UA, "-H", "Accept-Language: en-US,en;q=0.9", url],
                       capture_output=True)
    return r.stdout if binary else r.stdout.decode("utf-8", "replace")


def find_icons(page, base):
    icons = []
    for tag in re.findall(r"<link\b[^>]*>", page, re.I):
        rel = re.search(r'rel\s*=\s*["\']([^"\']+)', tag, re.I)
        href = re.search(r'href\s*=\s*["\']([^"\']+)', tag, re.I)
        if not rel or not href:
            continue
        rel = rel.group(1).lower()
        if "icon" not in rel:
            continue
        sizes = re.search(r'sizes\s*=\s*["\'](\d+)x', tag, re.I)
        size = int(sizes.group(1)) if sizes else (180 if "apple" in rel else 32)
        icons.append((size, urllib.parse.urljoin(base, html.unescape(href.group(1)))))
    icons.sort(reverse=True)
    return icons


def load_icon(data):
    im = Image.open(io.BytesIO(data))
    if im.format == "ICO":
        try:
            big = max(im.ico.sizes())
            im.size = big
            im = im.ico.getimage(big)
        except Exception:
            pass
    if getattr(im, "n_frames", 1) > 1:
        im.seek(0)
    return im.convert("RGBA")


def tile(im, side=160):
    t = Image.new("RGBA", (side, side), (255, 255, 255, 255))
    im = im.copy()
    im.thumbnail((side - 28, side - 28), Image.LANCZOS)
    if im.width < side - 60:  # small icons are scaled up a little, not past 2x
        f = min(2.0, (side - 40) / max(im.width, 1))
        im = im.resize((int(im.width * f), int(im.height * f)), Image.LANCZOS)
    t.paste(im, ((side - im.width) // 2, (side - im.height) // 2), im)
    return t.convert("RGB")


def lettermark(name, side=160):
    from PIL import ImageDraw, ImageFont
    t = Image.new("RGB", (side, side), (242, 242, 239))
    d = ImageDraw.Draw(t)
    letters = "".join(w[0] for w in re.split(r"[\s-]+", name) if w)[:2].upper() or "?"
    try:
        f = ImageFont.truetype("arialbd.ttf", 64)
    except Exception:
        f = ImageFont.load_default()
    w = d.textlength(letters, font=f)
    d.text(((side - w) / 2, side / 2 - 40), letters, font=f, fill=(15, 15, 15))
    return t


def chrome_dom(url):
    chrome = next((c for c in CHROME_CANDIDATES if os.path.exists(c)), None)
    if not chrome:
        return ""
    prof = tempfile.mkdtemp(prefix="wz-dom-")
    try:
        r = subprocess.run([chrome, "--headless=new", "--disable-gpu", "--no-first-run", f"--user-data-dir={prof}",
                            f"--user-agent={UA}", "--virtual-time-budget=8000", "--dump-dom", url],
                           capture_output=True, timeout=90)
        return r.stdout.decode("utf-8", "replace")
    except Exception:
        return ""
    finally:
        time.sleep(0.5)
        shutil.rmtree(prof, ignore_errors=True)


def get_logo(url, domain, name, explicit=None):
    cands = []
    if explicit:
        cands.append(explicit)
    page = curl(url)
    icons = find_icons(page, url) or find_icons(chrome_dom(url), url)
    cands += [src for _, src in icons if not src.lower().endswith(".svg")]
    cands += [urllib.parse.urljoin(url, "/apple-touch-icon.png"),
              f"https://icon.horse/icon/{domain}",
              f"https://www.google.com/s2/favicons?domain={domain}&sz=256",
              f"https://icons.duckduckgo.com/ip3/{domain}.ico"]
    def colorful(im):
        rgb = im.convert("RGB").resize((32, 32))
        b = rgb.tobytes()
        px = [b[i:i + 3] for i in range(0, len(b), 3)]
        return sum(max(p) - min(p) for p in px) / len(px) > 12

    real, horse = [], []
    for src in dict.fromkeys(cands):
        data = curl(src, binary=True)
        if not data or len(data) < 100:
            continue
        try:
            im = load_icon(data)
        except Exception:
            continue
        size = min(im.size)
        if explicit and src == explicit and size >= 32:
            return tile(im), src, size
        (horse if "icon.horse" in src else real).append((size, src, im))
    real.sort(key=lambda x: -x[0])
    horse = [h for h in horse if colorful(h[2])]
    if real and real[0][0] >= 64:
        size, src, im = real[0]
        return tile(im), src, size
    if horse and horse[0][0] >= 64:
        size, src, im = horse[0]
        return tile(im), src, size
    if real and real[0][0] >= 32:
        size, src, im = real[0]
        return tile(im), src, size
    return lettermark(name), "lettermark", 0


def screenshot(url, out_path):
    chrome = next((c for c in CHROME_CANDIDATES if os.path.exists(c)), None)
    if not chrome:
        return False, "no chrome"
    prof = tempfile.mkdtemp(prefix="wz-shot-")
    raw = os.path.join(prof, "shot.png")
    try:
        subprocess.run([chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run",
                        "--no-default-browser-check", f"--user-data-dir={prof}", f"--user-agent={UA}",
                        "--lang=en-US", "--window-size=1440,900", "--virtual-time-budget=9000",
                        f"--screenshot={raw}", url], capture_output=True, timeout=90)
        if not os.path.exists(raw):
            return False, "chrome produced no image"
        im = Image.open(raw).convert("RGB").crop((0, 0, 1440, 810)).resize((1280, 720), Image.LANCZOS)
        stat = ImageStat.Stat(im.convert("L"))
        if stat.stddev[0] < 12:
            return False, f"image looks blank (stddev {stat.stddev[0]:.1f})"
        im.save(out_path, quality=78, optimize=True, progressive=True)
        return True, "ok"
    finally:
        time.sleep(0.5)
        shutil.rmtree(prof, ignore_errors=True)


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    slug, url = argv[0], argv[1]
    force = "--force" in argv
    explicit = next((a.split("=", 1)[1] for a in argv if a.startswith("--logo-url=")), None)
    crop = next((a.split("=", 1)[1] for a in argv if a.startswith("--crop-logo=")), None)
    name = next((a.split("=", 1)[1] for a in argv if a.startswith("--name=")), None)
    if not re.fullmatch(r"[a-z0-9][a-z0-9-]*", slug):
        print("slug must be lowercase letters, digits and hyphens")
        return 2
    d = os.path.join(ROOT, "assets", "tools", slug)
    os.makedirs(d, exist_ok=True)
    meta_path = os.path.join(d, "meta.json")
    meta = json.load(open(meta_path, encoding="utf-8")) if os.path.exists(meta_path) else {}
    domain = urllib.parse.urlparse(url).netloc.replace("www.", "")
    meta.update(slug=slug, url=url, domain=domain, captured=time.strftime("%Y-%m-%d"))
    if name:
        meta["name"] = name

    logo_path = os.path.join(d, "logo.png")
    if crop:
        x, y, w, h = [int(v) for v in crop.split(",")]
        src_img = Image.open(os.path.join(d, "home.jpg")).convert("RGBA").crop((x, y, x + w, y + h))
        side = 160
        canvas = Image.new("RGBA", (max(w, h) + 24, max(w, h) + 24), (255, 255, 255, 255))
        canvas.paste(src_img, ((canvas.width - w) // 2, (canvas.height - h) // 2))
        canvas.convert("RGB").resize((side, side), Image.LANCZOS).save(logo_path, optimize=True)
        meta.update(logo_source=f"cropped from home.jpg at {crop}", logo_px=max(w, h), logo_flag="ok (cropped from screenshot)")
        json.dump(meta, open(meta_path, "w", encoding="utf-8"), indent=1)
        print("logo: cropped from screenshot", crop)
        return 0
    if force or explicit or not os.path.exists(logo_path):
        logo, src, size = get_logo(url, domain, meta.get("name") or slug.replace("-", " "), explicit)
        logo.save(logo_path, optimize=True)
        meta["logo_source"] = src
        meta["logo_px"] = size
        meta["logo_flag"] = "ok" if size >= 64 else ("small icon, check it" if size else "lettermark fallback, find an official logo")
        print(f"logo: {src} ({size}px) {meta['logo_flag']}")
    else:
        print("logo: kept existing")

    shot_path = os.path.join(d, "home.jpg")
    if force or not os.path.exists(shot_path):
        ok, why = screenshot(url, shot_path)
        meta["screenshot_ok"] = ok
        meta["screenshot_note"] = why
        print("screenshot:", why, "->", shot_path if ok else "")
    else:
        print("screenshot: kept existing")
    json.dump(meta, open(meta_path, "w", encoding="utf-8"), indent=1)
    print("folder:", os.path.relpath(d, ROOT).replace("\\", "/"))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
