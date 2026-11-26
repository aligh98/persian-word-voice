#!/usr/bin/env python3
"""Persian word pronunciation with ElevenLabs (eleven_v4), for language-learning apps.

The working state is one table, <work>/words.tsv. Each row is one word or expression.

  prepare  SRC --work DIR          read .xlsx/.csv/.txt, clean Unicode, flag homographs -> words.tsv
  list     DIR [--kind K] [--status S] [--ids ...]
  set      DIR ID [--spoken S] [--method base|carrier] [--sentence S] [--status S] [--note S]
  estimate DIR [--ids ... | --status S] [--method M]    ElevenLabs credits (1 credit per character)
  generate DIR [--ids ... | --status S] [--force] [--workers N] [--voice ID]
  review   DIR [--ids ... | --status S]                  listen.html with a "wrong/weak" checkbox per word
  export   DIR --out OUT [--names id|word] [--loudness [DB]]  final files + manifest.csv; --loudness evens out volume
  voices                                                 Persian voices in this ElevenLabs account

API key: environment variable ELEVENLABS_API_KEY. It is never written to disk.
"""
import argparse, base64, hashlib, random, csv, html, json, os, re, shutil, subprocess, sys, threading, time, unicodedata
import urllib.error, urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATA = HERE.parent / "data"
API = "https://api.elevenlabs.io/v1"
MODEL = "eleven_v4"
DEFAULT_VOICE = "RBEKbcQ62sKk2tsmPsAU"  # "Sara - Neutral Persian Female" (ElevenLabs Voice Library)
SETTINGS = {"stability": 0.5, "similarity_boost": 0.8}  # eleven_v4 ignores speed and style
SEED = 42
FORMAT = "mp3_44100_128"
COLS = ["id", "source", "original", "word", "meaning", "kind", "spoken", "method", "sentence",
        "status", "note", "made_from"]
MARKS = re.compile("[ً-ْٰ]")
HEADER_WORDS = {"کلمه", "کلمات", "واژه", "واژه‌ها", "لغت", "لغات", "word", "words", "term"}
INDEX_WORDS = {"ردیف", "#", "شماره", "id", "no", "no.", "row"}

for s in (sys.stdout, sys.stderr):
    try:
        s.reconfigure(encoding="utf-8")
    except Exception:
        pass


# ---------- text ----------

def normalize(s):
    """Persian letters instead of Arabic look-alikes, no tatweel, tidy spaces and ZWNJ.
    Arabic ي/ك/ى push the model towards an Arabic accent, so this always runs first."""
    s = unicodedata.normalize("NFC", str(s))
    s = s.replace("ي", "ی").replace("ى", "ی").replace("ك", "ک")
    s = s.replace("ـ", "")                      # tatweel
    s = s.replace(" ", " ").replace("‏", "").replace("‎", "")
    s = re.sub("‌+", "‌", s)               # one ZWNJ at most
    s = re.sub(" *‌ *", "‌", s)            # no space around ZWNJ
    s = re.sub(r"\s+", " ", s).strip(" ‌")
    s = s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "۰۱۲۳۴۵۶۷۸۹"))
    return s

def bare(s):
    return MARKS.sub("", normalize(s)).replace("‌", " ")

def load_homographs():
    table = {}
    p = DATA / "homographs.tsv"
    for line in p.read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        word, readings = line.split("\t", 1)
        table[bare(word)] = readings.strip()
    return table


# ---------- table ----------

def tsv_path(work):
    return Path(work) / "words.tsv"

def load(work):
    p = tsv_path(work)
    if not p.exists():
        sys.exit(f"no {p}. Run: prepare SRC --work {work}")
    with p.open(encoding="utf-8", newline="") as f:
        rows = list(csv.DictReader(f, delimiter="\t"))
    for r in rows:
        for c in COLS:
            r.setdefault(c, "")
    return rows

def save(work, rows):
    p = tsv_path(work)
    tmp = p.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS, delimiter="\t", extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    tmp.replace(p)

def pick(rows, ids=None, status=None, kind=None, method=None):
    out = rows
    if ids:
        want = set()
        for part in ids.split(","):
            part = part.strip()
            if "-" in part:
                a, b = part.split("-")
                want |= {f"{i:04d}" for i in range(int(a), int(b) + 1)}
            elif part:
                want.add(f"{int(part):04d}")
        out = [r for r in out if r["id"] in want]
    if status:
        out = [r for r in out if r["status"] in status.split(",")]
    if kind:
        out = [r for r in out if r["kind"] in kind.split(",")]
    if method:
        out = [r for r in out if r["method"] == method]
    return out

def text_for(r):
    return r["sentence"] if r["method"] == "carrier" else r["spoken"]


# ---------- prepare ----------

def read_source(src, column, meaning_column):
    src = Path(src)
    items = []  # (source, original, meaning)
    if src.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            sys.exit("pip install openpyxl")
        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
        for ws in wb.worksheets:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            col, mcol, start = find_columns(rows, column, meaning_column, ws.title)
            for r in rows[start:]:
                v = r[col] if col < len(r) else None
                if v is None or not str(v).strip():
                    continue
                m = r[mcol] if mcol is not None and mcol < len(r) and r[mcol] is not None else ""
                items.append((ws.title, str(v), str(m)))
    elif src.suffix.lower() in (".csv", ".tsv"):
        with src.open(encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f, delimiter="\t" if src.suffix.lower() == ".tsv" else ","))
        col, mcol, start = find_columns(rows, column, meaning_column, src.stem)
        for r in rows[start:]:
            if col < len(r) and r[col].strip():
                items.append((src.stem, r[col], r[mcol] if mcol is not None and mcol < len(r) else ""))
    else:
        for line in src.read_text(encoding="utf-8-sig").splitlines():
            if line.strip():
                items.append((src.stem, line, ""))
    return items

def find_columns(rows, column, meaning_column, title):
    """Return (word column, meaning column or None, first data row). Title and header rows are skipped."""
    head = [[normalize(c).lower() if c is not None else "" for c in r] for r in rows[:5]]
    column = normalize(column).lower() if column else None              # «معني» typed with Arabic ی still matches
    meaning_column = normalize(meaning_column).lower() if meaning_column else None
    mcol, col = None, None
    if column:
        for i, cells in enumerate(head):
            if column in cells:
                if meaning_column and meaning_column in cells:
                    mcol = cells.index(meaning_column)
                return cells.index(column), mcol, i + 1
        sys.exit(f"column «{column}» not found in the first 5 rows of «{title}»")
    start = 0
    for cells in head[:3]:  # title row and header row, e.g. «مصدرها» then «ردیف | کلمه | معنی»
        filled = [c for c in cells if c]
        if filled and (filled[0] == normalize(title).lower() or any(c in HEADER_WORDS | INDEX_WORDS for c in filled)
                       or (meaning_column and meaning_column in cells)):
            start += 1
            hit = next((j for j, c in enumerate(cells) if c in HEADER_WORDS), None)
            if hit is not None:
                col = hit
            if meaning_column and meaning_column in cells:
                mcol = cells.index(meaning_column)
        else:
            break
    if col is None:  # no header: the first column that holds text, not row numbers
        data = [r for r in rows[start:start + 30] if any(c is not None and str(c).strip() for c in r)]
        width = max((len(r) for r in data), default=0)
        def wordy(j):
            vals = [str(r[j]).strip() for r in data if j < len(r) and r[j] is not None and str(r[j]).strip()]
            return vals and sum(not re.fullmatch(r"[0-9۰-۹٠-٩.#\-\s]+", v) for v in vals) > len(vals) / 2
        col = next((j for j in range(width) if wordy(j)), 0)
    return col, mcol, start

def cmd_prepare(a):
    work = Path(a.work)
    work.mkdir(parents=True, exist_ok=True)
    if tsv_path(work).exists() and not a.force:
        sys.exit(f"{tsv_path(work)} exists (it holds your statuses). Use --force to rebuild it.")
    homo = load_homographs()
    rows, seen = [], {}
    for source, original, meaning in read_source(a.src, a.column, a.meaning_column):
        word = normalize(original)
        key = (bare(word), bare(meaning))
        if key in seen:
            seen[key]["note"] = (seen[key]["note"] + f" duplicate in {source};").strip()
            continue
        kind, note = "single", ""
        if MARKS.search(word):
            kind, note = "review", "source already has vowel marks: keep only marks that pick between readings"
        hits = [t for t in dict.fromkeys([bare(word)] + bare(word).split()) if t in homo]
        if hits:
            kind, note = "homograph", " / ".join(f"{t}: {homo[t]}" for t in hits)
        r = dict.fromkeys(COLS, "")
        r.update(id=f"{len(rows) + 1:04d}", source=source, original=original, word=word, meaning=normalize(meaning),
                 kind=kind, spoken=word, method="base", status="new", note=note)
        rows.append(r)
        seen[key] = r
    count = {}
    for r in rows:
        count[bare(r["word"])] = count.get(bare(r["word"]), 0) + 1
    for r in rows:  # same spelling listed twice with different meanings: the reading must be chosen per row
        if count[bare(r["word"])] > 1 and r["kind"] != "homograph":
            r["kind"], r["note"] = "homograph", "same spelling appears with another meaning in this list"
    save(work, rows)
    changed = sum(1 for r in rows if r["word"] != r["original"].strip())
    kinds = {k: sum(1 for r in rows if r["kind"] == k) for k in ("single", "homograph", "review")}
    print(f"{len(rows)} words -> {tsv_path(work)}")
    print(f"Unicode cleaned in {changed} words. single {kinds['single']}, homograph {kinds['homograph']}, review {kinds['review']}")
    print(f"credits for a base run: {sum(len(r['spoken']) for r in rows)}")


# ---------- list / set ----------

def cmd_list(a):
    rows = pick(load(a.work), a.ids, a.status, a.kind, a.method)
    for r in rows:
        extra = f"  «{r['sentence']}»" if r["method"] == "carrier" else ""
        print(f"{r['id']}\t{r['word']}\t{r['spoken']}\t{r['kind']}\t{r['method']}\t{r['status']}\t{r['meaning']}\t{r['note']}{extra}")
    print(f"({len(rows)} rows)", file=sys.stderr)

def check_sentence(sentence, spoken):
    s = normalize(sentence)
    if not re.search(r':\s*"' + re.escape(spoken) + r'"\s*$', s):
        return f'the sentence must end with : "{spoken}" (the spoken form, marks included)'
    lead = re.split(r"[\s‌،؛:.!؟?\"«»]+", bare(s[: s.rfind('"' + spoken)]))
    if bare(spoken).replace(" ", "") in [t for t in lead if t] or (" " in bare(spoken) and bare(spoken) in bare(s[: s.rfind('"' + spoken)])):
        return "the word must not appear earlier in the sentence: the lead-in should point to the word, not say it"
    return None

def cmd_set(a):
    rows = load(a.work)
    r = next((x for x in rows if x["id"] == f"{int(a.id):04d}"), None)
    if not r:
        sys.exit(f"no id {a.id}")
    if a.spoken is not None:
        sp = normalize(a.spoken)
        if bare(sp) != bare(r["word"]):
            sys.exit(f"spoken «{sp}» changes letters of «{r['word']}». Only vowel marks may be added; never respell.")
        r["spoken"] = sp
    if a.method:
        r["method"] = a.method
    if a.sentence is not None:
        r["sentence"] = re.sub(r"«([^«»]+)»\s*$", lambda m: chr(34) + m.group(1) + chr(34), normalize(a.sentence))
    if r["method"] == "carrier":
        if not r["sentence"]:
            sys.exit("method carrier needs --sentence")
        err = check_sentence(r["sentence"], r["spoken"])
        if err:
            sys.exit(err)
    for f in ("status", "note"):
        v = getattr(a, f)
        if v is not None:
            r[f] = v
    if text_for(r) != r["made_from"] and r["status"] in ("done", "ok"):
        r["status"] = "new"  # text changed: needs a new take
    save(a.work, rows)
    print(f"{r['id']}\t{r['word']}\tspoken «{r['spoken']}»\t{r['method']}\t{r['status']}" + (f"\t«{r['sentence']}»" if r["method"] == "carrier" else ""))


# ---------- estimate ----------

def cmd_estimate(a):
    rows = pick(load(a.work), a.ids, a.status, a.kind, a.method)
    base = sum(len(r["spoken"]) for r in rows)
    cur = sum(len(text_for(r)) for r in rows)
    print(f"{len(rows)} words. As set now: {cur} credits. All as base: {base} credits.")
    car = [r for r in rows if r["method"] == "carrier"]
    if car:
        print(f"{len(car)} carrier words: {sum(len(r['sentence']) for r in car)} credits "
              f"(base would be {sum(len(r['spoken']) for r in car)}).")
    try:
        sub = json.loads(api_get("/user/subscription"))
        left = sub["character_limit"] - sub["character_count"]
        print(f"account: {left} credits left this period ({sub.get('tier')}).")
    except SystemExit:
        raise
    except Exception as e:
        print(f"(could not read the account balance: {e})")


# ---------- ElevenLabs ----------

def key():
    k = os.environ.get("ELEVENLABS_API_KEY", "").strip()
    if not k:
        sys.exit("set ELEVENLABS_API_KEY first")
    return k

def api(path, body=None, tries=5):
    url = API + path
    data = json.dumps(body).encode() if body is not None else None
    for attempt in range(tries):
        req = urllib.request.Request(url, data=data, headers={"xi-api-key": key(), "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                return resp.read()
        except urllib.error.HTTPError as e:
            msg = e.read().decode("utf-8", "replace")[:400]
            if e.code in (429, 500, 502, 503, 504) and attempt < tries - 1:
                time.sleep(4 * (attempt + 1) + random.uniform(0, 2))  # jitter: parallel workers do not retry in lockstep
                continue
            if e.code == 401 or "authentication_error" in msg or "invalid_api_key" in msg:
                raise RuntimeError(f"401: the API key is wrong or lacks permission. {msg}")
            if e.code in (400, 404) and ("voice_not_found" in msg or "voice_id" in msg):
                raise RuntimeError(f"{e.code}: voice not available. Add it to My Voices from the Voice Library, or pass --voice. {msg}")
            if "quota" in msg.lower() or "credits" in msg.lower():
                raise RuntimeError(f"{e.code}: not enough ElevenLabs credits. {msg}")
            raise RuntimeError(f"HTTP {e.code}: {msg}")
        except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as e:
            if attempt < tries - 1:
                time.sleep(4 * (attempt + 1) + random.uniform(0, 2))  # jitter: parallel workers do not retry in lockstep
                continue
            raise RuntimeError(f"network: {e}")

def api_get(path):
    return api(path)

def tts_body(text):
    return {"text": text, "model_id": MODEL, "language_code": "fa", "voice_settings": SETTINGS, "seed": SEED}


# ---------- carrier cut ----------

TRIM = ("silenceremove=start_periods=1:start_threshold=-45dB:start_silence=0.03,areverse,"
        "silenceremove=start_periods=1:start_threshold=-45dB:start_silence=0.12,areverse")

def need_ffmpeg():
    if not shutil.which("ffmpeg"):
        raise RuntimeError("ffmpeg is needed for carrier cuts (install it and retry)")

def silence_gaps(path):
    r = subprocess.run(["ffmpeg", "-v", "info", "-i", str(path), "-af", "silencedetect=n=-40dB:d=0.06", "-f", "null", "-"],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    return list(zip(map(float, re.findall(r"silence_start: ([0-9.]+)", r.stderr)),
                    map(float, re.findall(r"silence_end: ([0-9.]+)", r.stderr))))

def cut_last_word(full, alignment, spoken, out):
    """Cut from the silence before the last occurrence of the word to the end.
    Timestamps alone can land 30-70 ms after the real onset, so the cut goes inside the silence gap."""
    chars = "".join(alignment["characters"])
    k = chars.rfind(spoken)
    if k < 0:  # the alignment may drop vowel marks or ZWNJ: fall back to the opening quote before the word
        q = chars.rstrip().rstrip('"').rstrip()
        k = max(chars.rfind('"', 0, len(q)), chars.rfind("«", 0, len(q))) + 1
        if k <= 0:
            raise RuntimeError("word not found in the alignment")
    t = alignment["character_start_times_seconds"][k]
    gaps = sorted([g for g in silence_gaps(full) if g[0] < t + 0.15], key=lambda g: abs(g[1] - t))
    start = max(gaps[0][0], gaps[0][1] - 0.05) if gaps and abs(gaps[0][1] - t) < 0.6 else max(0.0, t - 0.12)
    subprocess.run(["ffmpeg", "-y", "-v", "error", "-i", str(full), "-ss", f"{start:.3f}", "-af", TRIM,
                    "-c:a", "libmp3lame", "-b:a", "128k", str(out)], check=True)


# ---------- generate ----------

def make(work, r, voice):
    audio = Path(work) / "audio"
    audio.mkdir(exist_ok=True)
    out = audio / f"{r['id']}.mp3"
    if r["method"] == "carrier":
        need_ffmpeg()
        err = check_sentence(r["sentence"], r["spoken"])
        if err:
            raise RuntimeError(err)
        full_dir = audio / "carrier_full"
        full_dir.mkdir(exist_ok=True)
        full = full_dir / f"{r['id']}.mp3"
        d = json.loads(api(f"/text-to-speech/{voice}/with-timestamps?output_format={FORMAT}", tts_body(r["sentence"])))
        full.write_bytes(base64.b64decode(d["audio_base64"]))
        (full_dir / f"{r['id']}.json").write_text(json.dumps(d["alignment"], ensure_ascii=False), encoding="utf-8")
        cut_last_word(full, d["alignment"], r["spoken"], out)
    else:
        out.write_bytes(api(f"/text-to-speech/{voice}?output_format={FORMAT}", tts_body(r["spoken"])))
    return out

def cmd_generate(a):
    rows = load(a.work)
    # one voice per project: a list that switches voice halfway sounds like two speakers in the app
    cfg_path = Path(a.work) / "config.json"
    cfg = json.loads(cfg_path.read_text(encoding="utf-8")) if cfg_path.exists() else {}
    filtered = any((a.ids, a.status, a.kind, a.method))
    if a.voice and cfg.get("voice") and a.voice != cfg["voice"]:
        if filtered or not a.force:
            sys.exit(f"this project uses voice {cfg['voice']}. Switching voice remakes every word "
                     f"(about {sum(len(text_for(r)) for r in rows)} credits): generate {a.work} --voice {a.voice} --force")
    a.voice = a.voice or cfg.get("voice") or DEFAULT_VOICE
    todo = pick(rows, a.ids, a.status, a.kind, a.method)
    if not a.force:
        todo = [r for r in todo if r["made_from"] != text_for(r) or not (Path(a.work) / "audio" / f"{r['id']}.mp3").exists()]
    if not todo:
        print("nothing to do (all selected words already made from their current text; --force to redo)")
        return
    if not a.dry_run:
        cfg_path.write_text(json.dumps({**cfg, "voice": a.voice}), encoding="utf-8")
    credits = sum(len(text_for(r)) for r in todo)
    print(f"{len(todo)} words, about {credits} credits, voice {a.voice}, {MODEL}")
    if a.dry_run:
        return
    lock, done, failed = threading.Lock(), [0], []

    def one(r):
        try:
            make(a.work, r, a.voice)
            with lock:
                r["made_from"], r["status"] = text_for(r), "done"
                done[0] += 1
                if done[0] % 20 == 0:
                    save(a.work, rows)
                    print(f"  {done[0]}/{len(todo)}", flush=True)
        except Exception as e:
            with lock:
                failed.append((r["id"], r["word"], str(e)))
                if "credits" in str(e) or "401" in str(e):
                    fatal.append(str(e))
                    stop.set()  # no point sending the rest: every call would fail the same way

    def guarded(r):
        if not stop.is_set():
            one(r)

    stop, fatal = threading.Event(), []
    try:
        with ThreadPoolExecutor(max_workers=max(1, a.workers)) as ex:
            list(ex.map(guarded, todo))
    finally:
        save(a.work, rows)  # progress survives a stop (no credits, bad key, Ctrl+C)
    skipped = len(todo) - done[0] - len(failed)
    print(f"done {done[0]}, failed {len(failed)}" + (f", not started {skipped}" if skipped else ""))
    for f in failed:
        print("  FAILED", *f, sep="\t")
    if fatal:
        sys.exit(f"stopped: {fatal[0]}\nFix it, then run generate again: finished words are skipped.")


# ---------- review ----------

PAGE = """<!doctype html><html lang="fa" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>بازبینی تلفظ</title><style>
:root{--bg:#fff;--fg:#111;--mut:#666;--line:#e3e3e3;--hi:#fff4d6}
@media (prefers-color-scheme:dark){:root{--bg:#141414;--fg:#eee;--mut:#9a9a9a;--line:#2c2c2c;--hi:#3a3220}}
body{background:var(--bg);color:var(--fg);font-family:Vazirmatn,Tahoma,sans-serif;margin:0;padding:16px}
table{border-collapse:collapse;width:100%}td,th{border-bottom:1px solid var(--line);padding:6px 8px;text-align:right;vertical-align:middle}
th{position:sticky;top:0;background:var(--bg)}td.w{font-size:20px}small{color:var(--mut)}tr.bad{background:var(--hi)}
audio{height:30px;width:190px}input[type=text]{width:100%;min-width:120px;font:inherit}
#bar{position:sticky;bottom:0;background:var(--bg);padding:10px 0;border-top:1px solid var(--line)}textarea{width:100%;height:90px;font:inherit}
button{font:inherit;padding:6px 14px}</style>
<h2>بازبینی تلفظ: __N__ واژه</h2><p><small>هر واژه‌ای که تلفظش غلط یا ضعیف است را علامت بزنید و در صورت نیاز توضیح دهید. در پایان «ساخت فهرست» را بزنید و متن را برای ایجنت هوش مصنوعی بفرستید.</small></p>
<table><tr><th>#</th><th>کلمه</th><th>متن ارسالی</th><th>روش</th><th>صدا</th><th>غلط/ضعیف</th><th>توضیح</th></tr>__ROWS__</table>
<div id="bar"><button onclick="build()">ساخت فهرست</button> <small id="msg"></small><textarea id="out" readonly></textarea></div>
<script>
function build(){const l=[];document.querySelectorAll('tr[data-id]').forEach(t=>{if(t.querySelector('input[type=checkbox]').checked){l.push(t.dataset.id+'\\t'+t.dataset.w+'\\t'+t.querySelector('input[type=text]').value.trim())}});
const s=l.length?'غلط/ضعیف:\\n'+l.join('\\n'):'همه درست بود.';const o=document.getElementById('out');o.value=s;o.select();
try{navigator.clipboard.writeText(s).then(()=>{document.getElementById('msg').textContent='کپی شد.'},()=>{})}catch(e){}}
document.addEventListener('change',e=>{if(e.target.type==='checkbox')e.target.closest('tr').classList.toggle('bad',e.target.checked)});
</script></html>"""

def cmd_review(a):
    rows = [r for r in pick(load(a.work), a.ids, a.status, a.kind, a.method)
            if (Path(a.work) / "audio" / f"{r['id']}.mp3").exists()]
    out = []
    for r in rows:
        sent = html.escape(r["sentence"] if r["method"] == "carrier" else r["spoken"])
        out.append(f'<tr data-id="{r["id"]}" data-w="{html.escape(r["word"])}"><td><small>{r["id"]}</small></td>'
                   f'<td class=w>{html.escape(r["word"])}<br><small>{html.escape(r["meaning"])}</small></td><td>{sent}</td>'
                   f'<td><small>{r["method"]}</small></td><td><audio controls preload=none src="audio/{r["id"]}.mp3?v={hashlib.md5(r["made_from"].encode()).hexdigest()[:8]}"></audio></td>'
                   f'<td><input type=checkbox></td><td><input type=text></td></tr>')
    p = Path(a.work) / "listen.html"
    p.write_text(PAGE.replace("__N__", str(len(rows))).replace("__ROWS__", "".join(out)), encoding="utf-8")
    print(p.resolve())


# ---------- export ----------

def safe_name(s):
    s = MARKS.sub("", s).replace("‌", "_").replace(" ", "_")
    return re.sub(r'[\\/:*?"<>|]', "", s)

def level(src, dst, target_db, peak_db=-1.0):
    """Bring a clip to the same average loudness (RMS, dBFS) as every other clip, without clipping.
    RMS works on clips under 0.4 s, where EBU loudnorm has too little audio to measure."""
    r = subprocess.run(["ffmpeg", "-v", "info", "-i", str(src), "-af", "volumedetect", "-f", "null", "-"],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    mean = re.search(r"mean_volume: (-?[0-9.]+) dB", r.stderr)
    peak = re.search(r"max_volume: (-?[0-9.]+) dB", r.stderr)
    if not mean or not peak:
        raise RuntimeError(f"could not measure {src}")
    gain = min(target_db - float(mean.group(1)), peak_db - float(peak.group(1)))
    subprocess.run(["ffmpeg", "-y", "-v", "error", "-i", str(src), "-af", f"volume={gain:.2f}dB",
                    "-c:a", "libmp3lame", "-b:a", "128k", str(dst)], check=True)

def cmd_export(a):
    if a.loudness is not None:
        need_ffmpeg()
    rows = load(a.work)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    man, missing, flagged, used = [], [], [], set()
    for r in rows:
        src = Path(a.work) / "audio" / f"{r['id']}.mp3"
        if not src.exists():
            missing.append(r["id"])
            continue
        if r["status"] == "flagged":
            flagged.append(r["id"])
        name = (safe_name(r["word"]) if a.names == "word" else r["id"]) + ".mp3"
        if name in used:
            name = f"{name[:-4]}_{r['id']}.mp3"
        used.add(name)
        if a.loudness is not None:
            level(src, out / name, a.loudness)
        else:
            shutil.copyfile(src, out / name)
        man.append({"id": r["id"], "word": r["word"], "meaning": r["meaning"], "file": name, "method": r["method"],
                    "spoken": r["spoken"], "status": r["status"]})
    with (out / "manifest.csv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(man[0]) if man else ["id"])
        w.writeheader()
        w.writerows(man)
    print(f"{len(man)} files -> {out.resolve()} (+ manifest.csv)")
    if missing:
        print(f"no audio yet for {len(missing)}: {','.join(missing[:30])}{' ...' if len(missing) > 30 else ''}")
    if flagged:
        print(f"still flagged as wrong/weak: {','.join(flagged)}")


# ---------- voices ----------

def cmd_voices(a):
    d = json.loads(api_get("/voices"))
    for v in d.get("voices", []):
        lab = v.get("labels") or {}
        langs = {(x.get("language") or "") for x in (v.get("verified_languages") or [])}
        persian = "fa" in langs or "persian" in (v.get("name", "") + json.dumps(lab)).lower() or "فارسی" in v.get("name", "")
        if persian or a.all:
            print(f"{v['voice_id']}\t{v['name']}\t{v.get('category')}\t{lab.get('gender', '')}")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    def sel(q):
        q.add_argument("work")
        q.add_argument("--ids", help="e.g. 3,7,12-20")
        q.add_argument("--status", help="new,done,flagged,ok")
        q.add_argument("--kind", help="single,homograph,review")
        q.add_argument("--method", choices=["base", "carrier"])
        return q

    q = sub.add_parser("prepare"); q.add_argument("src"); q.add_argument("--work", required=True)
    q.add_argument("--column"); q.add_argument("--meaning-column"); q.add_argument("--force", action="store_true")
    q.set_defaults(f=cmd_prepare)
    sel(sub.add_parser("list")).set_defaults(f=cmd_list)
    q = sub.add_parser("set"); q.add_argument("work"); q.add_argument("id")
    q.add_argument("--spoken"); q.add_argument("--method", choices=["base", "carrier"]); q.add_argument("--sentence")
    q.add_argument("--status", choices=["new", "done", "flagged", "ok"]); q.add_argument("--note")
    q.set_defaults(f=cmd_set)
    sel(sub.add_parser("estimate")).set_defaults(f=cmd_estimate)
    q = sel(sub.add_parser("generate")); q.add_argument("--force", action="store_true"); q.add_argument("--dry-run", action="store_true")
    q.add_argument("--workers", type=int, default=2); q.add_argument("--voice", help="voice ID; saved for the project on first use (default: Sara)")
    q.set_defaults(f=cmd_generate)
    sel(sub.add_parser("review")).set_defaults(f=cmd_review)
    q = sub.add_parser("export"); q.add_argument("work"); q.add_argument("--out", required=True)
    q.add_argument("--names", choices=["id", "word"], default="id")
    q.add_argument("--loudness", type=float, nargs="?", const=-20.0, default=None,
                   help="match every clip to this average level in dBFS (default -20 when given without a value)")
    q.set_defaults(f=cmd_export)
    q = sub.add_parser("voices"); q.add_argument("--all", action="store_true"); q.set_defaults(f=cmd_voices)
    a = p.parse_args()
    try:
        a.f(a)
    except RuntimeError as e:
        sys.exit(str(e))


if __name__ == "__main__":
    main()
