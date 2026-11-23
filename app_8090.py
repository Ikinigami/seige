"""
app_8090.py  -  SIEGE web studio (tablet-friendly, dark mode).

Run:   python app_8090.py          ->  http://localhost:8090   (Codespaces: open the forwarded 8090 port)

PORT 8080 is reserved for the Yellow Stick Scarf video studio. This app refuses to start on 8080 and
only ever reads/writes inside ./output.
"""
import os
import re
import sys
import threading
import uuid

from flask import Flask, jsonify, request, send_from_directory, Response

import siege_engine as engine

PORT = int(os.environ.get("PORT", "8090"))
if PORT == 8080:
    sys.exit("Refusing to start: port 8080 is reserved for Yellow Stick Scarf. Use PORT=8090.")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024          # whole upload
MAX_FILE = 15 * 1024 * 1024                                  # per file
MAX_FILES = 10
SAFE_NAME = re.compile(r"^siege_v\d+(_FINAL)?\.xlsx$")

JOBS, JOBS_LOCK = {}, threading.Lock()


def _worker(job_id, sources, fresh):
    job = JOBS[job_id]

    def on_stage(key, status, detail):
        job["stages"][key] = {"status": status, "detail": detail}

    try:
        job["report"] = engine.run_pipeline(sources, fresh, on_stage)
        job["state"] = "done"
    except Exception as e:  # never let a thread die silently
        job["state"], job["error"] = "error", f"{type(e).__name__}: {e}"


@app.post("/api/run")
def api_run():
    with JOBS_LOCK:
        if any(j["state"] == "running" for j in JOBS.values()):
            return jsonify(error="A run is already in progress. Wait for it to finish."), 409
        sources = []
        text = (request.form.get("text") or "").strip()
        if text:
            sources.append({"kind": "text", "name": "pasted text", "text": text})
        files = [f for f in request.files.getlist("files") if f and f.filename][:MAX_FILES]
        for f in files:
            data = f.read(MAX_FILE + 1)
            if len(data) > MAX_FILE:
                return jsonify(error=f"{f.filename} is larger than 15 MB."), 413
            sources.append({"kind": "file", "name": os.path.basename(f.filename), "data": data})
        if not sources:
            return jsonify(error="Paste some text or choose a photo / CSV / TXT file first."), 400
        job_id = uuid.uuid4().hex[:10]
        JOBS[job_id] = {"state": "running", "report": None, "error": None,
                        "stages": {k: {"status": "idle", "detail": ""} for k, _ in engine.STAGES}}
        for old in list(JOBS)[:-20]:
            JOBS.pop(old, None)
    threading.Thread(target=_worker, args=(job_id, sources, request.form.get("fresh") == "1"),
                     daemon=True).start()
    return jsonify(job_id=job_id)


@app.get("/api/job/<job_id>")
def api_job(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify(error="Unknown job"), 404
    return jsonify(state=job["state"], stages=job["stages"], report=job["report"], error=job["error"])


@app.get("/api/state")
def api_state():
    return jsonify(engine.app_status() | {"stages": [{"key": k, "label": l} for k, l in engine.STAGES]})


@app.post("/api/approve")
def api_approve():
    body = request.get_json(silent=True) or {}
    if body.get("confirm") is not True:
        return jsonify(ok=False, error="Tick the confirmation first."), 400
    try:
        version = int(body.get("version"))
    except (TypeError, ValueError):
        return jsonify(ok=False, error="Bad version."), 400
    res = engine.approve(version, body.get("approved_by", ""))
    return jsonify(res), (200 if res.get("ok") else 400)


@app.get("/download/<name>")
def download(name):
    if not SAFE_NAME.match(name):
        return "Not found", 404
    return send_from_directory(engine.OUTPUT_DIR, name, as_attachment=True)


@app.get("/health")
def health():
    return jsonify(ok=True, port=PORT)


@app.get("/")
def index():
    return Response(PAGE, mimetype="text/html")


PAGE = r"""<!doctype html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SIEGE Studio</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600&family=Barlow+Condensed:wght@600;700&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#0e1621; --panel:#16212e; --panel2:#1c2a3a; --line:#2a3b4f; --text:#e6edf3; --muted:#8fa1b3;
  --active:#6cc4ff; --done:#5fd0a0; --flag:#f0b955; --err:#ef6b6b; --gate:#b79cff;
  --head:'Barlow Condensed','Arial Narrow',system-ui,sans-serif; --body:'Barlow',system-ui,-apple-system,'Segoe UI',sans-serif;
}
*{box-sizing:border-box}
html,body{margin:0;background:var(--bg);color:var(--text);font:16px/1.5 var(--body)}
body{padding:16px;max-width:1100px;margin:0 auto}
header{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end;justify-content:space-between;margin-bottom:14px}
h1{font:700 40px/1 var(--head);letter-spacing:.04em;margin:0}
header p{margin:2px 0 0;color:var(--muted)}
.chips{display:flex;gap:8px;flex-wrap:wrap}
.chip{border:1px solid var(--line);background:var(--panel);border-radius:999px;padding:4px 12px;font-size:14px;color:var(--muted)}
.chip.ok{color:var(--done);border-color:#2f6b57}.chip.bad{color:var(--flag);border-color:#7a5f27}
h2{font:600 22px/1.2 var(--head);margin:0 0 10px;letter-spacing:.02em}
h3{font:600 17px/1.2 var(--head);margin:18px 0 8px;color:var(--muted);letter-spacing:.03em}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:14px}

/* pipeline: a real sequence, so the step numbers carry meaning */
.pipe{display:flex;gap:0;overflow-x:auto;padding:4px 2px 10px;margin:0 0 14px;list-style:none;scrollbar-width:thin}
.node{position:relative;flex:1 1 100px;min-width:100px;padding:8px 8px 8px 9px;background:var(--panel);
  border:1.5px solid var(--line);border-radius:10px;margin-right:12px;transition:border-color .25s,background .25s}
.node:last-child{margin-right:0}
.node:not(:last-child)::after{content:"";position:absolute;right:-12px;top:50%;width:12px;height:2px;background:var(--line)}
.node .n{display:inline-grid;place-items:center;width:24px;height:24px;border-radius:50%;background:var(--panel2);
  font:600 13px var(--head);color:var(--muted);margin-right:6px}
.node b{font:600 15px var(--head);letter-spacing:.01em}
.node small{display:block;margin-top:4px;color:var(--muted);font-size:12.5px;line-height:1.3;min-height:32px;overflow-wrap:anywhere}
.node.active{border-color:var(--active);animation:pulse 1.1s ease-in-out infinite}
.node.active .n{background:var(--active);color:#06202f}
.node.done{border-color:var(--done)} .node.done .n{background:var(--done);color:#062a1d}
.node.flagged{border-color:var(--flag);background:#2a2518} .node.flagged .n{background:var(--flag);color:#2e2106}
.node.error{border-color:var(--err);background:#2d1a1d} .node.error .n{background:var(--err);color:#2d0a0a}
.node.waiting{border-color:var(--gate);animation:pulse 1.6s ease-in-out infinite} .node.waiting .n{background:var(--gate);color:#1d1238}
.node.idle{opacity:.7}
@keyframes pulse{0%,100%{box-shadow:0 0 0 0 rgba(108,196,255,.0)}50%{box-shadow:0 0 0 5px rgba(108,196,255,.18)}}
@media (prefers-reduced-motion:reduce){.node.active,.node.waiting{animation:none}}

textarea,input[type=text]{width:100%;background:var(--bg);color:var(--text);border:1px solid var(--line);border-radius:10px;
  padding:12px;font:16px/1.4 var(--body)}
textarea{min-height:130px;resize:vertical}
textarea:focus,input:focus,button:focus-visible,summary:focus-visible{outline:2px solid var(--active);outline-offset:2px}
.row{display:flex;flex-wrap:wrap;gap:12px;align-items:center;margin-top:12px}
.file{flex:1 1 260px;border:1.5px dashed var(--line);border-radius:10px;padding:12px;color:var(--muted)}
.file input{width:100%;color:var(--text)}
label.chk{display:flex;gap:10px;align-items:flex-start;min-height:44px;cursor:pointer}
label.chk input{width:22px;height:22px;margin-top:2px;flex:0 0 auto;accent-color:var(--done)}
button,.btn{font:600 17px var(--body);min-height:48px;padding:0 22px;border-radius:10px;border:0;cursor:pointer;
  background:var(--active);color:#06202f;text-decoration:none;display:inline-flex;align-items:center}
button.sec,.btn.sec{background:var(--panel2);color:var(--text);border:1px solid var(--line)}
button.go{background:var(--done);color:#062a1d}
button:disabled{opacity:.4;cursor:not-allowed}
.banner{border-radius:10px;padding:12px 14px;font:600 18px var(--head);letter-spacing:.03em;display:flex;flex-wrap:wrap;gap:6px 16px;align-items:center}
.banner.wait{background:#251f3b;border:1.5px solid var(--gate);color:#d8c9ff}
.banner.final{background:#143427;border:1.5px solid var(--done);color:#b4f0d3}
.banner span{font:400 15px var(--body);color:var(--text)}
ul.plain,ol.plain{margin:0;padding-left:20px} ul.plain li,ol.plain li{margin:6px 0}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-top:12px}
.kpi{background:var(--panel2);border-radius:10px;padding:10px 12px}
.kpi small{color:var(--muted);display:block}.kpi strong{font:600 22px var(--head)}
.tablewrap{overflow-x:auto;border:1px solid var(--line);border-radius:10px}
table{border-collapse:collapse;width:100%;min-width:640px;font-size:14.5px}
th,td{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{background:var(--panel2);font:600 15px var(--head)}
.pill{display:inline-block;border-radius:999px;padding:1px 10px;font-size:13px;font-weight:600}
.pill.HIGH{background:#4a2227;color:#ffb3b3}.pill.MEDIUM{background:#473a1c;color:#ffd98a}.pill.LOW{background:#1c3a2f;color:#a8e8cc}
.muted{color:var(--muted)} .err{color:var(--err);font-weight:600}
.hist{display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;padding:8px 0;border-bottom:1px solid var(--line)}
.hist:last-child{border:0}
@media (max-width:640px){body{padding:10px}h1{font-size:34px}}
</style></head>
<body>
<header>
  <div><h1>SIEGE</h1><p>Operations, Inventory &amp; Bookkeeping Studio</p></div>
  <div class="chips"><span class="chip" id="chipKey">Gemini key: checking</span><span class="chip">Port 8090</span><span class="chip" id="chipVer">No runs yet</span></div>
</header>

<ol class="pipe" id="pipe"></ol>

<section class="card">
  <h2>Send SIEGE something to read</h2>
  <textarea id="text" placeholder="Paste a receipt, an OR, stock notes, or a CSV (with a header row like Date, Vendor, OR No, Item, Qty, Unit Price, Amount, Category, Type)"></textarea>
  <div class="row">
    <div class="file">Photos (receipt, OR, invoice, stock note) or .csv / .txt files
      <input id="files" type="file" multiple accept="image/*,.csv,.txt,.tsv,.heic"></div>
  </div>
  <div class="row">
    <button id="runBtn">Run SIEGE</button>
    <label class="chk"><input type="checkbox" id="fresh"><span>Start a fresh ledger<br><small class="muted">Off = add to earlier versions and check duplicates across runs</small></span></label>
  </div>
  <p id="msg" class="muted" role="status" style="margin:10px 0 0"></p>
</section>

<section class="card" id="result" hidden></section>
<section class="card" id="histCard" hidden><h2>Versions</h2><div id="hist"></div></section>

<script>
const $ = s => document.querySelector(s);
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let STAGES = [], current = null, polling = null;

function paintPipe(stages){
  $('#pipe').innerHTML = STAGES.map((s,i)=>{
    const st = (stages && stages[s.key]) || {status:'idle',detail:''};
    return `<li class="node ${esc(st.status)}"><span class="n">${i+1}</span><b>${esc(s.label)}</b><small>${esc(st.detail || ({idle:'Waiting',active:'Working...'}[st.status]||''))}</small></li>`;
  }).join('');
}
const money = (x,c) => (x<0?'-':'') + c + Math.abs(x).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2});

function renderReport(r){
  current = r;
  const final = r.status === 'FINAL', s = r.summary, c = r.currency;
  const esc_rows = (r.escalations||[]).map(e=>`<tr><td><span class="pill ${esc(e.severity)}">${esc(e.severity)}</span><br><small class="muted">${e.held?'HELD':'posted'}</small></td>
    <td>${esc(e.issue)}<br><small class="muted">${esc(e.source)} &middot; ${esc(e.field)}${e.value?': '+esc(e.value):''}</small></td><td>${esc(e.description)}</td><td>${esc(e.detail)}</td></tr>`).join('');
  $('#result').hidden = false;
  $('#result').innerHTML = `
    <div class="banner ${final?'final':'wait'}">Status: ${esc(r.status)}
      <span>${final ? 'Approved by '+esc(r.approved_by)+' on '+esc(r.approved_at) : 'Version '+r.version+' is paused until you approve it.'}</span></div>

    <h3>What happened</h3>
    <ul class="plain">${r.what_happened.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>
    <div class="kpis">
      <div class="kpi"><small>Total sales</small><strong>${esc(money(s.sales,c))}</strong></div>
      <div class="kpi"><small>Total expenses</small><strong>${esc(money(s.expenses,c))}</strong></div>
      <div class="kpi"><small>Net</small><strong>${esc(money(s.net,c))}</strong></div>
      <div class="kpi"><small>Stock items</small><strong>${s.items}</strong></div>
      <div class="kpi"><small>Held for review</small><strong>${r.stats.held}</strong></div>
    </div>

    <h3>Output file</h3>
    <div class="row" style="margin-top:0">
      <a class="btn" href="/download/${esc(r.file)}">Download ${esc(r.file)}</a>
      ${r.final_file ? `<a class="btn go" href="/download/${esc(r.final_file)}">Download ${esc(r.final_file)}</a>` : ''}
    </div>
    <p class="muted">Saved in <code>output/</code>. Formulas recalculate when opened in Excel or Google Sheets.</p>

    <h3>Items to double-check / escalations (${(r.escalations||[]).length})</h3>
    ${esc_rows ? `<div class="tablewrap"><table><tr><th>Level</th><th>Issue</th><th>Item</th><th>What to check</th></tr>${esc_rows}</table></div>`
               : '<p class="muted">Nothing was flagged in this run. Still verify against the originals.</p>'}

    <h3>Human verification checklist</h3>
    <div id="checks">${r.checklist.map((t,i)=>`<label class="chk"><input type="checkbox" class="ck" ${final?'checked disabled':''}><span>${esc(t)}</span></label>`).join('')}</div>

    ${final ? '' : `<h3>Approve final</h3>
    <input type="text" id="who" placeholder="Your name (approver)" autocomplete="name">
    <div class="row"><button id="approveBtn" class="go" disabled>Approve Final</button>
    <span class="muted" id="approveHint">Tick every checklist item and enter your name.</span></div>`}

    <h3>Next improvements (reply with one to approve it)</h3>
    <ol class="plain">${r.improvements.map(x=>`<li>${esc(x)}</li>`).join('')}</ol>`;
  paintPipe(r.stages);
  if(!final){
    const upd = () => { const all=[...document.querySelectorAll('.ck')].every(x=>x.checked);
      $('#approveBtn').disabled = !(all && $('#who').value.trim()); };
    document.querySelectorAll('.ck').forEach(x=>x.addEventListener('change',upd));
    $('#who').addEventListener('input',upd);
    $('#approveBtn').addEventListener('click',approve);
  }
}

async function approve(){
  $('#approveBtn').disabled = true;
  const res = await fetch('/api/approve',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({version:current.version,approved_by:$('#who').value.trim(),confirm:true})});
  const d = await res.json();
  if(d.ok){ renderReport(d.report); loadState(); }
  else { $('#approveHint').textContent = d.error; $('#approveHint').className='err'; $('#approveBtn').disabled=false; }
}

async function run(){
  const fd = new FormData();
  fd.append('text', $('#text').value);
  fd.append('fresh', $('#fresh').checked ? '1' : '0');
  for(const f of $('#files').files) fd.append('files', f);
  $('#runBtn').disabled = true; $('#msg').textContent = 'Running...'; $('#msg').className='muted';
  const res = await fetch('/api/run',{method:'POST',body:fd});
  const d = await res.json();
  if(!res.ok){ $('#msg').textContent = d.error; $('#msg').className='err'; $('#runBtn').disabled=false; return; }
  $('#result').hidden = true;
  polling = setInterval(async ()=>{
    const j = await (await fetch('/api/job/'+d.job_id)).json();
    paintPipe(j.stages);
    if(j.state === 'running') return;
    clearInterval(polling); $('#runBtn').disabled = false;
    if(j.state === 'error'){ $('#msg').textContent = 'Crashed: '+j.error; $('#msg').className='err'; return; }
    if(!j.report.ok){ $('#msg').textContent = j.report.error; $('#msg').className='err'; return; }
    $('#msg').textContent = 'Done. Paused at NEEDS_HUMAN_APPROVAL.'; $('#msg').className='muted';
    renderReport(j.report); loadState();
    $('#result').scrollIntoView({behavior:'smooth'});
  }, 600);
}

async function loadState(first){
  const s = await (await fetch('/api/state')).json();
  STAGES = s.stages;
  $('#chipKey').textContent = s.gemini_key ? 'Gemini key: set ('+s.model+')' : 'Gemini key: missing (CSV/TXT still work)';
  $('#chipKey').className = 'chip ' + (s.gemini_key ? 'ok' : 'bad');
  $('#chipVer').textContent = s.version ? 'Latest: v'+s.version : 'No runs yet';
  $('#histCard').hidden = !s.history.length;
  $('#hist').innerHTML = [...s.history].reverse().map(h=>`<div class="hist"><span><b>v${h.version}</b> &middot; ${esc(h.created)} &middot; ${h.posted} posted, ${h.held} held</span>
    <span><a href="/download/${esc(h.file)}" style="color:var(--active)">${esc(h.file)}</a>${h.final_file?' &middot; <a href="/download/'+esc(h.final_file)+'" style="color:var(--done)">FINAL</a>':' &middot; <span class="muted">needs approval</span>'}</span></div>`).join('');
  if(first){ paintPipe(s.last_report ? s.last_report.stages : null); if(s.last_report) renderReport(s.last_report); }
}
$('#runBtn').addEventListener('click', run);
loadState(true);
</script>
</body></html>
"""

if __name__ == "__main__":
    engine.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"SIEGE Studio on http://0.0.0.0:{PORT}  (port 8080 untouched)")
    app.run(host="0.0.0.0", port=PORT, debug=False, threaded=True)
