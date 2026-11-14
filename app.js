import {
  Anthropic, Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
} from "./vendor.js?v=1";

const $ = (id) => document.getElementById(id);
const el = {
  title: $("title"), when: $("when"), place: $("place"), people: $("people"), author: $("author"), terms: $("terms"),
  recBtn: $("recBtn"), timer: $("timer"), liveMark: $("liveMark"), interim: $("interim"), transcript: $("transcript"),
  resumeBtn: $("resumeBtn"), clearBtn: $("clearBtn"), count: $("count"), style: $("style"),
  goBtn: $("goBtn"), stopBtn: $("stopBtn"), status: $("status"), sheet: $("sheet"), raw: $("raw"),
  editBtn: $("editBtn"), copyBtn: $("copyBtn"), shareBtn: $("shareBtn"), dlBtn: $("dlBtn"),
  settings: $("settings"), settingsBtn: $("settingsBtn"), apiKey: $("apiKey"), model: $("model"), forgetBtn: $("forgetBtn"),
  recHint: $("recHint"), keyBanner: $("keyBanner"), keyStatus: $("keyStatus"), storageBanner: $("storageBanner"), settingsForm: $("settingsForm"), standaloneBanner: $("standaloneBanner"), selfUrl: $("selfUrl"),
};

// ---------------------------------------------------------------- storage
const KEY = "meeting-minutes-app-v1";
const KEY_API = "meeting-minutes-app-key";
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
  del(k) { try { localStorage.removeItem(k); } catch {} },
};

let minutes = "";      // markdown subset
let docBlob = null;    // pre-built .docx (iOS share needs it ready inside the tap)

function persist() {
  store.set(KEY, {
    title: el.title.value, when: el.when.value, place: el.place.value, people: el.people.value,
    author: el.author.value, terms: el.terms.value, transcript: el.transcript.value,
    style: el.style.value, minutes,
  });
}

// ---------------------------------------------------------------- helpers
const pad = (n) => String(n).padStart(2, "0");
function nowLocal() {
  const d = new Date(); d.setMinutes(Math.floor(d.getMinutes() / 5) * 5);
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function fmtWhen(v) {
  if (!v) return "";
  const d = new Date(v); if (isNaN(d)) return v;
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日（${"日月火水木金土"[d.getDay()]}）${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function setStatus(msg, err = false) { el.status.textContent = msg || ""; el.status.classList.toggle("err", err); }
function updateCount() { el.count.textContent = el.transcript.value.replace(/\s/g, "").length.toLocaleString("ja-JP") + "字"; }

// ---------------------------------------------------------------- restore
const saved = store.get(KEY, null);
if (saved) {
  for (const k of ["title", "when", "place", "people", "author", "terms", "transcript", "style"]) if (saved[k]) el[k].value = saved[k];
  if (saved.minutes) setMinutes(saved.minutes);
}
if (!el.when.value) el.when.value = nowLocal();
updateCount();
for (const k of ["title", "when", "place", "people", "author", "terms", "transcript", "style"]) {
  el[k].addEventListener("input", () => { if (k === "transcript") updateCount(); persist(); if (minutes) buildDoc(); });
}

const isStandalone = window.navigator.standalone === true || matchMedia("(display-mode: standalone)").matches;
if (isStandalone) { el.standaloneBanner.hidden = false; el.selfUrl.textContent = location.href.split("#")[0]; }

// ---------------------------------------------------------------- settings
// localStorage が使えるか（Safari の「すべての Cookie をブロック」やプライベートブラウズ等で使えないことがある）
const storageOK = (() => {
  try { const k = "__probe__"; localStorage.setItem(k, "1"); const ok = localStorage.getItem(k) === "1"; localStorage.removeItem(k); return ok; }
  catch { return false; }
})();
try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch {}

let memKey = "";                                   // 保存できない環境でも、開いている間は使えるように
const getKey = () => store.get(KEY_API, "") || memKey;
function saveKey(v) {
  memKey = v;
  if (v) store.set(KEY_API, v); else store.del(KEY_API);
  refreshKeyBanner();
}
function refreshKeyBanner() {
  const k = getKey();
  el.keyBanner.hidden = !!k;
  if (!el.keyStatus) return;           // 古い画面（キャッシュ）と組み合わさっても止まらないように
  if (!storageOK) {
    el.keyStatus.textContent = "この画面ではキーを保存できません。閉じると消えるため、次回また入力が必要です。下の「保存できないとき」を確認してください。";
    el.keyStatus.className = "keystatus err";
  } else if (k && store.get(KEY_API, "")) {
    el.keyStatus.textContent = `保存済み（末尾 …${k.slice(-4)}）。次回から入力は不要です。`;
    el.keyStatus.className = "keystatus ok";
  } else {
    el.keyStatus.textContent = "まだ登録されていません。";
    el.keyStatus.className = "keystatus";
  }
}
if (!storageOK && el.storageBanner) el.storageBanner.hidden = false;
el.apiKey.value = getKey();
el.model.value = store.get(KEY + "-model", "claude-opus-5-5");
refreshKeyBanner();
el.settingsBtn.addEventListener("click", () => el.settings.showModal());
el.keyBanner.addEventListener("click", () => el.settings.showModal());
el.apiKey.addEventListener("input", () => saveKey(el.apiKey.value.trim()));
el.apiKey.addEventListener("change", () => saveKey(el.apiKey.value.trim()));
el.model.addEventListener("change", () => store.set(KEY + "-model", el.model.value));
el.settings.addEventListener("close", () => saveKey(el.apiKey.value.trim()));
if (el.settingsForm) el.settingsForm.addEventListener("submit", () => saveKey(el.apiKey.value.trim()));
el.forgetBtn.addEventListener("click", () => { el.apiKey.value = ""; saveKey(""); });

// ---------------------------------------------------------------- recording (Web Speech API)
const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
let rec = null, recording = false, pending = "", startedAt = 0, elapsedBefore = 0, tick = null, segTimer = null, wakeLock = null;
const SEGMENT_MS = 50_000; // iPhone では長く続けると確定しなくなるため、約50秒ごとに区切って再開する

if (!SR) {
  el.recBtn.disabled = true;
  el.recHint.textContent = "このブラウザでは音声認識が使えません。iPhone の Safari で開くか、ボイスメモの文字起こしを下の欄に貼り付けてください。";
}

function appendLine(text) {
  const t = text.trim(); if (!t) return;
  const cur = el.transcript.value;
  el.transcript.value = cur + (cur && !cur.endsWith("\n") ? "\n" : "") + t;
  el.transcript.scrollTop = el.transcript.scrollHeight;
  updateCount(); persist();
}

function startRecognizer() {
  rec = new SR();
  rec.lang = "ja-JP";
  rec.continuous = true;
  rec.interimResults = true;
  rec.onresult = (e) => {
    let interim = "";
    for (let i = e.resultIndex; i < e.results.length; i++) {
      const r = e.results[i];
      if (r.isFinal) { appendLine(r[0].transcript); pending = ""; }
      else interim += r[0].transcript;
    }
    pending = interim;
    el.interim.textContent = interim;
  };
  rec.onerror = (e) => {
    if (e.error === "not-allowed" || e.error === "service-not-allowed") {
      stopRecording();
      setRecHint("マイクまたは音声認識が許可されていません。iPhone の「設定」→「Safari」→「マイク」と、「設定」→「プライバシーとセキュリティ」→「音声認識」を確認してください。", true);
    }
  };
  rec.onend = () => {
    if (pending) { appendLine(pending); pending = ""; el.interim.textContent = ""; }
    clearTimeout(segTimer);
    if (!recording) return;
    setTimeout(() => {
      if (!recording) return;
      try { startRecognizer(); }
      catch { el.resumeBtn.hidden = false; setRecHint("文字起こしが止まりました。「認識を再開」を押してください。", true); }
    }, 200);
  };
  rec.start();
  segTimer = setTimeout(() => { try { rec && rec.stop(); } catch {} }, SEGMENT_MS);
}

function setRecHint(msg, warn) { el.recHint.textContent = msg; el.recHint.style.color = warn ? "var(--rec)" : ""; }

async function lockScreen() {
  try { if ("wakeLock" in navigator) wakeLock = await navigator.wakeLock.request("screen"); } catch {}
}
document.addEventListener("visibilitychange", () => {
  if (!recording) return;
  if (document.visibilityState === "visible") {
    lockScreen();
    appendLine("（この間、画面が消えていたため文字起こしが中断）");
    el.resumeBtn.hidden = false;
    setRecHint("画面が消えていたため文字起こしが止まった可能性があります。続けるには「認識を再開」を押してください。", true);
  }
});

function startRecording() {
  if (!SR) return;
  recording = true;
  pending = "";
  if (!el.transcript.value.trim()) {
    appendLine(`【録音開始 ${fmtWhen(nowLocal()) || ""}】`);
  } else {
    appendLine("【録音再開】");
  }
  try { startRecognizer(); } catch (err) { recording = false; setRecHint("録音を開始できませんでした。ページを読み込み直してください。", true); return; }
  startedAt = Date.now();
  tick = setInterval(renderTimer, 500);
  lockScreen();
  el.recBtn.classList.add("on"); el.recBtn.textContent = "停止"; el.recBtn.setAttribute("aria-pressed", "true");
  el.liveMark.hidden = false; el.transcript.readOnly = true; el.resumeBtn.hidden = true;
  setRecHint("録音中です。画面をつけたまま、Safari を開いたままにしてください。");
}

function stopRecording() {
  recording = false;
  clearTimeout(segTimer);
  try { rec && rec.stop(); } catch {}
  elapsedBefore += startedAt ? Date.now() - startedAt : 0; startedAt = 0;
  clearInterval(tick); renderTimer();
  try { wakeLock && wakeLock.release(); } catch {} wakeLock = null;
  el.recBtn.classList.remove("on"); el.recBtn.textContent = "録音開始"; el.recBtn.setAttribute("aria-pressed", "false");
  el.liveMark.hidden = true; el.transcript.readOnly = false; el.resumeBtn.hidden = true;
  setTimeout(() => { if (pending) { appendLine(pending); pending = ""; } el.interim.textContent = ""; }, 600);
  setRecHint("録音を停止しました。文字起こしを確認・修正してから、3 の「議事録を作成」を押してください。");
}

function renderTimer() {
  const ms = elapsedBefore + (startedAt ? Date.now() - startedAt : 0);
  const s = Math.floor(ms / 1000);
  el.timer.textContent = `${pad(Math.floor(s / 3600))}:${pad(Math.floor(s / 60) % 60)}:${pad(s % 60)}`;
}

el.recBtn.addEventListener("click", () => (recording ? stopRecording() : startRecording()));
el.resumeBtn.addEventListener("click", () => {
  el.resumeBtn.hidden = true;
  try { rec && rec.abort(); } catch {}
  try { startRecognizer(); setRecHint("録音中です。画面をつけたまま、Safari を開いたままにしてください。"); }
  catch { setRecHint("再開できませんでした。いったん「停止」してから、もう一度「録音開始」を押してください。", true); }
});
el.clearBtn.addEventListener("click", () => {
  if (recording) return;
  if (el.clearBtn.dataset.confirm !== "1") {
    el.clearBtn.dataset.confirm = "1"; el.clearBtn.textContent = "本当に消す？（もう一度押す）";
    setTimeout(() => { el.clearBtn.dataset.confirm = ""; el.clearBtn.textContent = "文字起こしを消す"; }, 4000);
    return;
  }
  el.clearBtn.dataset.confirm = ""; el.clearBtn.textContent = "文字起こしを消す";
  el.transcript.value = ""; elapsedBefore = 0; renderTimer(); updateCount(); persist();
});

// ---------------------------------------------------------------- prompt
function buildPrompt() {
  const st = el.style.value;
  const terms = el.terms.value.split("\n").map((s) => s.trim()).filter(Boolean);
  const meta = [
    `会議名：${el.title.value.trim() || "（未記入。内容から短い会議名を付け、末尾に［要確認］）"}`,
    `日時：${fmtWhen(el.when.value) || "（未記入）"}`,
    `場所：${el.place.value.trim() || "（未記入）"}`,
    `出席者：${el.people.value.trim() || "（未記入）"}`,
  ].join("\n");
  const detail = {
    brief: "「議事内容」の章は設けず、「議題」の次に「要点」の章を置き、会議全体の要点を3〜6項目の箇条書きにする。",
    std: "「議事内容」は議題ごとに ### 見出しを立て、論点と結論を2〜5項目の箇条書きで簡潔にまとめる。",
    full: "「議事内容」は議題ごとに ### 見出しを立て、論点・主な意見・結論を漏れなく箇条書きにする。発言者が文字起こしで明示されている場合のみ「（田中）」のように末尾に記す。",
  }[st];
  return `あなたは日本語の会議の議事録を作成する担当者です。以下の「文字起こし」から議事録を作成してください。

# 会議情報（利用者が入力）
${meta}

# 作成ルール
- 文字起こしは iPhone の音声認識の結果で、誤変換・聞き間違い・句読点の欠落・話者の区別なしを含む前提で読む。【】で囲んだ行は録音の区切りの記録であり、議事録には含めない。
- 意味が取れない箇所や誤認識と思われる語は、推測で補ったり直したりせず、その語や文の直後に［要確認］と付ける。そのような箇所は最後の「要確認事項」にも、元の言い回しを「」で引用して列挙する。
${terms.length ? `- 次の語は正しい表記である。音が近い誤変換と判断できる場合に限り、この表記に直してよい：${terms.join("、")}\n` : ""}- 文字起こしに無い事実（決定・担当者・期限・数値・固有名詞）を作らない。担当者や期限が述べられていなければ「未定」と書く。
- 会議情報が「（未記入）」の項目は、文字起こしから推定して埋めず「（未記入）」のまま出す。
- 「えー」「あの」などの言いよどみ、雑談、繰り返しは省く。
- 文体は「である」調。箇条書きは体言止めか簡潔な文にする。
- ${detail}
- 該当する内容が無い章は「- なし」と書く。

# 出力形式（この形だけを出力する。前置き・後書き・コードブロック・表・太字は使わない）
# 会議名
- 日時：…
- 場所：…
- 出席者：…

## 議題
- …

${st === "brief" ? "## 要点\n- …" : "## 議事内容\n### 議題の名前\n- …"}

## 決定事項
- …

## 保留・継続検討
- …

## 宿題（アクションアイテム）
- 内容／担当：…／期限：…

## 要確認事項
- …

# 文字起こし
<<<
${el.transcript.value.trim()}
>>>`;
}

// ---------------------------------------------------------------- Claude
let stream = null;
const PRICES = { "claude-opus-5-5": [4, 20], "claude-sonnet-5-5": [2, 10] }; // USD / 1M tokens

async function generate() {
  const apiKey = getKey();
  if (!apiKey) { el.settings.showModal(); setStatus("先に API キーを登録してください。", true); return; }
  if (recording) stopRecording();
  const text = el.transcript.value.trim();
  if (text.replace(/【[^】]*】/g, "").trim().length < 30) { setStatus("文字起こしが短すぎます（30字以上必要です）。", true); return; }
  if (!el.raw.hidden) toggleEdit();

  const model = el.model.value;
  const client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true, maxRetries: 2 });
  el.goBtn.disabled = true; el.stopBtn.hidden = false; el.editBtn.disabled = true; el.copyBtn.disabled = true;
  el.shareBtn.disabled = true; el.dlBtn.disabled = true;
  setStatus("Claude が文字起こしを読んでいます。長い会議では1〜3分かかります。");
  el.sheet.innerHTML = '<p class="empty">作成中…</p>';

  try {
    stream = client.beta.messages.stream({
      model,
      max_tokens: 32000,
      betas: ["server-side-fallback-2026-07-01"],
      fallbacks: "default",
      output_config: { effort: "medium" },
      messages: [{ role: "user", content: buildPrompt() }],
    });
    stream.on("text", (_delta, snapshot) => { el.sheet.innerHTML = render(snapshot); });
    const msg = await stream.finalMessage();
    if (msg.stop_reason === "refusal") {
      setStatus("この内容では議事録を作成できませんでした。", true);
      el.sheet.innerHTML = '<p class="empty">作成できませんでした。</p>';
      return;
    }
    const out = msg.content.filter((b) => b.type === "text").map((b) => b.text).join("").trim();
    setMinutes(out);
    persist();
    const [pin, pout] = PRICES[model] || PRICES["claude-opus-5-5"];
    const usd = (msg.usage.input_tokens * pin + msg.usage.output_tokens * pout) / 1e6;
    setStatus(msg.stop_reason === "max_tokens"
      ? "途中で切れました。「要点のみ」にするか、文字起こしを分けて作成してください。"
      : `作成しました（料金の目安：約 ${usd < 0.01 ? "1円未満" : "$" + usd.toFixed(2)}）。［要確認］の箇所を確認してください。`,
      msg.stop_reason === "max_tokens");
  } catch (err) {
    if (err instanceof Anthropic.APIUserAbortError) {
      setStatus("中止しました。");
    } else if (err instanceof Anthropic.AuthenticationError) {
      setStatus("API キーが正しくありません。「設定」で確認してください。", true);
    } else if (err instanceof Anthropic.PermissionDeniedError) {
      setStatus("この API キーではこのモデルを使えません。Anthropic のコンソールで権限を確認してください。", true);
    } else if (err instanceof Anthropic.RateLimitError) {
      setStatus("利用回数の上限に達しました。少し待ってからもう一度お試しください。", true);
    } else if (err instanceof Anthropic.BadRequestError) {
      setStatus(/credit|balance/i.test(err.message) ? "API の残高が不足しています。Anthropic のコンソールでクレジットを追加してください。" : "リクエストが受け付けられませんでした：" + err.message, true);
    } else if (err instanceof Anthropic.APIConnectionError) {
      setStatus("通信できませんでした。電波の良い場所でもう一度お試しください。", true);
    } else if (err instanceof Anthropic.APIError) {
      setStatus("Claude 側でエラーが起きました。時間をおいてお試しください。", true);
    } else {
      setStatus("作成できませんでした：" + (err && err.message ? err.message : String(err)), true);
    }
    if (minutes) el.sheet.innerHTML = render(minutes); else el.sheet.innerHTML = '<p class="empty">ここに議事録が表示されます。</p>';
  } finally {
    stream = null;
    el.goBtn.disabled = false; el.stopBtn.hidden = true;
    const has = !!minutes;
    el.editBtn.disabled = !has; el.copyBtn.disabled = !has;
  }
}
el.goBtn.addEventListener("click", generate);
el.stopBtn.addEventListener("click", () => stream && stream.abort());

// ---------------------------------------------------------------- render
const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const CHK = /[［\[]要確認[］\]]/g;
const CHK1 = /^[［\[]要確認[］\]]$/;
const inline = (s) => esc(s).replace(CHK, '<span class="chk">［要確認］</span>');

function parse(md) {
  // → [{kind:"h1"|"h2"|"h3"|"li"|"p", level, text}]
  const out = [];
  for (const raw of md.replace(/\r/g, "").split("\n")) {
    const line = raw.replace(/\t/g, "  ");
    let m;
    if ((m = line.match(/^(#{1,3})\s+(.*)$/))) out.push({ kind: "h" + m[1].length, text: m[2].trim() });
    else if ((m = line.match(/^(\s*)[-*・]\s+(.*)$/))) out.push({ kind: "li", level: Math.min(2, Math.floor(m[1].length / 2)), text: m[2].trim() });
    else if (line.trim() && !/^```/.test(line.trim())) out.push({ kind: "p", text: line.trim() });
  }
  return out;
}

function render(md) {
  const html = []; let depth = -1;
  const close = (to) => { while (depth > to) { html.push("</li></ul>"); depth--; } };
  for (const b of parse(md)) {
    if (b.kind === "li") {
      if (b.level > depth) { while (depth < b.level) { html.push("<ul><li>"); depth++; } }
      else { close(b.level); html.push("</li><li>"); }
      html.push(inline(b.text));
    } else {
      close(-1);
      html.push(b.kind === "p" ? `<p>${inline(b.text)}</p>` : `<${b.kind}>${inline(b.text)}</${b.kind}>`);
    }
  }
  close(-1);
  return html.join("") || '<p class="empty">作成中…</p>';
}

function toPlain(md) {
  return parse(md).map((b) => {
    if (b.kind === "h1") return b.text;
    if (b.kind === "h2") return "\n■ " + b.text;
    if (b.kind === "h3") return "□ " + b.text;
    if (b.kind === "li") return "　".repeat(b.level) + "・" + b.text;
    return b.text;
  }).join("\n");
}

function setMinutes(md) {
  minutes = md;
  el.sheet.innerHTML = md ? render(md) : '<p class="empty">ここに議事録が表示されます。</p>';
  el.editBtn.disabled = !md; el.copyBtn.disabled = !md;
  buildDoc();
}

function toggleEdit() {
  if (el.raw.hidden) {
    el.raw.value = minutes; el.raw.hidden = false; el.sheet.hidden = true; el.editBtn.textContent = "編集を終える";
  } else {
    setMinutes(el.raw.value); persist(); el.raw.hidden = true; el.sheet.hidden = false; el.editBtn.textContent = "直接編集";
  }
}
el.editBtn.addEventListener("click", toggleEdit);
el.raw.addEventListener("input", () => { minutes = el.raw.value; persist(); buildDocSoon(); });

el.copyBtn.addEventListener("click", async () => {
  const plain = toPlain(el.raw.hidden ? minutes : el.raw.value);
  try { await navigator.clipboard.writeText(plain); setStatus("コピーしました。"); }
  catch { setStatus("コピーできませんでした。「直接編集」を開いて全選択してください。", true); }
});

// ---------------------------------------------------------------- Word (.docx)
const FONT = { ascii: "Yu Mincho", eastAsia: "游明朝", hAnsi: "Yu Mincho" };
const FONT_G = { ascii: "Yu Gothic", eastAsia: "游ゴシック", hAnsi: "Yu Gothic" };

function runs(text, opts = {}) {
  const parts = text.split(/([［\[]要確認[］\]])/);
  return parts.filter(Boolean).map((p) => CHK1.test(p)
    ? new TextRun({ text: "［要確認］", color: "C62828", bold: true, font: FONT_G, size: 18 })
    : new TextRun({ text: p, ...opts }));
}

function metaTable(rows) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "9AA7B6" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map(([k, v]) => new TableRow({
      children: [
        new TableCell({ width: { size: 18, type: WidthType.PERCENTAGE }, borders, shading: { type: ShadingType.CLEAR, color: "auto", fill: "E3E9F3" },
          children: [new Paragraph({ children: [new TextRun({ text: k, bold: true, font: FONT_G, size: 20 })] })] }),
        new TableCell({ width: { size: 82, type: WidthType.PERCENTAGE }, borders,
          children: [new Paragraph({ children: runs(v || "（未記入）", { size: 20 }) })] }),
      ],
    })),
  });
}

function makeDoc(md) {
  const blocks = parse(md);
  const children = [];
  let title = el.title.value.trim() || "議事録";
  const metaRows = [];
  let i = 0;
  if (blocks[0] && blocks[0].kind === "h1") { title = blocks[0].text.replace(CHK, "").trim() || title; i = 1; }
  // 冒頭の「- 日時：…」などは表にまとめる
  while (blocks[i] && blocks[i].kind === "li" && /^(日時|場所|出席者)[：:]/.test(blocks[i].text)) {
    const m = blocks[i].text.match(/^(日時|場所|出席者)[：:]\s*(.*)$/); metaRows.push([m[1], m[2].trim()]); i++;
  }
  if (!metaRows.length) metaRows.push(["日時", fmtWhen(el.when.value)], ["場所", el.place.value.trim()], ["出席者", el.people.value.trim()]);
  if (el.author.value.trim()) metaRows.push(["作成者", el.author.value.trim()]);

  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "議 事 録", font: FONT_G, size: 22, color: "5B6878" })] }));
  children.push(new Paragraph({ heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: runs(title, { bold: true, size: 32, font: FONT }) }));
  children.push(metaTable(metaRows));
  children.push(new Paragraph({ text: "" }));

  for (; i < blocks.length; i++) {
    const b = blocks[i];
    if (b.kind === "h1") children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 80 }, children: runs(b.text, { bold: true, size: 28, font: FONT_G }) }));
    else if (b.kind === "h2") children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 80 },
      border: { left: { style: BorderStyle.SINGLE, size: 24, color: "26457A", space: 6 } }, children: runs(b.text, { bold: true, size: 24, font: FONT_G }) }));
    else if (b.kind === "h3") children.push(new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 160, after: 40 }, children: runs("□ " + b.text, { bold: true, size: 22, font: FONT_G }) }));
    else if (b.kind === "li") children.push(new Paragraph({ bullet: { level: b.level }, spacing: { after: 40 }, children: runs(b.text, { size: 21 }) }));
    else children.push(new Paragraph({ spacing: { after: 60 }, children: runs(b.text, { size: 21 }) }));
  }

  children.push(new Paragraph({ text: "" }));
  children.push(new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "※ 音声認識の文字起こしをもとに Claude で作成。［要確認］は録音で確認すること。", size: 16, color: "5B6878", font: FONT_G })] }));

  return new Document({
    creator: el.author.value.trim() || "会議録音メモ",
    title,
    styles: {
      default: { document: { run: { font: FONT, size: 21, color: "1D2733" } } },
      paragraphStyles: [
        { id: "Title", name: "Title", basedOn: "Normal", run: { font: FONT, size: 32, bold: true, color: "1D2733" } },
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", run: { font: FONT_G, size: 28, bold: true, color: "1D2733" } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", run: { font: FONT_G, size: 24, bold: true, color: "26457A" } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", run: { font: FONT_G, size: 22, bold: true, color: "1D2733" } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1418, bottom: 1418, left: 1418, right: 1418 } } },
      children,
    }],
  });
}

function fileName() {
  const d = el.when.value ? el.when.value.slice(0, 10).replace(/-/g, "") : "";
  const t = (el.title.value.trim() || "議事録").replace(/[\\/:*?"<>|\s]+/g, "_").slice(0, 40);
  return `${d ? d + "_" : ""}${t}_議事録.docx`;
}

let buildSeq = 0, buildTimer = null;
async function buildDoc() {
  const seq = ++buildSeq;
  const md = el.raw.hidden ? minutes : el.raw.value;
  if (!md) { docBlob = null; el.shareBtn.disabled = el.dlBtn.disabled = true; return; }
  try {
    const blob = await Packer.toBlob(makeDoc(md));
    if (seq !== buildSeq) return;
    docBlob = blob;
    el.shareBtn.disabled = el.dlBtn.disabled = false;
  } catch (err) {
    console.error(err);
    setStatus("Word ファイルを作れませんでした：" + err.message, true);
  }
}
function buildDocSoon() { clearTimeout(buildTimer); buildTimer = setTimeout(buildDoc, 600); }

const DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
el.shareBtn.addEventListener("click", async () => {
  if (!docBlob) return;
  const file = new File([docBlob], fileName(), { type: DOCX_TYPE });
  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: fileName() }); }
    catch (err) { if (err && err.name !== "AbortError") setStatus("共有できませんでした。「ファイルに保存」をお使いください。", true); }
  } else {
    download();
  }
});
function download() {
  if (!docBlob) return;
  const url = URL.createObjectURL(new Blob([docBlob], { type: DOCX_TYPE }));
  const a = document.createElement("a");
  a.href = url; a.download = fileName(); a.hidden = true; document.body.appendChild(a); a.click();
  setTimeout(() => a.remove(), 1000);
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
  setStatus("Word ファイルを保存しました（iPhone では「ファイル」アプリの「ダウンロード」に入ります）。");
}
el.dlBtn.addEventListener("click", download);

window.addEventListener("beforeunload", (e) => { if (recording) { e.preventDefault(); e.returnValue = ""; } });
