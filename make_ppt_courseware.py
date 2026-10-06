# -*- coding: utf-8 -*-
"""
听力PPT课件生成器  —— 由切分音频(网页课件切的, 或本程序自切)直接拼出可播放的 pptx。

用法:
    python make_ppt_courseware.py [试卷文件夹] [--recut] [--legacy]
    python make_ppt_courseware.py --rebuild-template [成品.pptx]
    python make_ppt_courseware.py --help
        --recut   强制重切音频(换过 mp3 时用)
        --legacy  额外产出 tm.txt/yw.txt/音频文件 (喂老的 .pptm 宏路)
        --rebuild-template  用一份能播的成品 pptx 重建模板

不传文件夹就对"本脚本所在的文件夹"下手; 中文输出由脚本自己设好 UTF-8, 无需任何包装。

文件夹里需要:
    - 听力 mp3        (如  XXX_英语听力.mp3)
    - 试卷 docx       (含第 1~20 题的题干与选项)
    - 听力原文 docx   (含 Text 1 ... Text 10)

试卷排版容错: 题干/选项同段或各占一段都认; 选项标点 A. / A． / A。 都认;
              题号漏写(如只有题干没有"14.")时按位置自动补; 提示语「听第N段录音/材料」都认。

音频来源(依次尝试, 先用现成的):
    1. audio\\seg01.mp3 .. seg10.mp3        网页课件产物, 最新鲜
    2. 音频文件\\Text 01.mp3 .. Text 10.mp3   若存在则用(老素材)
    3. 都没有 -> 本程序自带切分引擎, 按提示音就地切开(需 numpy + ffmpeg)
       分割点缓存 <P>听力课件_seg.json, 可手工改好分割点后重跑
    加 --recut 强制重切(跳过 audio\\ 与 音频文件\\, 且忽略已缓存的 _seg.json)。

产出(写到同一个文件夹):
    <P>听力课件.pptx         P = mp3名 -> 试卷名 -> 文件夹名 (同 make_courseware.py 的命名规则)
    只有这一个。加 --legacy 才额外产出 tm.txt / yw.txt / 音频文件\\Text NN.mp3
    (供老的 听力PPT课件生成器.pptm 宏路继续用)。

模板: 同目录 听力课件模板.pptx —— 11 页, 无宏, 由宏产物另存而来。
      每页的音频对象名 Text NN、文本框 myTextTm0/myTextBox、首页 TextBox01 会被整体替换。

原理: 纯 OOXML 手术(替换 ppt/media/*.mp3 字节 + 重写三个文本框), 不依赖 PowerPoint/COM。
"""
import os, re, sys, io, json, shutil, zipfile, posixpath, subprocess
import xml.etree.ElementTree as ET
import docx

# 控制台中文: 自己把标准输出/错误设成 UTF-8。
# 这样直接 `python make_ppt_courseware.py` 就不会乱码, 不需要任何 .bat 包装;
# 输出被管道接走、或重定向到文件时同样有效。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:                  # 没有控制台(如 pythonw)或已被别的对象接管
        pass

try:
    import numpy as np
except ImportError:                # 没装 numpy 不影响已有 audio\\seg*.mp3 的路线, 只是不能自带切分
    np = None

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE   = os.path.join(SCRIPT_DIR, '听力课件模板.pptx')
TEMPLATE_BAK = os.path.join(SCRIPT_DIR, '听力课件模板-备份.pptx')
N_TEXT     = 10
R_ID       = '{http://schemas.openxmlformats.org/package/2006/relationships}'

# ----------------------------------------------------------------- 命名规则
_AUDIO_TAILS = ("_英语听力", "英语听力", "_听力音频", "听力音频", "_听力", "听力")


def _strip_audio_tail(stem):
    for tail in _AUDIO_TAILS:
        if stem.endswith(tail) and len(stem) > len(tail):
            return stem[:-len(tail)].rstrip("_- ")
    return stem


def pick_prefix(folder):
    """课件基名: 依次取 mp3 名 -> 试卷名 -> 文件夹名(先有先用)。"""
    names = os.listdir(folder)
    cands = []
    mps = [f for f in names if f.lower().endswith(".mp3")
           and "听力" in f and "课件" not in f]
    mps.sort(key=lambda f: (0 if f.startswith(os.path.basename(folder)) else 1, f))
    cands += [_strip_audio_tail(os.path.splitext(f)[0]) for f in mps]
    docs = [f for f in names if f.lower().endswith((".docx", ".doc"))
            and not f.startswith("~$") and "课件" not in f]
    papers = [f for f in docs if ("试卷" in f or "试题卷" in f) and "答案" not in f]
    papers.sort(key=lambda f: (0 if "英语试卷" in f else 1, f))
    cands += [os.path.splitext(f)[0] for f in papers]
    cands.append(os.path.basename(folder))
    cands = [c for c in cands if c and c not in ("听力", "英语听力")]
    return cands[0] if cands else os.path.basename(folder)


# ----------------------------------------------------------------- 素材发现
def _pick_docx(folder, keys, exclude=()):
    hits = []
    for f in sorted(os.listdir(folder)):
        if not f.lower().endswith((".docx", ".doc")) or f.startswith("~$"):
            continue
        if "课件" in f:
            continue
        if any(k in f for k in keys) and not any(k in f for k in exclude):
            hits.append(os.path.join(folder, f))
    return hits


def _paras(path):
    return [p.text for p in docx.Document(path).paragraphs]


def _audio_hint(folder):
    """文件夹里的听力 mp3 —— 用来判断 音频文件\\ 是否过期, 以及现场切分的输入。"""
    for f in sorted(os.listdir(folder)):
        if f.lower().endswith(".mp3") and "听力" in f and "课件" not in f:
            return os.path.join(folder, f)
    return None


# ----------------------------------------------------------------- 音频分段(内置切分引擎)
# 移植自 听力网页课件生成器/make_courseware.py, 让本工具自带切分能力, 不必再拷外部脚本。
# 全曲每段正文前都有一个提示音(2593Hz 纯音 ~0.4s + 2048Hz 纯音 ~1.3s)。
# 提示音把全曲切成区间, 每区间 = 引导语 + 正文第1遍 + 停顿 + 正文第2遍。
# 从后向前: 拿区间末尾的语音当模板做归一化互相关, 回头找同一段的第1遍。
SEG_SUFFIX = "听力课件_seg.json"
SR        = 16000
CH_NF     = 2048            # 提示音检测帧长
CH_HOP    = 320             # 20ms
CH_LO_T   = 0.45            # 纯音占比阈值
CH_MIN_1, CH_MIN_2 = 0.30, 0.70      # 两段纯音最短时长
ISL_WIN   = 0.02            # 语音岛帧长
ISL_DB    = 8               # 噪声底之上多少 dB 算语音
ISL_GAP   = 0.12            # 岛合并间隙
BLK_GAP   = 0.80            # 正文第1/2遍之间的停顿下限
TPL_SEC   = 5.0             # 末尾模板长度
MIN_CORR  = 0.85            # 判定"同一段读两遍"的互相关阈值
MIN_SIDE  = 7.0             # 两遍各自最短时长
SEG_PAD_S, SEG_PAD_E = 0.50, 0.50    # 段首前导 / 段尾留白


def read_json(p):
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def write_json(p, obj):
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


_FF = {}


def tool(name):
    """定位 ffmpeg / ffprobe: 环境变量 -> PATH -> 常见安装目录; 找不到返回 None。"""
    if name in _FF:
        return _FF[name]
    found = None
    env = os.environ.get({"ffmpeg": "FFMPEG", "ffprobe": "FFPROBE"}.get(name, ""))
    cands = ([env] if env else []) + [shutil.which(name)]
    for c in cands:
        if c and os.path.exists(c):
            found = c; break
    if not found:                                   # 常见兜底, 不报错
        for d in (r"C:\ffmpeg\bin", r"D:\Program Files\ffmpeg\bin"):
            p = os.path.join(d, name + (".exe" if os.name == "nt" else ""))
            if os.path.exists(p):
                found = p; break
    _FF[name] = found
    return found


def load_pcm(mp3):
    """整曲解码为 16kHz 单声道 float32(-1..1)。"""
    raw = subprocess.run([tool("ffmpeg"), "-v", "error", "-i", mp3,
                          "-ac", "1", "-ar", str(SR), "-f", "s16le", "-"],
                         capture_output=True).stdout
    return np.frombuffer(raw, dtype="<i2").astype(np.float32) / 32768.0


def _tonal(x):
    """逐帧算 2593Hz / 2048Hz 两个频带的"纯音度"和峰值频率。"""
    win = np.hanning(CH_NF)
    nfr = (len(x) - CH_NF) // CH_HOP
    f = np.fft.rfftfreq(CH_NF, 1 / SR)
    m1, m2 = (f >= 2450) & (f <= 2750), (f >= 1990) & (f <= 2110)
    f1, f2 = f[m1], f[m2]
    T1 = np.zeros(nfr, np.float32); F1 = np.zeros(nfr, np.float32)
    T2 = np.zeros(nfr, np.float32); F2 = np.zeros(nfr, np.float32)
    for s in range(0, nfr, 4000):
        e = min(s + 4000, nfr)
        idx = np.arange(CH_NF)[None, :] + CH_HOP * np.arange(s, e)[:, None]
        P = np.abs(np.fft.rfft(x[idx] * win, axis=1)) ** 2
        A, B = P[:, m1], P[:, m2]
        T1[s:e] = A.max(axis=1) / (A.sum(axis=1) + 1e-12); F1[s:e] = f1[A.argmax(axis=1)]
        T2[s:e] = B.max(axis=1) / (B.sum(axis=1) + 1e-12); F2[s:e] = f2[B.argmax(axis=1)]
    return T1, F1, T2, F2


def _runs(mask, min_dur):
    out, i, n = [], 0, len(mask)
    while i < n:
        if mask[i]:
            j = i
            while j < n and mask[j]:
                j += 1
            if (j - 1 - i) * CH_HOP / SR + CH_NF / SR >= min_dur:
                out.append((i * CH_HOP / SR, ((j - 1) * CH_HOP + CH_NF) / SR))
            i = j
        else:
            i += 1
    return out


def find_chimes(x):
    """提示音 = 2593Hz 纯音(>=0.30s) 紧接 2048Hz 纯音(>=0.70s)。"""
    T1, F1, T2, F2 = _tonal(x)
    e1 = _runs((T1 > CH_LO_T) & (F1 > 2570) & (F1 < 2615), CH_MIN_1)
    e2 = _runs((T2 > CH_LO_T) & (F2 > 2030) & (F2 < 2065), CH_MIN_2)
    out = []
    for a, b in e1:
        for c, d in e2:
            if -0.10 <= c - b <= 0.60:
                out.append((a, d)); break
    return out


def _islands(x):
    W = int(SR * ISL_WIN)
    n = len(x) // W
    db = 20 * np.log10(np.sqrt(np.mean(x[:n * W].reshape(n, W) ** 2, axis=1)) + 1e-9)
    thr = max(float(np.percentile(db, 10)) + ISL_DB, -55)
    raw, i = [], 0
    while i < n:
        if db[i] > thr:
            j = i
            while j < n and db[j] > thr:
                j += 1
            raw.append((i, j)); i = j
        else:
            i += 1
    mg = []
    for a, b in raw:
        if mg and (a - mg[-1][1]) * ISL_WIN < ISL_GAP:
            mg[-1] = (mg[-1][0], b)
        else:
            mg.append((a, b))
    return [(a * ISL_WIN, b * ISL_WIN) for a, b in mg]


def _xcorr(sig, tmpl):
    """归一化互相关: cc[i] = <sig[i:i+n], tmpl> / (|.|*|tmpl|)。"""
    n = len(tmpl)
    N = 1 << int(np.ceil(np.log2(len(sig) + n)))
    c = np.fft.irfft(np.fft.rfft(sig, N) * np.fft.rfft(tmpl[::-1], N), N)[n - 1:len(sig)]
    cs = np.concatenate([[0.0], np.cumsum(sig.astype(np.float64) ** 2)])
    e = cs[n:] - cs[:len(cs) - n]
    return c / np.sqrt(e * float(np.sum(tmpl.astype(np.float64) ** 2)) + 1e-12)


def _blocks(seg):
    out, cur = [], []
    for t in seg:
        if cur and t[0] - cur[-1][1] >= BLK_GAP:
            out.append(cur); cur = []
        cur.append(t)
    if cur:
        out.append(cur)
    return out


def find_first_reading(x, seg):
    """从后向前: 末段语音当模板, 回头找同一段的第一遍。返回 (起, 止, 相关系数)。"""
    bl = _blocks(seg)
    if len(bl) < 2:
        return None
    s0, s1 = seg[0][0], seg[-1][1]
    gi, gmax = -1, 0.0
    for i in range(len(bl) - 1):
        g = bl[i + 1][0][0] - bl[i][-1][1]
        if g > gmax:
            gmax, gi = g, i
    if not (1.2 <= gmax <= 6.0):                # 最大块间隙=两遍之间的停顿
        return None
    c2 = bl[gi + 1][0][0]                       # 第 2 遍起点
    if c2 - s0 < MIN_SIDE or s1 - c2 < MIN_SIDE:
        return None
    T = min(TPL_SEC, s1 - c2)
    t0 = s1 - T
    tmpl = x[int(t0 * SR):int(t0 * SR) + int(T * SR)]
    sig = x[int(s0 * SR):int(s1 * SR)]
    cc = _xcorr(sig, tmpl)
    lim = int((c2 - 0.30 - s0) * SR) - len(tmpl)   # 匹配窗必须在第 2 遍之前收尾
    if lim < 1:
        return None
    i = int(np.argmax(cc[:lim + 1]))
    if cc[i] < MIN_CORR:
        return None
    d = t0 - (s0 + i / SR)                      # 第 2 遍相对第 1 遍的偏移
    a, b = c2 - d, s1 - d
    if a < s0 - 1.0 or b > c2 - 0.2:
        return None
    return (a, b, float(cc[i]))


def detect_segments(x, n):
    """返回 n 段单遍朗读的 (起, 止) 列表; 判不出返回 None。"""
    dur = len(x) / SR
    ch = find_chimes(x)
    isl = _islands(x)
    spans = []
    for k, (cs, ce) in enumerate(ch):
        nxt = ch[k + 1][0] if k + 1 < len(ch) else dur
        seg = [t for t in isl if t[0] >= ce + 0.05 and t[1] <= nxt + 0.05]
        spans.append({"idx": k, "cs": cs, "seg": seg,
                      "rep": find_first_reading(x, seg)})
    paired = [sp for sp in spans if sp["rep"]]
    plain  = [sp for sp in spans if not sp["rep"]]
    print("提示音 %d 个 / 读两遍 %d 个 / 预期 %d 段" % (len(ch), len(paired), n))
    if len(paired) == n:
        chosen = paired                       # 每段都读两遍(联考、周测)
    elif (len(paired) == n // 2 and len(plain) >= n // 2
          and plain[0]["cs"] < paired[0]["cs"]):
        chosen = plain[:n // 2] + paired      # 第一节只读一遍(高考卷)
    else:
        for sp in spans:
            m = ("两遍 %.2f~%.2f r=%.3f" % (sp["rep"][0], sp["rep"][1], sp["rep"][2])
                 if sp["rep"] else "")
            print("   #%02d %8.2f  岛%2d  %s" % (sp["idx"], sp["cs"], len(sp["seg"]), m))
        return None
    out = []
    for sp in chosen:
        if sp["rep"]:
            a, b = sp["rep"][0], sp["rep"][1]
        else:
            if not sp["seg"]:
                return None
            a, b = sp["seg"][0][0], sp["seg"][-1][1]
        out.append((max(0.0, a - SEG_PAD_S), b + SEG_PAD_E))
    return out


def cut(mp3, seg_map, n, out_dir):
    """按 seg_map 把 mp3 切成 out_dir\\seg01..segNN.mp3。"""
    os.makedirs(out_dir, exist_ok=True)
    made = []
    for j in range(1, n + 1):
        if j not in seg_map:
            continue
        s, e = seg_map[j]
        out = os.path.join(out_dir, "seg%02d.mp3" % j)
        r = subprocess.run([tool("ffmpeg"), "-y", "-hide_banner", "-loglevel", "error",
            "-ss", "%.2f" % s, "-to", "%.2f" % e, "-i", mp3,
            "-map", "0:a:0", "-c:a", "libmp3lame", "-b:a", "64k", out],
            capture_output=True, text=True, errors="replace")
        if r.returncode != 0:
            raise SystemExit("[ERR] 切分失败 seg%d: %s" % (j, r.stderr[-500:]))
        made.append((j, out, s, e))
    return made


def pick_seg_map(mp3, n, seg_json, force=False):
    """返回 {段号: (起, 止)}。force=True 时忽略缓存, 一律重新检测。
    缓存 <P>听力课件_seg.json 不早于 mp3 时复用(可手工编辑分割点后重跑)。"""
    if not force and os.path.isfile(seg_json) \
       and os.path.getmtime(seg_json) >= os.path.getmtime(mp3) - 1:
        cached = read_json(seg_json)
        m = {}
        for k, v in (cached or {}).items():
            try:
                m[int(k)] = (float(v[0]), float(v[1]))
            except (TypeError, ValueError, IndexError, KeyError):
                m = {}; break
        if len(m) == n:
            print("[OK] 采用已有分割点: %s" % os.path.basename(seg_json))
            return m
    if np is None:
        raise SystemExit("[ERR] 自动切分需要 numpy(pip install numpy); "
                         "或先把切好的 audio\\seg01-10.mp3 放进本文件夹")
    if not tool("ffmpeg"):
        raise SystemExit("[ERR] 找不到 ffmpeg(装好并加入 PATH, 或放 C:\\ffmpeg\\bin); "
                         "或先把切好的 audio\\seg01-10.mp3 放进本文件夹")
    auto = detect_segments(load_pcm(mp3), n)
    if not auto:
        raise SystemExit("[ERR] 按提示音自动定位失败(音频段数与文本不符?)。\n"
                         "     可手工建 %s(格式 {\"段号\": [起点, 终点] 秒})后重跑"
                         % os.path.basename(seg_json))
    m = {j + 1: v for j, v in enumerate(auto)}
    write_json(seg_json, {str(k): [round(a, 2), round(b, 2)] for k, (a, b) in m.items()})
    print("[OK] 已保存分割点: %s" % os.path.basename(seg_json))
    return m


def ensure_segments(folder, P, recut=False):
    """按 audio -> 音频文件 -> 现场切 的优先序返回 seg01..seg10 的路径。"""
    def paths_in(sub, pat):
        d = os.path.join(folder, sub)
        ps = [os.path.join(d, pat % i) for i in range(1, N_TEXT + 1)]
        return ps if all(os.path.isfile(p) for p in ps) else None

    if not recut:
        audio = paths_in("audio", "seg%02d.mp3")
        if audio:
            print("[OK] 音频来源: audio\\seg01-10.mp3")
            mp3 = _audio_hint(folder)
            if mp3 and os.path.getmtime(mp3) > max(os.path.getmtime(p) for p in audio) + 1:
                print("[!!] 但它比 %s 旧 —— 换过听力音频就加 --recut 重切"
                      % os.path.basename(mp3))
            return audio
        legacy = paths_in("音频文件", "Text %02d.mp3")
        if legacy:
            mp3 = _audio_hint(folder)
            if mp3 and os.path.getmtime(mp3) > max(os.path.getmtime(p) for p in legacy) + 1:
                print("[..] 音频文件\\ 比 %s 旧 -> 重切" % os.path.basename(mp3))
            else:
                print("[OK] 音频来源: 音频文件\\Text 01-10.mp3")
                return legacy

    mp3 = _audio_hint(folder)
    if not mp3:
        raise SystemExit("[ERR] 找不到听力 mp3(文件名含「听力」), 无法切分")
    print("[..] 现场切分音频: %s" % os.path.basename(mp3))
    seg_map = pick_seg_map(mp3, N_TEXT, os.path.join(folder, P + SEG_SUFFIX), force=recut)
    cut(mp3, seg_map, N_TEXT, os.path.join(folder, "audio"))
    audio = paths_in("audio", "seg%02d.mp3")
    if not audio:
        raise SystemExit("[ERR] 切分后仍未见 audio\\seg01-10.mp3")
    print("[OK] 切分完成 -> audio\\seg01-10.mp3")
    return audio


# ----------------------------------------------------------------- 试卷/原文解析
# 选项标记: A. / A． / A。 / A.. , 必须在行首、空白之后或中文标点之后
# (免得把 "…text.B…" 里紧跟在字母后面的那个字母当成选项)
_OPT = re.compile(r"(?:(?<=[\s。．；，、！？：])|^)([ABC])\s*[.．。]{1,2}\s*", re.M)
# 兜底: 有的卷子把选项点号漏了 / 误打成逗号("B The woman…", "A, Its patient…"),
# 只在严格式凑不齐 3 项时才用
_OPT_LOOSE = re.compile(r"(?:(?<=[\s。．；，、！？：])|^)([ABC])\s*"
                        r"(?:[.．。]{1,2}\s*|[,，、]\s*|\s+)", re.M)
_WS = re.compile(r"[ \t\u3000]+")
_QNUM = re.compile(r"^\d{1,2}\s*[.．、]")                      # 新题号 -> 选项块到此结束
_INSTR = re.compile(r"^(?:听第\s*\d+\s*段|听下面|第[一二三四五六]节|"
                    r"第[一二三四五六]部分|Text\s*\d+)")


def split_options(text, loose=False):
    """把 'A. ..  B. ..  C..Six' 或分行的 A./B./C. 拆成 ['A. ..','B. ..','C. Six']。
    容错: 全角句点、间距、Tab、选项各占一段(高考真题排版); loose=True 时容忍漏掉点号。"""
    s = text.strip()
    if not s:
        return []
    marks = list((_OPT_LOOSE if loose else _OPT).finditer(s))
    out = []
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(s)
        body = _WS.sub(" ", s[m.end():end].replace("\n", " ")).strip()
        out.append("%s. %s" % (m.group(1), body))
    return out


def collect_options(paras, i, max_lines=8):
    """第 i 段(题干)之后收集选项: 选项可同段, 也可 A./B./C. 各占一段。"""
    buf = []
    for j in range(i + 1, min(i + 1 + max_lines, len(paras))):
        t = paras[j].strip()
        if not t:
            continue
        if _QNUM.match(t) or _INSTR.match(t):
            break                                   # 下一题 / 下一段提示语
        buf.append(t)
        if len(split_options("\n".join(buf))) >= 3:
            break
    if not buf:
        return []
    txt = "\n".join(buf)
    opts = split_options(txt)
    if len(opts) != 3:                              # 兜底: 漏点号的排版
        loose = split_options(txt, loose=True)
        if len(loose) == 3 and [o[0] for o in loose] == ["A", "B", "C"]:
            opts = loose
    return opts


def scan_questions(paras):
    """听力部分 1~20 题 -> {题号: (题干, [选项...])}"""
    QA, pos, orphans = {}, {}, []
    for i, raw in enumerate(paras):
        t = raw.strip()
        if t.startswith("第二部分"):
            break
        m = re.match(r"^(\d{1,2})\s*[.．、]\s*(.*)$", t)
        if m:
            no = int(m.group(1))
            if not 1 <= no <= 20 or no in QA:
                continue
            stem = m.group(2).strip()
            opts = collect_options(paras, i) if stem else []
            if len(opts) != 3:
                continue    # 不是听力题(如"注意事项"里的 "1. 答题前…"), 末尾统一校验 1~20
            QA[no] = (stem, opts); pos[no] = i
            continue
        # 没带题号的裸题干(如 2025 高考卷漏了 "14."), 只要后面紧跟 3 个选项就先记下
        if len(t) >= 4 and not _INSTR.match(t) and not split_options(t):
            opts = collect_options(paras, i)
            if len(opts) == 3:
                orphans.append((i, t, opts))

    miss = [n for n in range(1, 21) if n not in QA]
    if miss and len(orphans) == len(miss):          # 按位置就近补题号: 前后必须是 n-1 / n+1
        for i, stem, opts in orphans:
            before = max([n for n in pos if pos[n] < i], default=0)
            after = min([n for n in pos if pos[n] > i], default=21)
            hit = [n for n in miss if n not in QA and before == n - 1 and after == n + 1]
            if len(hit) == 1:
                QA[hit[0]] = (stem, opts); pos[hit[0]] = i
                print("[..] 第%d题题干原文缺题号, 已按位置补上: %s" % (hit[0], stem[:40]))
    if sorted(QA) != list(range(1, 21)):
        lack = [n for n in range(1, 21) if n not in QA]
        raise SystemExit("[ERR] 听力题只认出 %d 道, 缺第 %s 题 (应为 1~20)\n"
                         "     多半是试卷 docx 里该题的题号或选项格式不同, 手工补一下再跑"
                         % (len(QA), "、".join(str(n) for n in lack)))
    return QA


def question_groups(paras):
    """题号 -> Text 段号: 第1~5段各1题; 第6~10段由『听第N段录音/材料,回答第X至Y题』定。"""
    groups = {t: [t] for t in range(1, 6)}
    for raw in paras:
        m = re.search(r"听第\s*(\d+)\s*段(?:录音|材料|对话|独白|短文)"
                      r"[，,、\s]*回答第\s*[:：、]?\s*([\d、至\-–~和及\s]+)题", raw)
        if not m:
            continue
        tn = int(m.group(1))
        nums = [int(x) for x in re.findall(r"\d+", m.group(2))]
        if not (6 <= tn <= N_TEXT) or not nums:
            continue
        groups[tn] = nums if len(nums) == 1 else list(range(nums[0], nums[-1] + 1))
    flat = [q for t in range(1, N_TEXT + 1) for q in groups.get(t, [])]
    if flat != list(range(1, 21)):
        lack = [n for n in range(1, 21) if n not in flat]
        raise SystemExit("[ERR] 题号->Text 分组不对, 没归到任何一段的是第 %s 题\n"
                         "     检查试卷里的『听第N段录音/材料, 回答第X至Y题』那句话"
                         % "、".join(str(n) for n in lack))
    return groups


def parse_transcripts(path):
    """听力原文 docx -> {Text 段号: [行...]}"""
    tx, cur = {}, None
    for p in docx.Document(path).paragraphs:
        t = p.text.strip()
        if not t:
            continue
        m = re.match(r"^Text\s+(\d+)\b", t)      # 标题行可能带时长杂质, 如 "Text 3 2'02''"
        if m and 1 <= int(m.group(1)) <= N_TEXT:
            cur = int(m.group(1))
            tx[cur] = []
            continue
        if cur is None:
            continue
        if re.match(r"^[—\-]\s*\d+\s*[—\-]$", t):      # 页码残渣
            continue
        tx[cur].append(p.text.rstrip())
    if sorted(tx) != list(range(1, N_TEXT + 1)):
        raise SystemExit("[ERR] 听力原文只认出 %s" % sorted(tx))
    return tx


def build_lines(groups, QA, tx):
    tm_parts, yw_parts = [], []
    for t in range(1, N_TEXT + 1):
        ls = []
        for q in groups[t]:
            stem, opts = QA[q]
            ls += ["%d. %s" % (q, stem), "\t".join(opts)]
        tm_parts.append(ls)
        yw_parts.append(["%d.%s" % (t, tx[t][0])] + ["   " + x for x in tx[t][1:]])
    return tm_parts, yw_parts


# ----------------------------------------------------------------- 模板体检
def _audio_slot(z, slide_name):
    """某页音频对象 -> (段号, 媒体部件名), 如 (7, 'ppt/media/media7.mp3'); 缺一不可。"""
    n = int(re.search(r"(\d+)", os.path.basename(slide_name)).group(1))
    try:
        rels = z.read("ppt/slides/_rels/slide%d.xml.rels" % n)
    except KeyError:
        return None, None
    xml = z.read(slide_name).decode("utf-8")
    for m in re.finditer(r"<p:pic[ >].*?</p:pic>", xml, re.S):
        blk = m.group(0)
        if "<a:audioFile" not in blk:
            continue
        nm = re.search(r'<p:cNvPr[^>]*name="([^"]*)"', blk)
        d = re.search(r"(\d+)", nm.group(1)) if nm else None
        rid = re.search(r'<a:audioFile[^>]*r:link="([^"]*)"', blk)
        tgt = rel_target(rels, rid.group(1)) if rid else None
        part = posixpath.normpath(posixpath.join("ppt/slides", tgt)).lstrip("/") if tgt else None
        return (int(d.group(1)) if d else None), part
    return None, None


def check_template(path):
    """模板结构体检; 返回问题清单(空 = 合格)。模板只给程序读, 别用 PowerPoint 打开编辑。"""
    if not os.path.isfile(path):
        return ["模板不存在: %s" % os.path.basename(path)]
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return ["打不开(可能损坏/正在被占用): %s" % e]
    names = z.namelist()
    slides = sorted((n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                    key=lambda s: int(re.search(r"(\d+)", s).group(1)))
    if len(slides) != N_TEXT + 1:
        return ["应有 %d 页, 实际 %d 页" % (N_TEXT + 1, len(slides))]
    probs, slots, parts = [], [], []
    for s in slides[1:]:
        no, part = _audio_slot(z, s)
        if no is None or not part:
            probs.append("%s 那一页没有可用的音频对象" % os.path.basename(s))
            continue
        if part not in names:
            probs.append("%s 的音频指向 %s, 包里却没有这个部件"
                         % (os.path.basename(s), part))
        slots.append(no); parts.append(part)
        xml = z.read(s).decode("utf-8")
        for tb in ("myTextTm0", "myTextBox"):
            if 'name="%s"' % tb not in xml:
                probs.append("%s 里找不到文本框 %s" % (os.path.basename(s), tb))
    if not probs and sorted(slots) != list(range(1, N_TEXT + 1)):
        probs.append("音频对象名没凑成 Text 01~10 各一个, 实际 %s" % slots)
    dup = sorted({p for p in parts if parts.count(p) > 1})
    if dup:                       # 两页共用一个部件 -> 有一段音频会被覆盖, 必须拦住
        probs.append("有两页指向同一个音频部件(会串页/丢音频): %s" % ", ".join(dup))
    if 'name="TextBox01"' not in z.read(slides[0]).decode("utf-8"):
        probs.append("第 1 页封面里找不到文本框 TextBox01")
    media = [os.path.basename(n) for n in names if re.match(r"ppt/media/media\d+\.mp3$", n)]
    if len(media) != N_TEXT:
        probs.append("ppt/media 里应有 %d 个 mp3, 实际 %d 个" % (N_TEXT, len(media)))
    return probs


def ensure_template():
    """模板不合格时尝试用自动备份还原; 还原不了就带着修复指引报错。"""
    probs = check_template(TEMPLATE)
    if not probs:
        return
    if not check_template(TEMPLATE_BAK):
        print("[..] 模板结构不对, 用自动备份还原")
        for p in probs:
            print("       -", p)
        shutil.copy2(TEMPLATE_BAK, TEMPLATE)
        return
    raise SystemExit(
        "[ERR] 模板 听力课件模板.pptx 结构不对(别用 PowerPoint 打开它编辑!):\n     - "
        + "\n     - ".join(probs)
        + "\n     修复: python make_ppt_courseware.py --rebuild-template <一个能正常播放的成品.pptx>"
          "\n           (不给成品也行, 程序会自动在本文件夹里找)")


# ----------------------------------------------------------------- OOXML 组装
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")   # XML 1.0 不允许的控制字符


def esc(t):
    return _ILLEGAL.sub("", t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def flush_lines(lines):
    """docx 里的软换行(w:br)会变成 '\\n'; 拆成独立段落, 别让 XML 里出现裸换行。"""
    out = []
    for x in lines:
        out += str(x).split("\n") if "\n" in str(x) else [x]
    return out


def rebuild_txbody(tb, lines):
    """保留原 bodyPr / lstStyle / rPr, 用 lines 重建段落。
    lstStyle 必须留着 —— 首页大标题的字号/字体/阴影全靠它, 丢了会变回默认样式。"""
    def grab(pat, dflt):
        m = re.search(pat, tb, re.S)
        return m.group(0) if m else dflt
    body = grab(r"<a:bodyPr[^>]*/>|<a:bodyPr[^>]*>.*?</a:bodyPr>", "<a:bodyPr/>")
    lst = grab(r"<a:lstStyle[^>]*/>|<a:lstStyle[^>]*>.*?</a:lstStyle>", "<a:lstStyle/>")
    rpr = grab(r"<a:rPr[^>]*/>|<a:rPr[^>]*>.*?</a:rPr>", "")
    ps = "".join(
        '<a:p><a:endParaRPr/></a:p>' if x == "" else
        '<a:p><a:r>%s<a:t xml:space="preserve">%s</a:t></a:r></a:p>' % (rpr, esc(x))
        for x in flush_lines(lines))
    return "<p:txBody>%s%s%s</p:txBody>" % (body, lst, ps)


def set_shape_text(xml, name, lines):
    for m in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
        blk = m.group(0)
        if not re.search(r'<p:cNvPr[^>]*name="%s"' % re.escape(name), blk):
            continue
        tb = re.search(r"<p:txBody>.*?</p:txBody>", blk, re.S).group(0)
        return xml[:m.start()] + blk.replace(tb, rebuild_txbody(tb, lines)) + xml[m.end():]
    raise SystemExit("[ERR] 模板里没找到文本框: %s" % name)


def rel_target(rels_bytes, rid):
    root = ET.fromstring(rels_bytes)
    for r in root.findall(R_ID + "Relationship"):
        if r.get("Id") == rid:
            return r.get("Target")
    return None


def process_slide(xml, rels, tm, yw, seg, media_out):
    """替换本页音频 + 两个文本框; 返回 (新xml, Text段号)"""
    for m in re.finditer(r"<p:pic[ >].*?</p:pic>", xml, re.S):
        blk = m.group(0)
        if "<a:audioFile" not in blk:
            continue
        name = re.search(r'<p:cNvPr[^>]*name="([^"]*)"', blk)
        no = int(re.search(r"(\d+)", name.group(1)).group(1)) if name else None
        rid = re.search(r'<a:audioFile[^>]*r:link="([^"]*)"', blk)
        tgt = rel_target(rels, rid.group(1)) if rid else None
        if no is None or not tgt:
            raise SystemExit("[ERR] 模板某页的音频对象不完整(名字或关系缺失), "
                             "用 --rebuild-template 重建模板")
        part = posixpath.normpath(posixpath.join("ppt/slides", tgt)).lstrip("/")
        if seg is not None:                    # seg=None -> 保留模板原音频(造模板用)
            if part in media_out:              # 体检已拦, 这里是最后一道保险
                raise SystemExit("[ERR] 模板里两页共用了同一个音频部件 %s, 会丢一段音频" % part)
            media_out[part] = seg[no]
        break
    else:
        raise SystemExit("[ERR] 某页没有音频对象")
    xml = set_shape_text(xml, "myTextTm0", tm[no - 1] + [""])
    xml = set_shape_text(xml, "myTextBox", yw[no - 1] + [""])
    return xml, no


def build_pptx(src, out, title, tm_parts, yw_parts, segs):
    """src=模板/源pptx, segs={段号:路径} 或 None(保留原音频)"""
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    slides = sorted((n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                    key=lambda s: int(re.search(r"(\d+)", s).group(1)))
    if len(slides) != N_TEXT + 1:
        raise SystemExit("[ERR] 模板应 %d 页, 实际 %d 页" % (N_TEXT + 1, len(slides)))

    new_slides, media_out, seen = {}, {}, []
    for s in slides[1:]:                       # 第 1 页是封面, 其余按 Text 段
        n = int(re.search(r"(\d+)", s).group(1))
        rx = "ppt/slides/_rels/slide%d.xml.rels" % n
        xml, no = process_slide(zin.read(s).decode("utf-8"), zin.read(rx),
                                tm_parts, yw_parts, segs, media_out)
        new_slides[s] = xml.encode("utf-8")
        seen.append(no)
    if sorted(seen) != list(range(1, N_TEXT + 1)):
        raise SystemExit("[ERR] 模板音频段号不成套: %s" % sorted(seen))

    xml = zin.read(slides[0]).decode("utf-8")
    xml = set_shape_text(xml, "TextBox01", [title])
    new_slides[slides[0]] = xml.encode("utf-8")

    core = zin.read("docProps/core.xml").decode("utf-8")
    core = re.sub(r"<dc:title>.*?</dc:title>",
                  "<dc:title>%s</dc:title>" % esc(title), core)

    tmp = out + ".tmp"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                name = info.filename
                if name in new_slides:
                    data = new_slides[name]
                elif name in media_out:
                    data = media_out[name]
                elif name == "docProps/core.xml":
                    data = core.encode("utf-8")
                else:
                    data = None
                if data is None:
                    zout.writestr(info, zin.read(name))
                else:
                    zout.writestr(info, data)
    except BaseException:
        try:
            os.remove(tmp)                     # 半成品别留在文件夹里
        except OSError:
            pass
        raise
    finally:
        zin.close()
    try:
        os.replace(tmp, out)
    except OSError as e:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise SystemExit("[ERR] 覆盖不了 %s (%s)\n"
                         "     多半是这份成品正被 PowerPoint 打开, 关掉它再跑一次"
                         % (os.path.basename(out), e))
    return sorted(media_out)


# ----------------------------------------------------------------- main
def find_good_pptx(folder):
    """在 folder 里找一个结构合格、能当模板源的成品 pptx(最新优先); 没有返回 None。"""
    cands = []
    for f in os.listdir(folder):
        if not f.lower().endswith(".pptx") or f.startswith("~$"):
            continue
        p = os.path.join(folder, f)
        try:
            cands.append((os.path.getmtime(p), p))
        except OSError:
            pass
    for _, p in sorted(cands, reverse=True):
        if not check_template(p):
            print("[..] 可作模板源的成品: %s" % os.path.basename(p))
            return p
    return None


def rebuild_template(src):
    """用一份结构完好的成品 pptx 重建模板(清空三个文本框, 保留音频槽)。"""
    build_pptx(src, TEMPLATE, '', [['']] * N_TEXT, [['']] * N_TEXT, None)
    probs = check_template(TEMPLATE)
    if probs:
        raise SystemExit("[ERR] 重建后仍不合格: %s" % probs)
    shutil.copy2(TEMPLATE, TEMPLATE_BAK)
    print("[OK] 模板已重建: %s" % TEMPLATE)


def main():
    argv   = sys.argv[1:]
    if any(a in ("-h", "--help", "-?", "/?") for a in argv):
        print((__doc__ or "").strip())
        print("\n另: --rebuild-template [成品.pptx] 不给成品时自动在本文件夹里找一个结构合格的")
        return
    args = [a for a in argv if not a.startswith("--")]
    recut  = "--recut"  in argv          # 强制重新切分(跳过 audio\/音频文件\, 忽略 _seg.json 缓存)
    legacy = "--legacy" in argv          # 额外产出 tm.txt/yw.txt/音频文件\ (喂老 .pptm 宏路); 默认不要

    if "--rebuild-template" in argv:
        src = os.path.abspath(args[0]) if args else find_good_pptx(SCRIPT_DIR)
        if not src:
            raise SystemExit("[ERR] 用法: python make_ppt_courseware.py --rebuild-template "
                             "<一个能正常播放的成品.pptx>\n"
                             "     不给参数时程序会自动在本文件夹里找, 但这次没找到合格的 pptx")
        if not os.path.isfile(src):
            raise SystemExit("[ERR] 找不到源 pptx: %s" % src)
        rebuild_template(src)
        return

    folder = os.path.abspath(args[0]) if args else SCRIPT_DIR
    if not os.path.isdir(folder):
        raise SystemExit("[ERR] 不是文件夹: %s" % folder)
    ensure_template()

    P = pick_prefix(folder)
    print("[OK] 课件基名 P = %s" % P)

    paper = _pick_docx(folder, ("试卷", "试题卷", "试题"), exclude=("答案", "解析", "原文"))
    trans = _pick_docx(folder, ("听力原文",), exclude=("答案", "解析"))
    if not paper:
        raise SystemExit("[ERR] 没找到试卷 docx(文件名里要有「试卷/试题卷/试题」)")
    if not trans:
        raise SystemExit("[ERR] 没找到听力原文 docx(文件名里要有「听力原文」)")
    print("[OK] 试卷   : %s" % os.path.basename(paper[0]))
    print("[OK] 听力原文: %s" % os.path.basename(trans[0]))

    segs = ensure_segments(folder, P, recut)
    pparas = _paras(paper[0])
    QA = scan_questions(pparas)
    groups = question_groups(pparas)
    tx = parse_transcripts(trans[0])
    print("[OK] 题目 %d 道 / 原文 %d 段 / 音频 %d 个" % (len(QA), len(tx), len(segs)))

    tm_parts, yw_parts = build_lines(groups, QA, tx)
    title = P + "听力"
    out = os.path.join(folder, title + "课件.pptx")
    seg = {i: open(segs[i - 1], "rb").read() for i in range(1, N_TEXT + 1)}
    done = build_pptx(TEMPLATE, out, title, tm_parts, yw_parts, seg)
    print("[OK] 已嵌入音频 %d 个 -> %s" % (len(done), ", ".join(done[:3]) + " ..."))
    shutil.copy2(TEMPLATE, TEMPLATE_BAK)     # 这次模板好用, 存一份自动备份以防万一

    if legacy:
        # 老 .pptm 宏路的素材(tm.txt/yw.txt/音频文件\); 默认不生成, 需要时加 --legacy
        tm = "*".join("\n".join(ls) + "\n" for ls in tm_parts).replace("\n", "\r\n")
        yw = "*".join("\n".join(ls) + "\n" for ls in yw_parts).replace("\n", "\r\n")
        io.open(os.path.join(folder, "tm.txt"), "wb").write(tm.encode("gbk"))
        io.open(os.path.join(folder, "yw.txt"), "wb").write(yw.encode("gbk"))
        aud = os.path.join(folder, "音频文件")
        os.makedirs(aud, exist_ok=True)
        for i in range(1, N_TEXT + 1):
            src = segs[i - 1]
            dst = os.path.join(aud, "Text %02d.mp3" % i)
            if os.path.exists(dst) and os.path.samefile(src, dst):
                continue                   # 音频本来就取自 音频文件\, 不必自拷
            shutil.copy2(src, dst)
        print("[OK] 已刷新 tm.txt / yw.txt / 音频文件\\Text 01-10.mp3 (--legacy)")
    print("[DONE] 成品: %s  (%d bytes)" % (out, os.path.getsize(out)))


if __name__ == "__main__":
    main()
