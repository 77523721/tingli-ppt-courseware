# -*- coding: utf-8 -*-
"""make_ppt_courseware 的自检套件。

跑法(不需要 pytest, 也不需要 numpy):
    python tests/test_courseware.py

它验的是**过去真出过错**的那几件事:
  - 试卷排版: 选项同段/各占一段/全角点/逗号误打/漏点号 -> 都必须认出 1~20 题
  - 提示语写"录音"还是"材料" -> 题号到 Text 的分组都必须正确
  - 题干漏写题号 -> 按位置补上并出声
  - 模板体检: 缺部件 / 两页共用一个部件 / 少文本框 -> 必须报出来, 不许静默放过
  - 端到端: 合成试卷+模板+音频 -> 产出 11 页 pptx, 音频原样嵌入, 文字落在对的页上
"""
import atexit
import contextlib
import importlib.util
import io
import os
import re
import shutil
import sys
import tempfile
import traceback
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fixtures as F                                          # noqa: E402

# 和主程序一样: 自己把输出设成 UTF-8, 管道/重定向到文件时中文不乱码
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

TMP = tempfile.mkdtemp(prefix="tingli-selftest-")
atexit.register(lambda: shutil.rmtree(TMP, ignore_errors=True))


def _load():
    # COURSEWARE_MODULE 可以指向别的版本, 用来确认"这套自检确实抓得住老 bug"
    path = os.environ.get("COURSEWARE_MODULE") or os.path.join(F.ROOT, "make_ppt_courseware.py")
    spec = importlib.util.spec_from_file_location("make_ppt_courseware", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


m = _load()


def _d(name):
    p = os.path.join(TMP, name)
    os.makedirs(p, exist_ok=True)
    return p


# ---------------------------------------------------------------- 选项切分
def test_split_options_inline():
    got = m.split_options("A. A business client.\tB. A job applicant.\tC. A new staff member.")
    assert got == ["A. A business client.", "B. A job applicant.", "C. A new staff member."], got


def test_split_options_styles():
    """六种真实排版, 拆出来必须一模一样。"""
    assert m.split_options("A. Alpha.\tB. Beta.\tC. Gamma.") == ["A. Alpha.", "B. Beta.", "C. Gamma."]
    assert m.split_options("A. Alpha.\nB. Beta.\nC. Gamma.") == ["A. Alpha.", "B. Beta.", "C. Gamma."]
    assert m.split_options("A．Alpha. B．Beta. C．Gamma.") == ["A. Alpha.", "B. Beta.", "C. Gamma."]
    assert m.split_options("A.Alpha.\tB.Beta.\tC.Gamma.") == ["A. Alpha.", "B. Beta.", "C. Gamma."]
    got = m.split_options("A. Alpha.\t\tB Beta.\tC. Gamma.", loose=True)
    assert got == ["A. Alpha.", "B. Beta.", "C. Gamma."], got
    got = m.split_options("A, Alpha.\tB. Beta.\tC. Gamma.", loose=True)
    assert got == ["A. Alpha.", "B. Beta.", "C. Gamma."], got


def test_split_options_edges():
    assert m.split_options("A. Alpha.") == ["A. Alpha."]
    assert m.split_options("What is the man doing?") == []
    assert m.split_options("   ") == []
    # 严格式对漏点号只认得出 2 项 —— 这一步是"该走兜底"的信号
    assert len(m.split_options("A. Alpha.\tB Beta.\tC. Gamma.")) == 2


# ---------------------------------------------------------------- 试卷解析
def test_scan_questions_all_styles():
    """每种排版都要凑齐 1~20 题, 且选项文字完全相同。"""
    for style in F.OPT_STYLES:
        paras = F.paper_lines(style=style)
        QA = m.scan_questions(paras)
        assert sorted(QA) == list(range(1, 21)), (style, sorted(QA))
        for q in range(1, 21):
            assert QA[q][0] == "What is the answer to item %d?" % q, (style, q, QA[q])
            assert QA[q][1] == F.EXPECTED_OPTS, (style, q, QA[q][1])


def test_scan_questions_skips_non_listening():
    """注意事项里的 "1．答卷前…" 不是听力题, 不能占掉第 1 题。"""
    QA = m.scan_questions(F.paper_lines())
    assert QA[1][0] == "What is the answer to item 1?"


def test_scan_questions_repairs_missing_number():
    """题干漏写题号(2025 高考Ⅰ卷真实情况) -> 按位置补上并打印提示。"""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        QA = m.scan_questions(F.paper_lines(style="paragraph", unnumbered=(14,)))
    assert sorted(QA) == list(range(1, 21))
    assert QA[14][0] == "What is the answer to item 14?"
    assert QA[14][1] == F.EXPECTED_OPTS
    assert "第14题" in buf.getvalue(), buf.getvalue()


def test_scan_questions_error_names_missing_numbers():
    """连着的两题都没题号 -> 不敢猜, 必须报错并点名是第几题。"""
    try:
        m.scan_questions(F.paper_lines(unnumbered=(6, 7)))
    except SystemExit as e:
        msg = str(e)
        assert "缺第" in msg and "6" in msg and "7" in msg, msg
    else:
        raise AssertionError("应当报错")


def test_question_groups_wording_and_layouts():
    for word in ("录音", "材料"):
        for groups in (F.GROUPS_GAOKAO, F.GROUPS_ZHOUCE):
            for style in ("inline_dot", "paragraph"):
                g = m.question_groups(F.paper_lines(style=style, groups=groups, word=word))
                want = {t: [t] for t in range(1, 6)}
                want.update(groups)
                assert g == want, (word, style, g)


def test_question_groups_error_names_ungrouped():
    paras = [x for x in F.paper_lines() if not x.startswith("听第")]
    try:
        m.question_groups(paras)
    except SystemExit as e:
        assert "没归到任何一段" in str(e), str(e)
    else:
        raise AssertionError("应当报错")


def test_parse_transcripts():
    p = F.write_transcript(os.path.join(_d("tr"), "transcript.docx"))
    tx = m.parse_transcripts(p)
    assert sorted(tx) == list(range(1, 11))
    assert tx[1] == ["M: This is text 1, line one.", "W: And this is text 1, line two."], tx[1]
    assert all("— 2 —" not in ln for ls in tx.values() for ln in ls), "页码残渣没被丢掉"


# ---------------------------------------------------------------- 命名规则
def test_pick_prefix():
    d = _d("pfx1")
    open(os.path.join(d, "某卷_英语听力.mp3"), "wb").close()
    assert m.pick_prefix(d) == "某卷"
    d = _d("pfx2")
    open(os.path.join(d, "黄石市2026-2027学年上学期高三9月质量监测英语听力.mp3"), "wb").close()
    assert m.pick_prefix(d) == "黄石市2026-2027学年上学期高三9月质量监测"
    d = _d("pfx3")
    open(os.path.join(d, "2024年普通高等学校招生全国统一考试英语试题听力音频.mp3"), "wb").close()
    assert m.pick_prefix(d) == "2024年普通高等学校招生全国统一考试英语试题"


# ---------------------------------------------------------------- 文本安全
def test_esc_and_flush_lines():
    assert m.esc("a\x0bb\x0cc<&>") == "abc&lt;&amp;&gt;"
    assert m.esc('Tom & "Jerry" <b>') == 'Tom &amp; "Jerry" &lt;b&gt;'
    assert m.flush_lines(["一行\n二行", "ok"]) == ["一行", "二行", "ok"]


# ---------------------------------------------------------------- 模板体检
def test_check_template_accepts_good():
    p = F.write_template(os.path.join(_d("tpl"), "good.pptx"))
    assert m.check_template(p) == []


def test_check_template_missing_file():
    assert m.check_template(os.path.join(TMP, "根本没有.pptx")) != []


def test_check_template_catches_missing_media():
    p = F.write_template(os.path.join(_d("tpl"), "nomedia.pptx"), drop_media=5)
    probs = m.check_template(p)
    assert probs, "缺部件没被发现"
    assert any("media5" in x for x in probs), probs


def test_check_template_catches_shared_media():
    """两页共用一个音频部件 —— 这正是旧版会静默丢一段音频的坏法。"""
    p = F.write_template(os.path.join(_d("tpl"), "dup.pptx"), duplicate=(3, 1))
    probs = m.check_template(p)
    assert probs, "共用部件没被发现"
    assert any("同一个音频部件" in x for x in probs), probs


def test_check_template_catches_missing_textbox():
    p = F.write_template(os.path.join(_d("tpl"), "notb.pptx"), drop_textbox=4)
    probs = m.check_template(p)
    assert probs and any("myTextBox" in x for x in probs), probs


# ---------------------------------------------------------------- 端到端
def test_build_pptx_end_to_end():
    d = _d("e2e")
    tpl = F.write_template(os.path.join(d, "template.pptx"))
    audio = F.write_audio_segments(os.path.join(d, "audio"))
    segs = {i: open(audio[i - 1], "rb").read() for i in range(1, 11)}

    import docx
    paper = F.write_paper(os.path.join(d, "paper.docx"), style="paragraph")
    trans = F.write_transcript(os.path.join(d, "transcript.docx"))
    pparas = [p.text for p in docx.Document(paper).paragraphs]
    QA = m.scan_questions(pparas)
    groups = m.question_groups(pparas)
    tx = m.parse_transcripts(trans)
    tm_parts, yw_parts = m.build_lines(groups, QA, tx)

    out = os.path.join(d, "out.pptx")
    m.build_pptx(tpl, out, "测试卷听力", tm_parts, yw_parts, segs)

    z = zipfile.ZipFile(out)
    slides = [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)]
    assert len(slides) == 11, "应当 11 页, 实际 %d" % len(slides)

    for i in range(1, 11):                       # 音频必须原样嵌入
        assert z.read("ppt/media/media%d.mp3" % i) == segs[i], "media%d 字节不对" % i

    s2 = z.read("ppt/slides/slide2.xml").decode("utf-8")
    assert "1. What is the answer to item 1?" in s2
    assert "A. Alpha" in s2 and "B. Beta" in s2 and "C. Gamma" in s2
    assert "text 1, line one" in s2

    s7 = z.read("ppt/slides/slide7.xml").decode("utf-8")     # Text 06 承载第 6、7 题
    assert "6. What is the answer to item 6?" in s7
    assert "7. What is the answer to item 7?" in s7
    assert "text 6, line one" in s7

    assert "测试卷听力" in z.read("ppt/slides/slide1.xml").decode("utf-8")
    assert "测试卷听力" in z.read("docProps/core.xml").decode("utf-8")
    assert m.check_template(out) == [], "产物自己应当是合格模板"


# ---------------------------------------------------------------- 零依赖跑法
def main():
    tests = sorted((n, f) for n, f in globals().items()
                   if n.startswith("test_") and callable(f))
    bad = []
    for name, fn in tests:
        try:
            fn()
            print("  [OK]   %s" % name)
        except KeyboardInterrupt:
            raise
        except BaseException as e:               # SystemExit 也要抓, 免得整轮被打断
            bad.append(name)
            print("  [FAIL] %s: %s" % (name, e))
            traceback.print_exc()
    print("\n%d/%d 通过" % (len(tests) - len(bad), len(tests)))
    if bad:
        print("失败: " + ", ".join(bad))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
