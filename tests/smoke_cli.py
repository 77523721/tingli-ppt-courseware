# -*- coding: utf-8 -*-
"""命令行冒烟: 造一个假卷文件夹, 真的把 make_ppt_courseware.py 跑一遍。

单元测试验的是函数, 这个验的是**整条路**: 找素材 -> 认试卷 -> 认原文 -> 取 audio\\ ->
替换模板音频 -> 写文本框 -> 存备份。跑法:
    python tests/smoke_cli.py
不需要 numpy, 也不需要 ffmpeg(用现成的 audio\\ 分段)。
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fixtures as F                                          # noqa: E402

# 和主程序一样: 自己把输出设成 UTF-8
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def run(cwd, *args):
    return subprocess.run([sys.executable, "make_ppt_courseware.py"] + list(args),
                          cwd=cwd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def main():
    work = tempfile.mkdtemp(prefix="tingli-cli-")
    try:
        shutil.copy2(os.path.join(F.ROOT, "make_ppt_courseware.py"), work)
        F.write_template(os.path.join(work, "听力课件模板.pptx"))
        F.write_paper(os.path.join(work, "2027届周测英语试卷.docx"), style="paragraph")
        F.write_transcript(os.path.join(work, "2027届周测英语听力原文.docx"))
        F.write_audio_segments(os.path.join(work, "audio"))

        # ---- --help 应当能跑且退出码为 0
        r = run(work, "--help")
        assert r.returncode == 0, "--help 退出码 %d\n%s" % (r.returncode, r.stderr)
        assert "--recut" in r.stdout and "--rebuild-template" in r.stdout, r.stdout

        # ---- 正常出片
        r = run(work)
        sys.stdout.write(r.stdout)
        sys.stderr.write(r.stderr)
        assert r.returncode == 0, "出片失败, 退出码 %d" % r.returncode

        out = os.path.join(work, "2027届周测英语试卷听力课件.pptx")
        assert os.path.isfile(out), "没看到成品: %s" % os.listdir(work)
        with zipfile.ZipFile(out) as z:                  # 用完必须关, 否则 Windows 锁住文件
            slides = [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)]
            media = [n for n in z.namelist() if re.match(r"ppt/media/media\d+\.mp3$", n)]
            s2 = z.read("ppt/slides/slide2.xml").decode("utf-8")
        assert len(slides) == 11, "应当 11 页, 实际 %d" % len(slides)
        assert len(media) == 10, "应当 10 个音频部件, 实际 %d" % len(media)
        assert "1. What is the answer to item 1?" in s2, "第 1 题没落在第 2 页上"
        assert os.path.isfile(os.path.join(work, "听力课件模板-备份.pptx")), "没留备份"

        # ---- 模板被改坏 -> 应当自动从备份还原后照常出片
        bad = os.path.join(work, "听力课件模板.pptx")
        good = os.path.join(work, "听力课件模板-备份.pptx")
        F.write_template(bad, drop_media=5)               # 覆盖成残缺模板
        r = run(work)
        assert r.returncode == 0, "自愈失败, 退出码 %d\n%s" % (r.returncode, r.stderr)
        assert "还原" in r.stdout, "没走自愈分支:\n%s" % r.stdout
        with zipfile.ZipFile(bad) as z:
            restored = [n for n in z.namelist() if re.match(r"ppt/media/media\d+\.mp3$", n)]
        assert len(restored) == 10, "模板没被还原(%d 个 mp3)" % len(restored)
        assert os.path.getsize(good) > 0

        # ---- 模板和备份都坏 -> 必须报错退出, 不能假装成功
        F.write_template(bad, duplicate=(3, 1))
        F.write_template(good, duplicate=(3, 1))
        r = run(work)
        assert r.returncode != 0, "坏模板竟然跑成功了"
        assert "同一个音频部件" in r.stdout + r.stderr, r.stdout + r.stderr

        print("  [OK]   CLI 冒烟通过")
        return 0
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
