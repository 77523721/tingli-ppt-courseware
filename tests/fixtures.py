# -*- coding: utf-8 -*-
"""CI 自检用的**合成素材** —— 运行时现造, 仓库里不放任何二进制文件。

为什么这么干:
  - 真实试卷/音频有版权也有体积, 不该进仓库;
  - 造素材的过程本身就是"试卷排版有哪几种写法"的活文档;
  - 任何机器上都能自证, 不依赖 L:\\ 下面的文件。
"""
import os
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RP = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

# 两组"题号 -> Text 段"的分法, 真实卷子里最常见
GROUPS_ZHOUCE = {6: [6, 7], 7: [8, 9, 10], 8: [11, 12, 13], 9: [14, 15, 16], 10: [17, 18, 19, 20]}
GROUPS_GAOKAO = {6: [6, 7], 7: [8, 9, 10], 8: [11, 12, 13], 9: [14, 15, 16, 17], 10: [18, 19, 20]}

# 选项的几种真实排版: 同段 / 各占一段 / 全角点 / 逗号误打 / 漏点号
# 五种写法拆出来必须是同一结果, 所以选项正文统一用 "Alpha." 这种带句点的写法
OPT_STYLES = {
    "inline_dot":  lambda: ["A. Alpha.\tB. Beta.\tC. Gamma."],
    "paragraph":   lambda: ["A. Alpha.", "B. Beta.", "C. Gamma."],
    "fullwidth":   lambda: ["A．Alpha. B．Beta. C．Gamma."],
    "comma_typo":  lambda: ["A, Alpha.\tB. Beta.\tC. Gamma."],
    "dotless":     lambda: ["A. Alpha.\tB Beta.\tC. Gamma."],
}

EXPECTED_OPTS = ["A. Alpha.", "B. Beta.", "C. Gamma."]


def stem(q, numbered=True):
    return ("%d. What is the answer to item %d?" % (q, q)) if numbered \
        else ("What is the answer to item %d?" % q)


# ---------------------------------------------------------------- 试卷 docx
def paper_lines(style="inline_dot", groups=None, word="录音", unnumbered=()):
    """造一份 1~20 题的听力试卷文本(段落列表)。

    style:      OPT_STYLES 里的键
    word:       提示语写"录音"还是"材料"
    unnumbered: 这些题的题干**漏写题号**(真实卷子里 2025 高考Ⅰ卷第 14 题就是这样)
    """
    # 第 1~5 段各 1 题, 第 6~10 段按 groups 走
    all_groups = {t: [t] for t in range(1, 6)}
    all_groups.update(groups or GROUPS_GAOKAO)
    opts = OPT_STYLES[style]
    ls = ["2027届黄石一中高三年级英语周测试卷",
          # 注意事项里的"1."不是听力题, 解析器必须跳过它
          "1．答卷前，先将自己的姓名、准考证号填写在试卷上。",
          "2．选择题的作答：选出答案后，用2B铅笔把答题卡上对应题目的答案标号涂黑。",
          "第一部分 听力(共两节,满分30分)",
          "第一节(共5小题;每小题1.5分,满分7.5分)",
          "听下面5段录音。每段录音后有一个小题,从题中所给的A、B、C三个选项中选出最佳选项。"]
    for t in sorted(all_groups):
        if t > 5:
            nums = all_groups[t]
            if len(nums) == 2:
                tail = "%d、%d" % (nums[0], nums[1])
            elif nums[-1] - nums[0] == len(nums) - 1:
                tail = "%d至%d" % (nums[0], nums[-1])
            else:
                tail = "、".join(str(x) for x in nums)
            ls.append("听第%d段%s，回答第%s题。" % (t, word, tail))
        for q in all_groups[t]:
            ls.append(stem(q, numbered=(q not in unnumbered)))
            ls += opts()
    ls += ["第二部分 阅读(共两节,满分50分)", "21. A reading question?", "A. One. B. Two. C. Three. D. Four."]
    return ls


def write_paper(path, **kw):
    import docx
    d = docx.Document()
    for line in paper_lines(**kw):
        d.add_paragraph(line)
    d.save(path)
    return path


# ---------------------------------------------------------------- 原文 docx
def transcript_lines():
    ls = ["听力原文"]
    for t in range(1, 11):
        ls.append("Text %d" % t)
        ls.append("M: This is text %d, line one." % t)
        ls.append("W: And this is text %d, line two." % t)
        if t == 4:
            ls.append("— 2 —")                       # 页码残渣, 必须被丢掉
        if t == 5:
            ls.append("")
    return ls


def write_transcript(path):
    import docx
    d = docx.Document()
    for line in transcript_lines():
        d.add_paragraph(line)
    d.save(path)
    return path


def write_audio_segments(audio_dir, n=10):
    """造 10 个假的分段音频, 内容可辨认, 便于比对是否被原样嵌进 pptx。"""
    os.makedirs(audio_dir, exist_ok=True)
    out = []
    for i in range(1, n + 1):
        p = os.path.join(audio_dir, "seg%02d.mp3" % i)
        with open(p, "wb") as f:
            f.write(b"FAKE-MP3-%02d" % i + bytes([i]) * 64)
        out.append(p)
    return out


# ---------------------------------------------------------------- 模板 pptx
def _txbox(name):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="9" name="%s"/><p:cNvSpPr txBox="1"/>'
            '<p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:endParaRPr/></a:p></p:txBody></p:sp>' % name)


def _slide_xml(k, with_textbox=True):
    """第 k 段(1~10)的页面: 一个音频对象 + 两个文本框。"""
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<p:sld xmlns:p="%s" xmlns:a="%s" xmlns:r="%s"><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Text %02d"/><p:cNvPicPr/>'
            '<p:nvPr><a:audioFile r:link="rId2"/></p:nvPr></p:nvPicPr>'
            '<p:blipFill/><p:spPr/></p:pic>'
            '%s%s'
            '</p:spTree></p:cSld></p:sld>'
            % (P, A, R, k, _txbox("myTextTm0"), _txbox("myTextBox") if with_textbox else ""))


def _cover_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<p:sld xmlns:p="%s" xmlns:a="%s" xmlns:r="%s"><p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>%s</p:spTree></p:cSld></p:sld>' % (P, A, R, _txbox("TextBox01")))


def _rels_xml(media_no):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="%s">'
            '<Relationship Id="rId1" Type="%s/media" Target="../media/media%d.mp3"/>'
            '<Relationship Id="rId2" Type="%s/audio" Target="../media/media%d.mp3"/>'
            '</Relationships>' % (RP, R, media_no, R, media_no))


CORE = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>旧标题</dc:title></cp:coreProperties>')

CONTENT_TYPES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Types xmlns="%s">'
                 '<Default Extension="mp3" ContentType="audio/mpeg"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Default Extension="rels" '
                 'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '</Types>' % CT)


def write_template(path, duplicate=None, drop_media=None, drop_textbox=None):
    """造一个**结构上够用**的最小模板(不是能在 PowerPoint 里放映的成品)。

    duplicate=(第几页, 媒体号)  让某页指向别的音频部件(两页共用一个)
    drop_media=K                抽掉 mediaK.mp3
    drop_textbox=第几页          抽掉该页的 myTextBox
    """
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("docProps/core.xml", CORE)
        z.writestr("ppt/slides/slide1.xml", _cover_xml())
        z.writestr("ppt/slides/_rels/slide1.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   '<Relationships xmlns="%s"/>' % RP)
        for k in range(1, 11):
            slide_no = k + 1
            media_no = k
            if duplicate and duplicate[0] == slide_no:
                media_no = duplicate[1]
            z.writestr("ppt/slides/slide%d.xml" % slide_no,
                       _slide_xml(k, with_textbox=(drop_textbox != slide_no)))
            z.writestr("ppt/slides/_rels/slide%d.xml.rels" % slide_no,
                       _rels_xml(media_no))
            if drop_media == k:
                continue
            z.writestr("ppt/media/media%d.mp3" % k, b"TEMPLATE-PLACEHOLDER-%02d" % k)
    return path
