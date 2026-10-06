# tingli-ppt-courseware

听力 PPT 课件生成器 —— 把「听力 mp3 + 试卷 docx + 听力原文 docx」拼成可直接上课的
`<卷名>听力课件.pptx`。纯 OOXML 手术（替换 `ppt/media/*.mp3` 字节 + 重写文本框），
**不依赖 PowerPoint / COM / 宏**。

## 文件

| 文件 | 说明 |
|---|---|
| `听力PPT课件使用说明.md` | **完整说明书，先看这个** |
| `make_ppt_courseware.py` | 主程序（纯 Python） |
| `tests/` | 自检套件（合成素材，仓库里不含任何真实试卷或音频） |
| `.github/workflows/ci.yml` | 每次 push 自动跑上面这套自检 |

## 怎么跑

```
python make_ppt_courseware.py
```

不带参数就处理脚本所在的文件夹。中文输出由脚本自己设成 UTF-8，
**不需要 `.bat` 包装**（管道/重定向也不会乱码）。

## 本地自检

```
python tests/test_courseware.py      # 解析 + 模板体检 + 端到端打包
python tests/smoke_cli.py            # 命令行冒烟(真的跑一遍出片)
```

只需要 `python-docx`。**不需要** pytest、numpy 或 ffmpeg。

想确认这套自检真的抓得住老 bug，可以拿旧版跑一遍：

```
set COURSEWARE_MODULE=旧版路径\make_ppt_courseware.py
python tests/test_courseware.py
```

## 仓库里为什么不放试卷和音频

版权与体积。`.gitignore` 排除全部 `*.mp3` / `*.docx` / `*.pptx`、`audio\`、`音频文件\`，
CI 用**运行时现造**的合成素材（见 `tests/fixtures.py`），所以任何机器上都能自证。
