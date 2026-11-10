# 圆环连招 · ring-combo

![效果](preview.gif)

7 页连在一起的封面 + 目录 + 章节，页与页之间用「平滑」切换。四个带十字缝的圆环里透出底图，一路变形：

1. 画面中间四个同心圆环，透出放大的底图。
2. 圆环一边转一边放大到画面外，底图拉回全景，封面标题和署名出现。
3. 四个圆环缩小排成一行，变成目录，每个环里一个编号。
4. 到第 7 页：1 号圆环放大停在左边，每翻一页顺时针转 90°，编号跟着变，右边章节文字上下滚动切换。

适合职业规划大赛、毕业答辩、述职汇报、读书分享这种「封面 → 目录 → 分 4 部分讲」的 PPT。

## 让 AI 做

直接对 AI 说，例如：

> 用圆环连招做一份职业规划大赛的 PPT，封面写「大学生职业 / 生涯规划」，四个章节是认识自我、探索职业、规划路径、行动落地，用这张图。

没说的内容 AI 会帮你补上，做完把 .pptx 和一个网页预览交给你。用 PowerPoint 或 WPS 打开，按 F5 放映，点一下翻一页。

## 自己运行

```bash
pip install -r ../../requirements.txt
python generate.py --title "大学生职业" --title2 "生涯规划" --reporter "汇报人：XXX" \
  --section "Know Yourself|认识自我|Strengths. Interests. Values.\nStart with who you are." \
  --section "Explore Careers|探索职业|Find where you fit." \
  --section "Plan the Path|规划路径|One clear goal. Three steps." \
  --section "Take Action|行动落地|Small steps, every day." \
  --image 我的图.jpg --out 职业规划.pptx --preview 预览.html
```

## 能换什么

下面这些都可以在对话里直接跟 AI 说，AI 会换成对应的参数。

| 参数 | 作用 | 默认值 |
|---|---|---|
| `--title` | 封面标题第一行，5 个字左右最合适。传 `""` 不显示 | `大学生职业` |
| `--title2` | 封面标题第二行，往右错开、字更大，4 个字左右最合适。传 `""` 不显示 | `生涯规划` |
| `--reporter` | 封面署名条，例如「汇报人：XXX」。传 `""` 不显示 | `汇报人：XXX` |
| `--section` | 一个章节，**正好传 4 次**。格式 `英文标题\|中文小字\|正文`，正文用 `\n` 换行，最多 2 行；中文小字和正文可以留空 | 认识自我 / 探索职业 / 规划路径 / 行动落地 |
| `--contents` | 目录页大字。传 `""` 不显示 | `目 录` |
| `--contents-en` | 目录页大字后面的渐隐英文。传 `""` 不显示 | `CONTENTS` |
| `--part-label` | 章节页编号下面的小字。传 `""` 不显示 | `PART` |
| `--image` | 底图，横图最佳（16:9），建议 2K 以上。整套 7 页共用这一张 | `example.jpg` |
| `--title-font` | 封面标题字体文件 | `fonts/MaShanZheng-Regular.ttf` |
| `--en-font` / `--en-bold-font` | 英文正文 / 英文标题和编号字体 | `fonts/Montserrat-Regular.ttf` / `fonts/Montserrat-Bold.ttf` |
| `--cn-font` / `--cn-bold-font` | 中文小字 / 中文粗字（目录大字、署名）字体 | `fonts/NotoSansSC-Medium.ttf` / `fonts/NotoSansSC-Black.ttf` |
| `--duration` | 平滑切换时长，单位毫秒 | `1600` |
| `--out` | 输出的 PPT 文件名 | `圆环连招.pptx` |
| `--preview` | 同时输出一个网页预览 | 不输出 |

## 搭配建议

| 场景 | 封面标题 | 四个章节 |
|---|---|---|
| 职业规划大赛 | 大学生职业 / 生涯规划 | Know Yourself · Explore Careers · Plan the Path · Take Action |
| 毕业答辩 | 毕业设计 / 答辩汇报 | Background · Method · Results · Conclusion |
| 述职汇报 | 年度述职 / 报告 | Review · Results · Problems · Next Steps |

章节标题用英文短词、正文一两句话最有高级感，不要堆小字。底图选画面中间有主体（人、路、建筑）的横图，开场小圆环里能看到主体，效果最好。

## 说明

- 所有文字都已转成形状，换一台没装这些字体的电脑打开，效果也一样；但打开后文字不能再直接改，要改就重新运行一次。
- 圆环里透出的是幻灯片背景，所以底图要通过 `--image` 换，不要在 PPT 里直接改背景以外的图片。
- 平滑切换需要 PowerPoint 2019 及以上、Microsoft 365 或 WPS。旧版 PowerPoint 只能看到淡入淡出。
- 网页预览只用来看效果，正式放映以 PPT 为准。
