# 数字过渡页 · number-transition

![效果](preview.gif)

一组章节过渡页，每节一页：

- 整页铺一张图，左边一个超大的斜体数字，完整地放在画面里；右边是一块斜边的半透明色块，上面是英文标题和一句小字。
- 进入这一页时，底图从放大状态缓缓缩回原大；同时数字、色块和文字作为一个整体从右边飞入，越飞越慢（强调动画 + 飞入动画，不用点）。
- 色块颜色默认从每张图里自动取，图和色块天然搭配。

适合职业规划大赛、答辩、述职、年终汇报里「第一部分、第二部分……」的章节页。

## 让 AI 做

直接对 AI 说，例如：

> 用数字过渡页给我的职业规划做 3 个章节页：认识自我、探索职业、行动计划，标题用英文，下面配一句中文。

没说的内容 AI 会帮你补上，做完把 .pptx 和一个网页预览交给你。用 PowerPoint 或 WPS 打开，按 F5 放映，点一下翻一页。

## 自己运行

```bash
pip install -r ../../requirements.txt
python generate.py --images 1.jpg 2.jpg 3.jpg \
  --titles "KNOW|YOURSELF" "EXPLORE|OPTIONS" "MAKE|A PLAN" \
  --bodies "认识自己，找准方向" "走进行业，看清选择" "拆解目标，持续行动" \
  --out 章节页.pptx --preview 预览.html
```

有几张图就生成几页，`--numbers`、`--titles`、`--bodies`、`--colors` 都按页的顺序一一对应。

## 能换什么

下面这些都可以在对话里直接跟 AI 说，AI 会换成对应的参数。

| 参数 | 作用 | 默认值 |
|---|---|---|
| `--images` | 每页一张底图，横图最佳（16:9），建议 2K 以上。俯拍的风景、城市、纹理最好看 | 文件夹里的 `example1.jpg`～`example3.jpg` |
| `--numbers` | 每页的大数字，一般是 1 个字符 | `1` `2` `3` …… |
| `--titles` | 每页的标题，用 `\|` 分行。2 行最好看，每行不超过 9 个字母 | `KNOW\|YOURSELF` 等 3 组 |
| `--bodies` | 每页标题下的一句小字，10 个字左右最合适，也可以用 `\|` 分成 2 行。传 `""` 不显示 | 3 句职业规划文案 |
| `--colors` | 每页色块颜色（十六进制，不带 #） | 从每张图里自动取 |
| `--digit-font` | 大数字字体文件，衬线体最好看，会自动做成斜体 | `fonts/NotoSerifDisplay-Medium.ttf` |
| `--title-font` | 标题字体文件 | `fonts/Cinzel-Regular.ttf` |
| `--body-font` | 小字字体文件，写中文要用中文字体 | `fonts/NotoSansSC-Medium.ttf` |
| `--zoom-ms` | 底图缩回原大的时长，单位毫秒 | `2000` |
| `--fly-ms` | 色块飞入的时长，单位毫秒 | `1700` |
| `--fly-decel` | 飞入时减速的比例，0 是匀速，越接近 1 收尾越慢 | `0.9` |
| `--out` | 输出的 PPT 文件名 | `数字过渡页.pptx` |
| `--preview` | 同时输出一个网页预览 | 不输出 |

## 搭配建议

| 场景 | 标题示例 | 底图 |
|---|---|---|
| 职业规划大赛 | KNOW YOURSELF / EXPLORE OPTIONS / MAKE A PLAN | 花海、海岸、城市俯拍 |
| 年终汇报 | LOOKING BACK / KEY RESULTS / NEXT STEPS | 公司、城市、团队照片 |
| 励志演讲 | DREAM / BELIEVE / ACHIEVE | 山峰、日出、跑道 |

## 说明

- 所有文字都已转成形状，换一台没装这些字体的电脑打开，效果也一样；但打开后文字不能再直接改，要改就重新运行一次。
- 动画用的是「放大/缩小」强调动画和「飞入」进入动画，PowerPoint 2010 及以上、Microsoft 365、WPS 都能播放。
- 网页预览只用来看效果，正式放映以 PPT 为准。
