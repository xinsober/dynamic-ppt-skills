# 镂空开场 · hollow-opening

![效果](preview.gif)

3 页的电影感开场，页与页之间用「平滑」切换：

1. 全黑。
2. 一行镂空大字从中间切成上下两半，上半从右、下半从左滑入，拼成完整的一行，透过字能看到底图。
3. 上下两半黑色分别移出画面，整张底图露出来并稍微放大，标题、副标题、署名淡入。

适合职业规划大赛、演讲、开学或毕业典礼、年终汇报、年会的第一页。

## 让 AI 做

直接对 AI 说，例如：

> 用镂空开场做一页年终汇报开场，大字 2026，标题「凝心聚力 再创辉煌」，汇报人写张三，用这张图。

没说的内容 AI 会帮你补上，做完把 .pptx 和一个网页预览交给你。用 PowerPoint 或 WPS 打开，按 F5 放映，点一下翻一页。

## 自己运行

```bash
pip install -r ../../requirements.txt
python generate.py --word "MY FUTURE" --title "规划未来 职引人生" \
  --subtitle "大学生职业生涯规划大赛" --reporter "参赛选手：XXX　XX学院" \
  --image 我的图.jpg --out 职业规划.pptx --preview 预览.html
```

## 能换什么

下面这些都可以在对话里直接跟 AI 说，AI 会换成对应的参数。

| 参数 | 作用 | 默认值 |
|---|---|---|
| `--word` | 镂空大字。英文 6～10 个字母效果最好，中文建议 4 个字，数字也可以 | `MY FUTURE` |
| `--title` | 第 3 页大标题，两句各 4 个字最稳。传空字符串 `""` 不显示 | `规划未来 职引人生` |
| `--subtitle` | 副标题，10～16 个字最合适。传 `""` 不显示 | `大学生职业生涯规划大赛` |
| `--reporter` | 署名条，例如「汇报人：XXX」「主讲人：XXX」。传 `""` 不显示 | `汇报人：XXX` |
| `--image` | 底图，横图最佳（16:9 左右），建议 2K 以上 | `example.jpg` |
| `--zoom` | 第 2 页底图放大倍数，越大字里露出的画面越局部 | `1.14` |
| `--image-x` | 底图左右取景，-1 露最左边，0 居中，1 露最右边 | `0` |
| `--image-y` | 底图上下取景，-1 露最上面，0 居中，1 露最下面 | `-0.15` |
| `--word-font` | 大字字体文件 | `fonts/Anton-Regular.ttf` |
| `--title-font` | 标题字体文件 | `fonts/MaShanZheng-Regular.ttf` |
| `--subtitle-font` | 副标题字体文件 | `fonts/NotoSerifSC-Bold.ttf` |
| `--reporter-font` | 署名字体文件 | `fonts/NotoSansSC-Medium.ttf` |
| `--subtitle-color` | 副标题颜色，可以按底图色调换 | `#F2CF5B`（金色） |
| `--duration` | 平滑切换时长，单位毫秒 | `2000` |
| `--out` | 输出的 PPT 文件名 | `镂空开场.pptx` |
| `--preview` | 同时输出一个网页预览 | 不输出 |

## 搭配建议

| 场景 | 大字 | 标题 | 大字字体 |
|---|---|---|---|
| 职业规划大赛 | `MY FUTURE` | 规划未来 职引人生 | Anton |
| 励志演讲 | `DREAM BIG` | 心向远方 步履不停 | BigShouldersDisplay-Black |
| 年终汇报 | `2026` | 凝心聚力 再创辉煌 | ArchivoBlack-Regular |
| 中文大字 | `向光而行` | 不负韶华 砥砺前行 | NotoSansSC-Black |

底图选上半是天空、下半有城市或山景的横图，镂空字里能同时透出天色和轮廓，效果最好。

## 说明

- 所有文字都已转成形状，换一台没装这些字体的电脑打开，效果也一样；但打开后文字不能再直接改，要改就重新运行一次。
- 平滑切换需要 PowerPoint 2019 及以上、Microsoft 365 或 WPS。旧版 PowerPoint 只能看到直接切页。
- 网页预览只用来看效果，正式放映以 PPT 为准。
