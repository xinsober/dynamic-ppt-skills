---
name: dynamic-ppt-skills
description: 生成动态 PPT（.pptx），包括平滑切换、动画等效果。用户想做电影感开场页、动态封面、炫酷转场，或提到职业规划大赛、演讲、年终汇报的动态 PPT 时使用。
---

# 动态 PPT 模板

每个模板是 `templates/<模板名>/generate.py` 一个程序，运行后直接输出 .pptx，动态效果用 PowerPoint / WPS 自带的切换和动画实现。

## 模板列表

| 模板 | 文件夹 | 效果 | 什么时候用 |
|---|---|---|---|
| 镂空开场 | `templates/hollow-opening` | 黑屏 → 镂空大字上下两半滑入拼合 → 黑色移出露出全图和标题，共 3 页 | 开场页、封面页，需要一行大字 + 一行标题 |
| 圆环连招 | `templates/ring-combo` | 四个圆环透出底图：开场 → 封面 → 目录 → 4 个章节页，大圆环每页转 90°，共 7 页 | 要一整套封面 + 目录 + 4 个章节；需要 1 张图、4 个章节名 |
| 旋转转盘 | `templates/rotary-dial` | 左边半个转盘每页转 90°，同时换底图、文字上下滚动，共 4 页 | 分 4 部分讲、每页想换画面；需要 4 张图（不够会循环用） |
| 多图轮播 | `templates/photo-carousel` | 上下两条弧形色带夹着一排照片，进入页面后照片自动从右往左匀速滚动，共 1 页 | 活动照片多、想一页展示完；需要 4～10 张照片、一句标题 |
| 数字过渡页 | `templates/number-transition` | 整页铺图，左边一个超大斜体数字，右边斜边半透明色块配标题；底图缓缓缩回，数字、色块、标题从右飞入，每节 1 页 | 章节过渡页；每节需要 1 张图、一个短英文标题 |
| 唱片机结束页 | `templates/record-ending` | 左边圆环透出底图像唱片一直转，唱臂搭在上面，右边毛笔标题和小字像片尾字幕一样往上滚，共 1 页 | 结束页、感谢页；需要 1 张图、一个短标题、2～3 段短句 |

## 调用步骤

1. 安装依赖（只需一次）：`pip install -r requirements.txt`
2. 读对应模板文件夹里的 `README.md`，确认有哪些参数、推荐字数。
3. 从用户的要求里整理出参数。用户没给的内容：
   - 文字：按场景写简短有力的文案，并告诉用户可以改；
   - 图片：问用户要横图（16:9 左右，2K 以上；多图轮播要多张照片）。用户没有就先用模板文件夹里自带的示例图（`example*.jpg`），并说明之后可以替换。
4. 运行，同时输出网页预览：
   ```bash
   python templates/hollow-opening/generate.py --word "MY FUTURE" --title "规划未来 职引人生" \
     --subtitle "大学生职业生涯规划大赛" --reporter "参赛选手：XXX" \
     --image 用户的图.jpg --out 开场.pptx --preview 预览.html
   ```
5. 把 .pptx 交给用户，告诉他们用 PowerPoint 2019+ / Microsoft 365 / WPS 打开后按 F5 放映（模板 README 里写了放映要求）。

## 注意

- 文字转成了形状，生成后在 PPT 里不能直接改字；要改就改参数重新生成。
- 大字字体、标题字体在 `fonts/` 里选（见 `fonts/README.md`），也可以传用户自己的 .ttf / .otf 路径。中文大字要用 `fonts/NotoSansSC-Black.ttf` 这类粗黑体。
- 所有参数都可以用 `python templates/<模板名>/generate.py --help` 查看。
