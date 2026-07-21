# Campus Poster Display Skill

一个可复用的 Agent Skill：把目录中的竖版海报批量转换为适合校区大屏、一体机或数字标牌播放的 16:9 PowerPoint。

每张海报完整居中显示，两侧使用同源模糊暗背景；PPTX 包含淡入淡出、定时自动翻页和循环播放设置，并在生成后检查关键 OOXML。

## Features

- JPG、JPEG、PNG 批量输入，扩展名不区分大小写
- EXIF 方向修正和透明像素安全铺白
- 前景海报完整等比显示，不拉伸、不裁切
- 1920×1080 模糊暗背景
- 自定义每页停留时间，默认 10 秒
- Fade 转场、禁用点击前进、使用计时、循环播放
- 生成后校验 16:9、页数、计时、转场和循环设置
- 对不适合远距离观看的超长海报给出拆页警告

## Install as an Agent Skill

```bash
git clone https://github.com/zhusq5-gif/campus-poster-display.git
mkdir -p ~/.codex/skills
cp -R campus-poster-display ~/.codex/skills/campus-poster-display
```

仓库根目录名与 `SKILL.md` 中的 `name: campus-poster-display` 一致，可直接作为 Agent Skill 使用。

## Python Setup

Requires Python 3.10+.

```bash
python3 -m pip install -r requirements.txt
```

## Usage

```bash
python3 scripts/generate_pptx.py \
  --input "/path/to/posters" \
  --output "/path/to/output/campus-posters.pptx" \
  --seconds 10
```

支持参数：

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `--input` | 当前目录 | 只读取目录第一层的 JPG/JPEG/PNG |
| `--output` | `<input>/artifacts/poster-display.pptx` | 输出 PPTX 路径 |
| `--seconds` | `10` | 每页停留秒数，必须大于 0 |

## Important Playback Note

脚本自检证明 PPTX 中已经写入计时和循环属性，但 PowerPoint、WPS、LibreOffice、Keynote 及不同数字标牌播放器的实现可能不同。正式使用前，请在目标电脑和屏幕上全屏播放至少一整轮。

如果播放器不能稳定遵循 PPT 计时，建议从 PowerPoint/WPS 导出 MP4，再使用稳定播放器全屏循环。

## Privacy and Content Safety

不要公开处理含有未授权学生姓名、可识别照片、学校、成绩单或其他敏感信息的海报。发布前应确认素材版权和展示授权。

## License

[MIT](LICENSE)
