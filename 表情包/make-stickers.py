# -*- coding: utf-8 -*-
"""大肥鱼表情包生成器：手写 SVG -> headless Edge 渲染 PNG。

用法：python make-stickers.py [输出目录]
"""
import os
import sys

OUT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__)))
W = H = 512
INK = "#1b1b1b"
BODY = "#9ed8f6"
BODY_D = "#7cc6ee"
BELLY = "#fff6e5"
BLUSH = "#ffb3c1"
PINK = "#ff8fa3"

# ---------- 可复用的身体零件 ----------


def body(tilt=0.0, belly=True):
    """圆滚滚的大肥鱼身体：尾巴 + 大圆身 + 肚皮 + 侧鳍 + 顶鳍 + 呆毛。"""
    return f"""
  <g transform="rotate({tilt} 256 300)">
    <!-- 尾巴 -->
    <path d="M404 300 C440 268 468 262 486 276 C470 292 470 308 486 324 C468 338 440 332 404 300 Z"
          fill="{BODY_D}" stroke="{INK}" stroke-width="8" stroke-linejoin="round"/>
    <!-- 头顶小鳍（小莫西干） -->
    <path d="M240 138 C252 110 276 96 294 116 C272 124 258 136 250 152 Z"
          fill="{BODY_D}" stroke="{INK}" stroke-width="8" stroke-linejoin="round"/>
    <!-- 呆毛 -->
    <path d="M206 148 C200 122 216 104 232 112 C218 122 214 134 220 152"
          fill="none" stroke="{INK}" stroke-width="8" stroke-linecap="round"/>
    <!-- 身体 -->
    <ellipse cx="252" cy="300" rx="168" ry="152" fill="{BODY}" stroke="{INK}" stroke-width="9"/>
    <!-- 肚皮 -->
    {"<ellipse cx='244' cy='340' rx='112' ry='94' fill='" + BELLY + "' stroke='" + INK + "' stroke-width='7'/>" if belly else ""}
    <!-- 两侧小鳍 -->
    <path d="M104 300 C74 292 52 306 50 326 C70 336 94 330 110 316 Z"
          fill="{BODY_D}" stroke="{INK}" stroke-width="8" stroke-linejoin="round"/>
    <path d="M400 336 C428 328 452 342 456 362 C436 372 412 364 396 350 Z"
          fill="{BODY_D}" stroke="{INK}" stroke-width="8" stroke-linejoin="round"/>
  </g>"""


def eyes(kind="happy"):
    """豆豆眼：happy=弯月笑眼 / open=大圆豆豆眼 / sparkle=星星眼 / closed=眯眼（上凸弧）。"""
    if kind == "happy":
        return f"""
    <path d="M162 252 C176 226 214 226 228 252" fill="none" stroke="{INK}"
          stroke-width="13" stroke-linecap="round"/>
    <path d="M284 252 C298 226 336 226 350 252" fill="none" stroke="{INK}"
          stroke-width="13" stroke-linecap="round"/>"""
    if kind == "closed":
        return f"""
    <path d="M168 246 C182 230 214 230 228 246" fill="none" stroke="{INK}"
          stroke-width="13" stroke-linecap="round"/>
    <path d="M284 246 C298 230 330 230 344 246" fill="none" stroke="{INK}"
          stroke-width="13" stroke-linecap="round"/>"""
    if kind == "open":
        return f"""
    <circle cx="198" cy="240" r="31" fill="#ffffff" stroke="{INK}" stroke-width="7"/>
    <circle cx="198" cy="240" r="21" fill="{INK}"/>
    <circle cx="186" cy="226" r="10" fill="#ffffff"/>
    <circle cx="208" cy="254" r="5" fill="#ffffff"/>
    <circle cx="318" cy="240" r="31" fill="#ffffff" stroke="{INK}" stroke-width="7"/>
    <circle cx="318" cy="240" r="21" fill="{INK}"/>
    <circle cx="306" cy="226" r="10" fill="#ffffff"/>
    <circle cx="328" cy="254" r="5" fill="#ffffff"/>"""
    # sparkle：星星眼
    return f"""
    <circle cx="198" cy="240" r="31" fill="#ffffff" stroke="{INK}" stroke-width="7"/>
    <circle cx="198" cy="240" r="22" fill="{INK}"/>
    <path d="M198 219 L206 234 L221 241 L206 248 L198 263 L190 248 L175 241 L190 234 Z"
          fill="#ffffff"/>
    <circle cx="318" cy="240" r="31" fill="#ffffff" stroke="{INK}" stroke-width="7"/>
    <circle cx="318" cy="240" r="22" fill="{INK}"/>
    <path d="M318 219 L326 234 L341 241 L326 248 L318 263 L310 248 L295 241 L310 234 Z"
          fill="#ffffff"/>"""


def blush():
    return f"""
    <ellipse cx="132" cy="322" rx="29" ry="18" fill="{BLUSH}" opacity="0.95"/>
    <ellipse cx="384" cy="322" rx="29" ry="18" fill="{BLUSH}" opacity="0.95"/>"""


def mouth(kind="w"):
    if kind == "w":
        return f"""
    <path d="M236 318 C246 332 258 332 268 318 C278 332 290 332 300 318"
          fill="none" stroke="{INK}" stroke-width="9" stroke-linecap="round"/>"""
    if kind == "open_o":
        return f"""
    <ellipse cx="268" cy="324" rx="25" ry="21" fill="#ff8a9a" stroke="{INK}" stroke-width="8"/>"""
    if kind == "grin":
        return f"""
    <path d="M222 312 C244 352 292 352 314 312 Z" fill="#ff8a9a" stroke="{INK}"
          stroke-width="8" stroke-linejoin="round"/>
    <path d="M236 316 L250 334 L264 316 L278 334 L292 316" fill="none" stroke="{INK}"
          stroke-width="6" stroke-linecap="round"/>"""
    # 波浪嘴（委屈）
    return f"""
    <path d="M230 340 C244 322 258 322 272 340 C286 358 300 358 314 340"
          fill="none" stroke="{INK}" stroke-width="9" stroke-linecap="round"/>"""


def text_line(txt, y=478, size=52, fill="#1b1b1b"):
    """贴纸描边字：粗白描边的黑体字。"""
    return f"""
  <text x="256" y="{y}" text-anchor="middle" font-size="{size}"
        font-family="'Microsoft YaHei UI','Microsoft YaHei','Noto Sans SC',sans-serif"
        font-weight="900" fill="{fill}" stroke="#ffffff" stroke-width="10"
        paint-order="stroke" stroke-linejoin="round">{txt}</text>"""


def svg(inner):
    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">
  <title>大肥鱼表情包</title>{inner}
</svg>
"""


# ---------- 六张表情 ----------

FISH = body()
E = eyes
B = blush
M = mouth
T = text_line

stickers = {}

stickers["01_本鱼超棒的"] = svg(f"""
  <g transform="translate(0,-14)">{FISH}
    {E("happy")}{B()}{M("grin")}
    <path d="M124 150 L134 174 L158 184 L134 194 L124 218 L114 194 L90 184 L114 174 Z" fill="#ffe066" stroke="{INK}" stroke-width="6"/>
    <path d="M388 132 L396 152 L416 160 L396 168 L388 188 L380 168 L360 160 L380 152 Z" fill="#ffe066" stroke="{INK}" stroke-width="6"/>
  </g>
  {T("本鱼超棒的！")}""")

stickers["02_饿饿饭饭"] = svg(f"""
  <g transform="translate(0,-14)">{FISH}
    {E("open")}{B()}{M("wave")}
    <path d="M196 268 C214 276 219 302 210 322 C201 342 187 342 180 323 C172 302 179 278 196 268 Z"
          fill="#7fcdf5" stroke="{INK}" stroke-width="6"/>
    <path d="M336 268 C354 276 359 302 350 322 C341 342 327 342 320 323 C312 302 319 278 336 268 Z"
          fill="#7fcdf5" stroke="{INK}" stroke-width="6"/>
  </g>
  {T("饿饿，饭饭")}""")

stickers["03_大白饭喂本鱼"] = svg(f"""
  <g transform="translate(0,-12)">{FISH}
    {E("sparkle")}{B()}{M("open_o")}
    <!-- 空碗 -->
    <path d="M276 156 C278 194 302 210 332 210 C362 210 386 194 388 156 Z"
          fill="#ffffff" stroke="{INK}" stroke-width="8" stroke-linejoin="round"/>
    <path d="M272 156 L392 156" stroke="{INK}" stroke-width="8" stroke-linecap="round"/>
    <path d="M300 150 C312 132 352 132 364 150" fill="none" stroke="{INK}" stroke-width="6" stroke-linecap="round"/>
    <path d="M290 214 C288 236 284 246 276 258" fill="none" stroke="{INK}" stroke-width="7" stroke-linecap="round"/>
    <path d="M336 214 L344 256" fill="none" stroke="{INK}" stroke-width="7" stroke-linecap="round"/>
    <ellipse cx="236" cy="358" rx="15" ry="10" fill="#ffffff" stroke="{INK}" stroke-width="5"/>
    <ellipse cx="296" cy="374" rx="13" ry="9" fill="#ffffff" stroke="{INK}" stroke-width="5"/>
    <path d="M116 150 L124 172 L146 180 L124 188 L116 210 L108 188 L86 180 L108 172 Z" fill="{PINK}"/>
    <path d="M424 196 L430 212 L446 218 L430 224 L424 240 L418 224 L402 218 L418 212 Z" fill="{PINK}"/>
  </g>
  {T("大白饭！喂本鱼！", size=44)}""")

stickers["04_本鱼在工作"] = svg(f"""
  <g transform="translate(0,-14)">{FISH}
    {E("open")}{B()}{M("w")}
    <!-- 小笔记本 -->
    <rect x="186" y="382" width="164" height="18" rx="8" fill="#d0d0d0" stroke="{INK}" stroke-width="7"/>
    <rect x="176" y="356" width="184" height="28" rx="9" fill="#e4e4e4" stroke="{INK}" stroke-width="7"/>
    <path d="M196 370 L238 370" stroke="{INK}" stroke-width="5" stroke-linecap="round" opacity="0.55"/>
    <path d="M196 380 L226 380" stroke="{INK}" stroke-width="5" stroke-linecap="round" opacity="0.55"/>
    <!-- 打字的小鳍 -->
    <path d="M186 340 C164 330 146 340 148 356 C160 372 184 366 194 350 Z"
          fill="{BODY_D}" stroke="{INK}" stroke-width="7" stroke-linejoin="round"/>
    <path d="M352 340 C374 330 392 340 390 356 C378 372 354 366 344 350 Z"
          fill="{BODY_D}" stroke="{INK}" stroke-width="7" stroke-linejoin="round"/>
    <!-- 灵感小火花 -->
    <path d="M420 224 L426 240 L442 246 L426 252 L420 268 L414 252 L398 246 L414 240 Z" fill="#ffe066" stroke="{INK}" stroke-width="5"/>
  </g>
  {T("本鱼在工作哦")}""")

stickers["05_交给本鱼啦"] = svg(f"""
  <g transform="translate(0,-14)">{FISH}
    {E("happy")}{B()}{M("grin")}
    <!-- 灯泡 -->
    <path d="M128 96 L128 138" stroke="{INK}" stroke-width="7" stroke-linecap="round"/>
    <path d="M128 120 C150 116 162 134 156 152 C151 168 136 176 122 172 C108 168 100 154 104 138 C107 127 116 121 128 120 Z"
          fill="#fff2a8" stroke="{INK}" stroke-width="7" stroke-linejoin="round"/>
    <path d="M116 158 L134 154" stroke="{INK}" stroke-width="6" stroke-linecap="round"/>
    <path d="M418 214 L426 234 L446 242 L426 250 L418 270 L410 250 L390 242 L410 234 Z" fill="#ffe066" stroke="{INK}" stroke-width="5"/>
  </g>
  {T("交给本鱼啦～")}""")

stickers["06_冲鸭本鱼来了"] = svg(f"""
  <g transform="translate(-14,-8) scale(0.88) translate(6,-6)">{FISH}
    {E("open")}{B()}{M("open_o")}
    <path d="M212 200 C162 202 118 204 86 206" fill="none" stroke="#ffffff" stroke-width="14" stroke-linecap="round" opacity="0.95"/>
    <path d="M220 268 C170 270 126 272 94 274" fill="none" stroke="#ffffff" stroke-width="14" stroke-linecap="round" opacity="0.95"/>
    <path d="M212 336 C164 338 124 340 92 342" fill="none" stroke="#ffffff" stroke-width="14" stroke-linecap="round" opacity="0.95"/>
    <circle cx="64" cy="160" r="9" fill="#ffffff" opacity="0.8"/>
    <circle cx="48" cy="392" r="7" fill="#ffffff" opacity="0.7"/>
  </g>
  {T("冲鸭！本鱼来了！", size=42)}""")

os.makedirs(OUT, exist_ok=True)
for name, content in stickers.items():
    path = os.path.join(OUT, name + ".svg")
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    print("wrote", path)
print("total", len(stickers))
