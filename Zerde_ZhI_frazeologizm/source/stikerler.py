"""«Фразеостикерлер»: WhatsApp-қа арналған стикерлер (512×512, мөлдір фон, WebP < 100 КБ) және 96×96 белгіше.
python3 stikerler.py ../Fraze_stikerler"""
import io, os, sys, zipfile
from PIL import Image, ImageDraw, ImageFilter, ImageFont

OUT = sys.argv[1] if len(sys.argv) > 1 else "../Fraze_stikerler"
os.makedirs(OUT, exist_ok=True)
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
EMOJI = "/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf"
S = 512

# (тіркес чатқа ыңғайлы түрде, сөздіктегі мағынасына сүйенген қысқа түсінік, эмодзи, түс)
STICKERS = [
    ("Асығың алшысынан түссін!", "Жолың болсын!", "🍀🎉", "#FFD166"),
    ("Екі езуім екі құлағымда!", "Қатты қуаныштымын", "😁", "#FF9F68"),
    ("Арамыздан қыл өтпейді", "Нағыз доспыз", "🫂", "#F78FB3"),
    ("Қас пен көздің арасында!", "Лезде, тез арада", "⚡👀", "#7FD1F7"),
    ("Алақаныма түкірдім!", "Іске кірістім", "💪", "#9BE7A0"),
    ("Бармағымды тістедім…", "Өкініп тұрмын", "😖", "#B8A9F5"),
    ("Ит өлген жер екен!", "Өте алыс", "🧭🏔", "#F5C26B"),
    ("Құлағымды түрдім", "Тыңдап тұрмын", "👂", "#8FE3D8"),
    ("Қой аузынан шөп алмас", "Момын, жуас", "🐑😇", "#E0E7EE"),
    ("Бес саусағымдай білемін", "Жетік білемін", "🖐", "#FFB4A2"),
    ("Қол ұшын берейін", "Көмектесейін", "🤝", "#A5C8FF"),
    ("Жыл көрдің бе, ай көрдің бе?", "Көп ұйықтадың ғой!", "😴🌙", "#C9B6E4"),
]


def emoji_img(ch, size):
    f = ImageFont.truetype(EMOJI, 109)  # Noto Color Emoji тек 109 өлшемінде салынады
    im = Image.new("RGBA", (160, 160), (0, 0, 0, 0))
    ImageDraw.Draw(im).text((80, 80), ch, font=f, embedded_color=True, anchor="mm")
    im = im.crop(im.getbbox())
    return im.resize((size, round(size * im.height / im.width)), Image.LANCZOS)


def fit_text(draw, text, maxw, sizes):
    for sz in sizes:
        f = ImageFont.truetype(BOLD, sz)
        words, lines, cur = text.split(), [], ""
        for w in words:
            t = (cur + " " + w).strip()
            if draw.textlength(t, font=f) <= maxw:
                cur = t
            else:
                lines.append(cur); cur = w
        lines.append(cur)
        if len(lines) <= 3 and all(draw.textlength(l, font=f) <= maxw for l in lines):
            return f, lines
    return f, lines


def sticker(text, sub, emojis, color, tilt):
    big = 2  # екі есе үлкен салып, кейін кішірейтеміз – жиектер тегіс шығады
    W = S * big
    im = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    m, r = 34 * big, 70 * big
    box = (m, m, W - m, W - m)
    sh = Image.new("RGBA", (W, W), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle((box[0] + 6 * big, box[1] + 10 * big, box[2] + 6 * big, box[3] + 10 * big), r, fill=(0, 0, 0, 90))
    im.alpha_composite(sh.filter(ImageFilter.GaussianBlur(10 * big)))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle(box, r, fill="white")                                      # стикердің ақ жиегі
    inner = (box[0] + 13 * big, box[1] + 13 * big, box[2] - 13 * big, box[3] - 13 * big)
    d.rounded_rectangle(inner, r - 13 * big, fill=color)
    # эмодзи (бәрі бір биіктікте), тіркес және түсінік – картаның ортасында
    f, lines = fit_text(d, text, 372 * big, [s * big for s in (44, 41, 38, 35, 32)])
    EH = (128 if len(lines) < 3 else 104) * big
    em = []
    for e in emojis:
        x = emoji_img(e, 200 * big)
        k = min(EH / x.height, (EH + 22 * big if len(emojis) == 1 else EH) / x.width)
        em.append(x.resize((round(x.width * k), round(x.height * k)), Image.LANCZOS))
    eh = max(e.height for e in em)
    lh = round(f.size * 1.18)
    fs = ImageFont.truetype(BOLD, 25 * big)
    ph, g1, g2 = 50 * big, 16 * big, 18 * big
    total = eh + g1 + len(lines) * lh + g2 + ph
    y = (inner[1] + inner[3]) // 2 - total // 2
    gap = 10 * big
    x = (W - (sum(e.width for e in em) + gap * (len(em) - 1))) // 2
    for e in em:
        im.alpha_composite(e, (x, y + (eh - e.height) // 2)); x += e.width + gap
    y += eh + g1
    for i, l in enumerate(lines):
        d.text((W // 2, y + i * lh), l, font=f, fill="#1E1B2E", anchor="ma")
    y += len(lines) * lh + g2
    pw = d.textlength(sub, font=fs) + 44 * big
    d.rounded_rectangle(((W - pw) / 2, y, (W + pw) / 2, y + ph), ph // 2, fill="#1E1B2E")
    d.text((W // 2, y + ph // 2), sub, font=fs, fill="white", anchor="mm")
    im = im.rotate(tilt, resample=Image.BICUBIC)
    return im.resize((S, S), Image.LANCZOS)


zf = zipfile.ZipFile(os.path.join(OUT, "Fraze_stikerler_WhatsApp.zip"), "w", zipfile.ZIP_DEFLATED)
for i, (t, sub, emo, col) in enumerate(STICKERS, 1):
    im = sticker(t, sub, list(emo), col, [-3, 2.5, -2, 3][i % 4])
    name = f"{i:02d}"
    im.save(os.path.join(OUT, name + ".png"))
    buf = io.BytesIO()
    for q in (90, 80, 70, 60):
        buf = io.BytesIO(); im.save(buf, "WEBP", quality=q, method=6)
        if buf.tell() < 100_000:
            break
    open(os.path.join(OUT, name + ".webp"), "wb").write(buf.getvalue())
    zf.writestr(name + ".webp", buf.getvalue())
    print(name, t, f"{buf.tell() // 1024} КБ")
tray = Image.open(os.path.join(OUT, "01.png")).resize((96, 96), Image.LANCZOS)
tray.save(os.path.join(OUT, "tray.png"), optimize=True)
zf.write(os.path.join(OUT, "tray.png"), "tray.png")
zf.close()

# барлық стикерді бір суретте көрсететін парақ
sheet = Image.new("RGB", (4 * 300 + 40, 3 * 300 + 40), "#ECE5DD")  # WhatsApp чатының фоны
for i in range(len(STICKERS)):
    st = Image.open(os.path.join(OUT, f"{i + 1:02d}.png")).resize((300, 300), Image.LANCZOS)
    sheet.paste(st, (20 + (i % 4) * 300, 20 + (i // 4) * 300), st)
sheet.save(os.path.join(OUT, "Barlygy.png"))
print("done", OUT)
