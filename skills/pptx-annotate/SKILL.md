---
name: pptx-annotate
description: Annotate screenshots inside PowerPoint presentations with numbered markers. Use this skill whenever the user has a PPTX with screenshots and wants to add numbered callouts/circles ON the images, paired with numbered labels in text boxes next to the images. Also use when the user wants to fix Hebrew titles to be right-aligned, replace mixed Hebrew/English titles with Hebrew-only, or generally improve annotation layout in presentation slides. Trigger on phrases like "הוסף מספרים לתמונה", "annotations on screenshots", "numbered callouts", "circles on slides", "מספרים על הצילום מסך", "label the screenshots".
---

# PPTX Annotation Skill

Adds dual numbered markers to screenshot-based presentations:
- 🔴 **Red circles** next to text label boxes (outside image, visible on the right panel)
- 🔵 **Blue circles** on the image itself (overlaid on the screenshot, pointing to the relevant UI element)

Both use the same numbers so the viewer can match image location → text explanation.

---

## Workflow

### 1. Unpack

```bash
python3 /mnt/skills/public/pptx/scripts/office/unpack.py input.pptx /home/claude/unpacked/
```

### 2. Analyze structure

Read slide XMLs to understand:
- Image position (look for `<p:pic>` → `<a:xfrm>` → `<a:off>` and `<a:ext>`)
- Existing annotation circles (ellipse shapes with color `E53935`)
- Title text box (shape named `Text 1`)
- Text label boxes (white rect shapes to the right of the image)

```bash
python3 -c "
from defusedxml import minidom
EMU = 914400
doc = minidom.parse('/home/claude/unpacked/ppt/slides/slide2.xml')
ns_p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
for sp in doc.getElementsByTagNameNS(ns_p, 'sp'):
    cNvPr = sp.getElementsByTagName('p:cNvPr')
    # ... inspect shapes
"
```

### 3. Apply three transformations

Run a single Python script that does all three. See **Script Template** below.

**Transformation A — Hebrew-only title, right-aligned:**
- Find shape named `Text 1`
- Replace text with Hebrew-only version (strip English parts)
- Set `algn="r" rtl="1"` on `<a:pPr>`
- Set `rtlCol="1"` on `<a:bodyPr>`

**Transformation B — Add blue circles ON the image:**
- For each annotation, define fractional position `(fx, fy)` where `0.0` = top/left of image, `1.0` = bottom/right
- Convert to EMU: `x = img_x + fx * img_w - circle_r`, `y = img_y + fy * img_h - circle_r`
- Append blue circle + number XML before `</p:spTree>`
- Blue color: `1565C0`, white number text

**Transformation C — Fix red circle rendering order:**
- Red circles (ellipse, color `E53935`) must appear AFTER white label boxes in XML order
- Otherwise the white boxes render on top and hide the red circles
- Extract all red ellipse+number pairs, remove them, re-insert just before the blue circles

### 4. Fix text box overflow

Check if label boxes extend beyond slide right edge (slide width = 10"):
```python
# Typical overflow values:
OLD_BOX_X = "6400800"  # 7.0" 
OLD_BOX_W = "2880360"  # 3.15" → ends at 10.15" (overflow!)
# Fix:
NEW_BOX_X = str(int(6.45 * 914400))  # 5897880
NEW_BOX_W = str(int(3.45 * 914400))  # 3154680
```

### 5. Pack and QA

```bash
python3 /mnt/skills/public/pptx/scripts/office/pack.py /home/claude/unpacked/ output.pptx --original input.pptx

# Visual QA
python3 /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 130 output.pdf /home/claude/qa_slide
# Then view: /home/claude/qa_slide-02.jpg etc.
```

---

## Script Template

```python
import re, os
EMU = 914400

# Image boundaries (typical for this layout — verify per file)
IMG_X = 182880   # 0.2"
IMG_Y = 914400   # 1.0"
IMG_W = 5669280  # 6.2"
IMG_H = 3401568  # 3.72"
CIRCLE_D = 320040  # 0.35" diameter

BLUE_COLOR = "1565C0"
RED_COLOR  = "E53935"

def img_pos(fx, fy):
    """Fractional position on image → absolute EMU (top-left of circle)"""
    x = IMG_X + int(fx * IMG_W) - CIRCLE_D // 2
    y = IMG_Y + int(fy * IMG_H) - CIRCLE_D // 2
    return x, y

def make_circle_xml(x_emu, y_emu, number, shape_id, color):
    return f'''      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="{shape_id}" name="Circle_{color}_{number}"/>
          <p:cNvSpPr/><p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x_emu}" y="{y_emu}"/>
            <a:ext cx="{CIRCLE_D}" cy="{CIRCLE_D}"/></a:xfrm>
          <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
            <a:prstDash val="solid"/></a:ln>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/>
          <a:p><a:endParaRPr lang="en-IL"/></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="{shape_id+1}" name="Num_{color}_{number}"/>
          <p:cNvSpPr/><p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x_emu}" y="{y_emu}"/>
            <a:ext cx="{CIRCLE_D}" cy="{CIRCLE_D}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/><a:ln/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr marL="0" indent="0" algn="ctr"><a:buNone/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1100" b="1" dirty="0">
                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
                <a:latin typeface="Calibri" pitchFamily="34" charset="0"/>
              </a:rPr>
              <a:t>{number}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>'''

def fix_title(content, hebrew_title):
    sp_start = content.find('name="Text 1"')
    if sp_start == -1: return content
    sp_start = content.rfind('<p:sp>', 0, sp_start)
    sp_end = content.find('</p:sp>', sp_start) + len('</p:sp>')
    block = content[sp_start:sp_end]
    block = re.sub(r'<a:pPr marL="0" indent="0"([^>]*)>',
                   '<a:pPr marL="0" indent="0" algn="r" rtl="1">', block)
    block = block.replace('rtlCol="0"', 'rtlCol="1"')
    block = re.sub(r'<a:t>[^<]+</a:t>', f'<a:t>{hebrew_title}</a:t>', block, count=1)
    return content[:sp_start] + block + content[sp_end:]

def add_blue_circles(content, positions):
    ids = [int(m) for m in re.findall(r'<p:cNvPr id="(\d+)"', content)]
    next_id = max(ids) + 1 if ids else 100
    xml = ""
    for i, (fx, fy) in enumerate(positions):
        x, y = img_pos(fx, fy)
        xml += make_circle_xml(x, y, i+1, next_id, BLUE_COLOR) + "\n"
        next_id += 2
    return content.replace('    </p:spTree>', xml + '    </p:spTree>', 1)

def fix_red_circle_order(content):
    """Move red circles to AFTER white boxes so they render on top"""
    sp_pat = re.compile(r'      <p:sp>.*?</p:sp>', re.DOTALL)
    all_sps = list(sp_pat.finditer(content))
    pairs = []
    i = 0
    while i < len(all_sps):
        block = all_sps[i].group(0)
        if 'prst="ellipse"' in block and RED_COLOR in block:
            if i + 1 < len(all_sps):
                nxt = all_sps[i+1].group(0)
                if 'FFFFFF' in nxt and ('sz="1100"' in nxt or 'sz="1000"' in nxt):
                    pairs.append((block, nxt))
                    i += 2; continue
        i += 1
    if not pairs: return content
    for e, n in pairs:
        content = content.replace(e + '\n' + n, '', 1)
    blue_pos = content.find('name="Circle_1565C0_')
    if blue_pos != -1:
        sp_start = content.rfind('      <p:sp>', 0, blue_pos)
    else:
        sp_start = content.rfind('    </p:spTree>')
    insert = ''.join(e + '\n' + n + '\n' for e, n in pairs)
    return content[:sp_start] + insert + content[sp_start:]

# --- Per-slide data: (hebrew_title, [(fx, fy), ...]) ---
# fx, fy = fraction of image (0.0–1.0) where circle center should appear
SLIDE_DATA = {
    2: ("כניסה לקופיילוט", [(0.32, 0.62), (0.32, 0.85), (0.73, 0.62)]),
    # ... add more slides
}

SLIDES_DIR = '/home/claude/unpacked/ppt/slides'
for slide_num, (title, positions) in SLIDE_DATA.items():
    path = os.path.join(SLIDES_DIR, f'slide{slide_num}.xml')
    with open(path) as f: content = f.read()
    content = fix_title(content, title)
    content = add_blue_circles(content, positions)
    content = fix_red_circle_order(content)
    with open(path, 'w') as f: f.write(content)
    print(f"slide{slide_num}: {title} + {len(positions)} blue circles")
```

---

## Positioning Blue Circles

When determining `(fx, fy)` positions for each annotation, look at what the annotation text describes and estimate where that element appears in the screenshot:

| UI zone | Approximate fx, fy |
|---|---|
| Top toolbar / navbar | fy ≈ 0.05–0.15 |
| Left sidebar | fx ≈ 0.08–0.15 |
| Center content | fx ≈ 0.45–0.60 |
| Right panel / button | fx ≈ 0.80–0.92 |
| Bottom bar | fy ≈ 0.85–0.95 |
| Middle of screen | fx ≈ 0.50, fy ≈ 0.50 |

After the first render, do a visual QA pass and adjust any circles that land on the wrong element.

---

## Key Pitfalls

1. **Red circles hidden behind white boxes** — Always run `fix_red_circle_order()`. The ellipse shapes must appear after the white label boxes in XML order.

2. **Text box overflow** — Label boxes default to x=7.0" + w=3.15" = 10.15" which overflows a 10" slide. Fix by shifting to x=6.45" w=3.45".

3. **Title alignment with mixed Hebrew/English** — Setting `algn="r"` alone is not enough. Must also set `rtl="1"` on `<a:pPr>` AND `rtlCol="1"` on `<a:bodyPr>`. And the text should be Hebrew-only for reliable RTL rendering.

4. **Circle ordering in XML = z-order** — Last shape in XML renders on top. Blue circles (on image) are added last so they appear above the image. Red circles must come after white boxes.

5. **Pack with `--original`** — Always pass the original file to preserve media relationships and avoid orphaned images.
