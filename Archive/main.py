#!/usr/bin/env python3
"""
Generate a PPTX presentation guiding Tamil speakers through
Japanese Hiragana and Katakana, including dakuten (が, ぎ, ぐ...) 
and handakuten (ぱ, ぴ, ぷ...), with improved sizing so slides
do not overflow in a 16:9 layout.

We use 'Option 2': Hardcode the slide width & height in inches
(13.333" x 7.5") instead of retrieving them from the Presentation object.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

################################################################
# Hard-coded slide dimensions in inches (16:9 ratio)
################################################################
SLIDE_WIDTH_INCHES = 13.333
SLIDE_HEIGHT_INCHES = 7.5

################################################################
# 1) Basic (Non-Dakuten) Tables
################################################################
HIRAGANA_TABLE = [
  ["あ", "か", "さ", "た", "な", "は", "ま", "や", "ら", "わ"],
  ["い", "き", "し", "ち", "に", "ひ", "み", "",   "り", "" ],
  ["う", "く", "す", "つ", "ぬ", "ふ", "む", "ゆ", "る", "" ],
  ["え", "け", "せ", "て", "ね", "へ", "め", "",   "れ", "" ],
  ["お", "こ", "そ", "と", "の", "ほ", "も", "よ", "ろ", "を"],
]
HIRAGANA_N = "ん"

KATAKANA_TABLE = [
  ["ア", "カ", "サ", "タ", "ナ", "ハ", "マ", "ヤ", "ラ", "ワ"],
  ["イ", "キ", "シ", "チ", "ニ", "ヒ", "ミ", "",   "リ", "" ],
  ["ウ", "ク", "ス", "ツ", "ヌ", "フ", "ム", "ユ", "ル", "" ],
  ["エ", "ケ", "セ", "テ", "ネ", "ヘ", "メ", "",   "レ", "" ],
  ["オ", "コ", "ソ", "ト", "ノ", "ホ", "モ", "ヨ", "ロ", "ヲ"],
]
KATAKANA_N = "ン"

################################################################
# 2) Dakuten/Handakuten Tables
################################################################
HIRAGANA_DAKUTEN_TABLE = [
  ["が", "ぎ", "ぐ", "げ", "ご"],
  ["ざ", "じ", "ず", "ぜ", "ぞ"],
  ["だ", "ぢ", "づ", "で", "ど"],
  ["ば", "び", "ぶ", "べ", "ぼ"],
  ["ぱ", "ぴ", "ぷ", "ぺ", "ぽ"],  # handakuten row
]

KATAKANA_DAKUTEN_TABLE = [
  ["ガ", "ギ", "グ", "ゲ", "ゴ"],
  ["ザ", "ジ", "ズ", "ゼ", "ゾ"],
  ["ダ", "ヂ", "ヅ", "デ", "ド"],
  ["バ", "ビ", "ブ", "ベ", "ボ"],
  ["パ", "ピ", "プ", "ペ", "ポ"],
]

################################################################
# 3) Romaji + Tamil Mappings
################################################################
ROMAJI_TAMIL_MAP = {
    # Basic Hiragana
    "あ": ("A","அ"),  "か": ("Ka","க"),  "さ": ("Sa","ச"),  "た": ("Ta","த"),  "な": ("Na","ந"),
    "は": ("Ha","ஹ"), "ま": ("Ma","ம"),  "や": ("Ya","ய"),  "ら": ("Ra","ற"),  "わ": ("Wa","வ"),
    "い": ("I","இ"),  "き": ("Ki","கி"), "し": ("Shi","சி"),"ち": ("Chi","தி"),"に": ("Ni","நி"),
    "ひ": ("Hi","ஹி"),"み": ("Mi","மி"), "り": ("Ri","றி"),
    "う": ("U","உ"),  "く": ("Ku","கு"), "す": ("Su","சு"), "つ": ("Tsu","து"),"ぬ": ("Nu","நு"),
    "ふ": ("Fu","ஹு"),"む": ("Mu","மு"), "ゆ": ("Yu","யு"), "る": ("Ru","று"),
    "え": ("E","எ"),  "け": ("Ke","கே"),"せ": ("Se","செ"), "て": ("Te","தே"),"ね": ("Ne","நே"),
    "へ": ("He","ஹே"),"め": ("Me","மே"), "れ": ("Re","றே"),
    "お": ("O","ஒ"),  "こ": ("Ko","கொ"),"そ": ("So","சொ"), "と": ("To","தொ"),"の": ("No","நொ"),
    "ほ": ("Ho","ஹொ"),"も": ("Mo","மொ"), "よ": ("Yo","யொ"), "ろ": ("Ro","றொ"), "を": ("Wo","வொ"),
    "ん": ("N","ன்"),

    # Basic Katakana
    "ア": ("A","அ"),  "カ": ("Ka","க"),  "サ": ("Sa","ச"),  "タ": ("Ta","த"),  "ナ": ("Na","ந"),
    "ハ": ("Ha","ஹ"), "マ": ("Ma","ம"),  "ヤ": ("Ya","ய"),  "ラ": ("Ra","ற"),  "ワ": ("Wa","வ"),
    "イ": ("I","இ"),  "キ": ("Ki","கி"), "シ": ("Shi","சி"),"チ": ("Chi","தி"),"ニ": ("Ni","நி"),
    "ヒ": ("Hi","ஹி"),"ミ": ("Mi","மி"), "リ": ("Ri","றி"),
    "ウ": ("U","உ"),  "ク": ("Ku","கு"), "ス": ("Su","சு"), "ツ": ("Tsu","து"),"ヌ": ("Nu","நு"),
    "フ": ("Fu","ஹு"),"ム": ("Mu","மு"), "ユ": ("Yu","யு"), "ル": ("Ru","று"),
    "エ": ("E","எ"),  "ケ": ("Ke","கே"),"セ": ("Se","செ"), "テ": ("Te","தே"),"ネ": ("Ne","நே"),
    "ヘ": ("He","ஹே"),"メ": ("Me","மே"), "レ": ("Re","றே"),
    "オ": ("O","ஒ"),  "コ": ("Ko","கொ"),"ソ": ("So","சொ"), "ト": ("To","தொ"),"ノ": ("No","நொ"),
    "ホ": ("Ho","ஹொ"),"モ": ("Mo","மொ"), "ヨ": ("Yo","யொ"), "ロ": ("Ro","றொ"), "ヲ": ("Wo","வொ"),
    "ン": ("N","ன்"),

    # Hiragana Dakuten
    "が": ("Ga","க"), "ぎ": ("Gi","கி"), "ぐ": ("Gu","கு"), "げ": ("Ge","கே"), "ご": ("Go","கொ"),
    "ざ": ("Za","ச"), "じ": ("Ji","சி"), "ず": ("Zu","சு"), "ぜ": ("Ze","சே"), "ぞ": ("Zo","சொ"),
    "だ": ("Da","த"), "ぢ": ("Ji","தி"), "づ": ("Zu","து"), "で": ("De","தே"), "ど": ("Do","தொ"),
    "ば": ("Ba","ப"), "び": ("Bi","பி"), "ぶ": ("Bu","பு"), "べ": ("Be","பே"), "ぼ": ("Bo","பொ"),
    "ぱ": ("Pa","ப"), "ぴ": ("Pi","பி"), "ぷ": ("Pu","பு"), "ぺ": ("Pe","பே"), "ぽ": ("Po","பொ"),

    # Katakana Dakuten
    "ガ": ("Ga","க"), "ギ": ("Gi","கி"), "グ": ("Gu","கு"), "ゲ": ("Ge","கே"), "ゴ": ("Go","கொ"),
    "ザ": ("Za","ச"), "ジ": ("Ji","சி"), "ズ": ("Zu","சு"), "ゼ": ("Ze","சே"), "ゾ": ("Zo","சொ"),
    "ダ": ("Da","த"), "ヂ": ("Ji","தி"), "ヅ": ("Zu","து"), "デ": ("De","தே"), "ド": ("Do","தொ"),
    "バ": ("Ba","ப"), "ビ": ("Bi","பி"), "ブ": ("Bu","பு"), "ベ": ("Be","பே"), "ボ": ("Bo","பொ"),
    "パ": ("Pa","ப"), "ピ": ("Pi","பி"), "プ": ("Pu","பு"), "ペ": ("Pe","பே"), "ポ": ("Po","பொ"),
}

################################################################
# 4) Series for Focus Slides
################################################################
GOJUON_SERIES = [
    ("A Series",  ["あ", "い", "う", "え", "お"],  ["ア", "イ", "ウ", "エ", "オ"]),
    ("Ka Series", ["か", "き", "く", "け", "こ"],  ["カ", "キ", "ク", "ケ", "コ"]),
    ("Sa Series", ["さ", "し", "す", "せ", "そ"],  ["サ", "シ", "ス", "セ", "ソ"]),
    ("Ta Series", ["た", "ち", "つ", "て", "と"],  ["タ", "チ", "ツ", "テ", "ト"]),
    ("Na Series", ["な", "に", "ぬ", "ね", "の"],  ["ナ", "ニ", "ヌ", "ネ", "ノ"]),
    ("Ha Series", ["は", "ひ", "ふ", "へ", "ほ"],  ["ハ", "ヒ", "フ", "ヘ", "ホ"]),
    ("Ma Series", ["ま", "み", "む", "め", "も"],  ["マ", "ミ", "ム", "メ", "モ"]),
    ("Ya Series", ["や", "ゆ", "よ"],             ["ヤ", "ユ", "ヨ"]),
    ("Ra Series", ["ら", "り", "る", "れ", "ろ"],  ["ラ", "リ", "ル", "レ", "ロ"]),
    ("Wa/N Series",["わ", "を", "ん"],            ["ワ", "ヲ", "ン"]),
]

DAKUTEN_SERIES = [
    ("Ga Series", ["が", "ぎ", "ぐ", "げ", "ご"], ["ガ", "ギ", "グ", "ゲ", "ゴ"]),
    ("Za Series", ["ざ", "じ", "ず", "ぜ", "ぞ"], ["ザ", "ジ", "ズ", "ゼ", "ゾ"]),
    ("Da Series", ["だ", "ぢ", "づ", "で", "ど"], ["ダ", "ヂ", "ヅ", "デ", "ド"]),
    ("Ba Series", ["ば", "び", "ぶ", "べ", "ぼ"], ["バ", "ビ", "ブ", "ベ", "ボ"]),
    ("Pa Series", ["ぱ", "ぴ", "ぷ", "ぺ", "ぽ"], ["パ", "ピ", "プ", "ペ", "ポ"]),
]

################################################################
# 5) Helper Functions
################################################################
def create_table_for_syllabary(
    slide,
    table_data,
    slide_title,
    top=1.0,
    left=0.5,
    col_width=1.0,
    row_height=1.2,
    font_main=30,
    font_sub=14
):
    """
    Creates a table on the given slide. We do NOT retrieve slide dimensions
    from 'slide' or 'prs' (to avoid the .presentation error). Instead, we rely on
    SLIDE_WIDTH_INCHES / SLIDE_HEIGHT_INCHES for auto-scaling.

    Each cell shows:
      - Big script char on top
      - Smaller line with Romaji | Tamil below
    """
    # Hard-coded 16:9 slide inches
    slide_width_in = SLIDE_WIDTH_INCHES
    slide_height_in = SLIDE_HEIGHT_INCHES

    rows = len(table_data)
    cols = max(len(r) for r in table_data) if rows > 0 else 0

    # Title at top
    title_shape = slide.shapes.add_textbox(
        Inches(left),
        Inches(top),
        Inches(col_width * cols),
        Inches(0.5)
    )
    tf = title_shape.text_frame
    tf.text = slide_title
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    tf.paragraphs[0].runs[0].font.size = Pt(font_main + 6)
    tf.paragraphs[0].runs[0].font.bold = True

    # Start the table below the title
    table_top = top + 0.7
    table_left = left
    table_width = col_width * cols
    table_height = row_height * rows

    # Simple scaling approach if table is too wide:
    margin_w = 1.0  # total left+right margin in inches
    available_w = slide_width_in - margin_w
    if (table_width > available_w):
        scale_factor = available_w / table_width
        col_width *= scale_factor
        table_width = col_width * cols

    # Similarly for height
    margin_h = 1.5  # top + bottom margin
    available_h = slide_height_in - margin_h
    if (table_height > available_h):
        scale_factor_h = available_h / table_height
        row_height *= scale_factor_h
        table_height = row_height * rows

    # Now create the actual table
    graphic_frame = slide.shapes.add_table(
        rows, cols,
        Inches(table_left),
        Inches(table_top),
        Inches(table_width),
        Inches(table_height)
    )
    table = graphic_frame.table

    # Set uniform column widths & row heights
    for col_idx in range(cols):
        table.columns[col_idx].width = Inches(col_width)
    for row_idx in range(rows):
        table.rows[row_idx].height = Inches(row_height)

    # Fill each cell
    for r_idx, row_data in enumerate(table_data):
        for c_idx, char in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)
            if not char:
                cell.text = ""
                continue

            romaji, tamil = ROMAJI_TAMIL_MAP.get(char, ("", ""))
            cell.text_frame.clear()
            cell.text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

            # 1) Big character
            p1 = cell.text_frame.add_paragraph()
            p1.text = char
            p1.alignment = PP_ALIGN.CENTER
            run1 = p1.runs[0]
            run1.font.size = Pt(font_main)
            run1.font.bold = True

            # 2) Smaller line: Romaji | Tamil
            p2 = cell.text_frame.add_paragraph()
            p2.text = f"{romaji}  |  {tamil}"
            p2.alignment = PP_ALIGN.CENTER
            run2 = p2.runs[0]
            run2.font.size = Pt(font_sub)
            run2.font.bold = False

def create_centered_textbox(
    slide,
    text,
    left,
    top,
    width,
    height,
    font_size=60,
    bold=True,
    alignment=PP_ALIGN.CENTER
):
    """
    Creates a textbox shape on the given slide, centered horizontally & vertically,
    with the given font size & boldness.
    """
    textbox = slide.shapes.add_textbox(
        Inches(left), Inches(top),
        Inches(width), Inches(height)
    )
    text_frame = textbox.text_frame
    text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    p = text_frame.paragraphs[0]
    p.alignment = alignment

    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)  # black text

    return textbox

################################################################
# 6) Main Generation Logic
################################################################
def main():
    prs = Presentation()
    # Hard-code 16:9: 13.333" x 7.5" 
    prs.slide_width = Inches(SLIDE_WIDTH_INCHES)
    prs.slide_height = Inches(SLIDE_HEIGHT_INCHES)

    blank_layout = prs.slide_layouts[6]  # blank layout

    # 1) Basic Hiragana Overview
    slide1 = prs.slides.add_slide(blank_layout)
    create_table_for_syllabary(
        slide1,
        HIRAGANA_TABLE,
        "Hiragana (ひらがな) Overview",
        top=0.5, left=0.5,
        col_width=1.1, row_height=1.0,
        font_main=32, font_sub=14
    )
    create_centered_textbox(
        slide1,
        f"{HIRAGANA_N}\nN | ன்",  
        left=11.0, top=5.5, width=1.2, height=1.0,
        font_size=24, bold=False
    )

    # 2) Basic Katakana Overview
    slide2 = prs.slides.add_slide(blank_layout)
    create_table_for_syllabary(
        slide2,
        KATAKANA_TABLE,
        "Katakana (カタカナ) Overview",
        top=0.5, left=0.5,
        col_width=1.1, row_height=1.0,
        font_main=32, font_sub=14
    )
    create_centered_textbox(
        slide2,
        f"{KATAKANA_N}\nN | ன்",
        left=11.0, top=5.5, width=1.2, height=1.0,
        font_size=24, bold=False
    )

    # 3) Hiragana Dakuten + Handakuten
    slide3 = prs.slides.add_slide(blank_layout)
    create_table_for_syllabary(
        slide3,
        HIRAGANA_DAKUTEN_TABLE,
        "Hiragana Dakuten/Handakuten",
        top=0.5, left=0.7,
        col_width=1.1, row_height=1.0,
        font_main=32, font_sub=14
    )

    # 4) Katakana Dakuten + Handakuten
    slide4 = prs.slides.add_slide(blank_layout)
    create_table_for_syllabary(
        slide4,
        KATAKANA_DAKUTEN_TABLE,
        "Katakana Dakuten/Handakuten",
        top=0.5, left=0.7,
        col_width=1.1, row_height=1.0,
        font_main=32, font_sub=14
    )

    # 5) Focus Slides for Basic Gojuon
    for (series_name, hira_list, kata_list) in GOJUON_SERIES:
        # Series Overview Slide
        overview_slide = prs.slides.add_slide(blank_layout)
        create_centered_textbox(
            overview_slide,
            series_name,
            left=2.0, top=1.0,
            width=8.0, height=1.0,
            font_size=60, bold=True
        )
        # Show the full row
        hira_str = " ".join(hira_list)
        kata_str = " ".join(kata_list)
        create_centered_textbox(
            overview_slide,
            f"Hiragana: {hira_str}\nKatakana: {kata_str}",
            left=2.0, top=3.0,
            width=8.0, height=2.0,
            font_size=36, bold=False
        )

        # Individual Focus Slides
        for i in range(len(hira_list)):
            h = hira_list[i]
            k = kata_list[i]
            (romaji_h, tamil_h) = ROMAJI_TAMIL_MAP.get(h, ("", ""))
            (romaji_k, tamil_k) = ROMAJI_TAMIL_MAP.get(k, ("", ""))

            focus_slide = prs.slides.add_slide(blank_layout)
            # Large side-by-side
            create_centered_textbox(
                focus_slide,
                f"{h}    {k}",
                left=3.0, top=2.0,
                width=7.0, height=1.5,
                font_size=120, bold=True
            )
            # Smaller Romaji | Tamil (from Hiragana mapping)
            create_centered_textbox(
                focus_slide,
                f"{romaji_h} | {tamil_h}",
                left=3.0, top=4.0,
                width=7.0, height=1.0,
                font_size=50, bold=False
            )

    # 6) Focus Slides for Dakuten Series
    for (series_name, hira_list, kata_list) in DAKUTEN_SERIES:
        # Series Overview Slide
        overview_slide = prs.slides.add_slide(blank_layout)
        create_centered_textbox(
            overview_slide,
            series_name,
            left=2.0, top=1.0,
            width=8.0, height=1.0,
            font_size=60, bold=True
        )
        # Show the full row
        hira_str = " ".join(hira_list)
        kata_str = " ".join(kata_list)
        create_centered_textbox(
            overview_slide,
            f"Hiragana: {hira_str}\nKatakana: {kata_str}",
            left=2.0, top=3.0,
            width=8.0, height=2.0,
            font_size=36, bold=False
        )

        # Individual Focus Slides
        for i in range(len(hira_list)):
            h = hira_list[i]
            k = kata_list[i]
            (romaji_h, tamil_h) = ROMAJI_TAMIL_MAP.get(h, ("", ""))
            (romaji_k, tamil_k) = ROMAJI_TAMIL_MAP.get(k, ("", ""))

            focus_slide = prs.slides.add_slide(blank_layout)
            # Large side-by-side
            create_centered_textbox(
                focus_slide,
                f"{h}    {k}",
                left=3.0, top=2.0,
                width=7.0, height=1.5,
                font_size=120, bold=True
            )
            # Smaller Romaji | Tamil
            create_centered_textbox(
                focus_slide,
                f"{romaji_h} | {tamil_h}",
                left=3.0, top=4.0,
                width=7.0, height=1.0,
                font_size=50, bold=False
            )

    # 7) Save
    output_name = "Japanese_Guide_for_Tamil_Speakers_v4.pptx"
    prs.save(output_name)
    print(f"Presentation saved as '{output_name}'.")

if __name__ == "__main__":
    main()
