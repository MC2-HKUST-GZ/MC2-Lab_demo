from __future__ import annotations

import html
import re
from pathlib import Path

from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    HRFlowable,
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
HTML_PATH = REPO_ROOT / "public/legacy/news-ai-as-catalyst-workshop.html"
OUTPUT_PATH = REPO_ROOT / "output/doc/workshop-participants-en.pdf"


def clean_html_text(raw: str) -> str:
    text = re.sub(r"<br\s*/?>", " ", raw)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def parse_participants(html_text: str) -> list[dict[str, str | Path]]:
    articles = re.findall(r'<article class="speaker-card-news".*?</article>', html_text, flags=re.S)
    participants: list[dict[str, str | Path]] = []

    for block in articles:
        img_match = re.search(r'<img\s+src="([^"]+)"', block)
        name_match = re.search(r"<h5>(.*?)</h5>", block, flags=re.S)
        role_match = re.search(r'<p class="speaker-card-news__role">(.*?)</p>', block, flags=re.S)
        org_match = re.search(r'<p class="speaker-card-news__org">(.*?)</p>', block, flags=re.S)
        bio_match = re.search(r'<p class="speaker-card-news__bio">(.*?)</p>', block, flags=re.S)

        if not all([img_match, name_match, role_match, org_match, bio_match]):
            continue

        image_src = img_match.group(1).split("?", 1)[0]
        image_path = REPO_ROOT / "public" / image_src.lstrip("/")

        participants.append(
            {
                "name": clean_html_text(name_match.group(1)),
                "role": clean_html_text(role_match.group(1)),
                "org": clean_html_text(org_match.group(1)),
                "bio": clean_html_text(bio_match.group(1)),
                "image_path": image_path,
            }
        )

    return participants


def make_image_flowable(image_path: Path, max_width: float, max_height: float) -> Image | Spacer:
    if not image_path.exists():
        return Spacer(max_width, max_height)

    with PILImage.open(image_path) as image:
        width, height = image.size

    if width <= 0 or height <= 0:
        return Spacer(max_width, max_height)

    ratio = min(max_width / width, max_height / height)
    draw_width = width * ratio
    draw_height = height * ratio
    return Image(str(image_path), width=draw_width, height=draw_height)


def build_pdf(participants: list[dict[str, str | Path]]) -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    page_width, page_height = A4
    doc = SimpleDocTemplate(
        str(OUTPUT_PATH),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=16 * mm,
    )

    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="CoverTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=30,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#0B2545"),
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CoverBody",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=11,
            leading=16,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#5B6472"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="ProfileName",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=20,
            leading=24,
            textColor=colors.HexColor("#102A43"),
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ProfileMeta",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10.5,
            leading=14,
            textColor=colors.HexColor("#486581"),
            spaceAfter=3,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ProfileBio",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=11,
            leading=17,
            textColor=colors.HexColor("#1F2933"),
            alignment=TA_LEFT,
        )
    )
    styles.add(
        ParagraphStyle(
            name="PageFoot",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=12,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#7B8794"),
        )
    )

    story = [
        Spacer(1, 42 * mm),
        Paragraph("AI as Catalyst Workshop Participants", styles["CoverTitle"]),
        Paragraph("English bios in current website order", styles["CoverBody"]),
        Spacer(1, 16 * mm),
        HRFlowable(width="35%", thickness=1.2, color=colors.HexColor("#D9E2EC"), lineCap="round"),
        Spacer(1, 12 * mm),
        Paragraph(
            f"This PDF includes {len(participants)} speakers and organizers, with one profile per page for easier reading and sharing.",
            styles["CoverBody"],
        ),
        PageBreak(),
    ]

    image_max_width = 56 * mm
    image_max_height = 72 * mm
    text_width = page_width - doc.leftMargin - doc.rightMargin - image_max_width - 10 * mm

    for index, participant in enumerate(participants, start=1):
        image = make_image_flowable(participant["image_path"], image_max_width, image_max_height)
        profile_text = [
            Paragraph(f"{index}. {html.escape(participant['name'])}", styles["ProfileName"]),
            Paragraph(html.escape(participant["role"]), styles["ProfileMeta"]),
            Paragraph(html.escape(participant["org"]), styles["ProfileMeta"]),
        ]

        header = Table(
            [[image, profile_text]],
            colWidths=[image_max_width, text_width],
            hAlign="LEFT",
        )
        header.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ]
            )
        )

        story.extend(
            [
                header,
                Spacer(1, 8 * mm),
                HRFlowable(width="100%", thickness=0.8, color=colors.HexColor("#D9E2EC"), lineCap="round"),
                Spacer(1, 8 * mm),
                Paragraph(html.escape(participant["bio"]), styles["ProfileBio"]),
                Spacer(1, 8 * mm),
                Paragraph(f"Participant {index} of {len(participants)}", styles["PageFoot"]),
            ]
        )

        if index != len(participants):
            story.append(PageBreak())

    def add_page_number(canvas, document) -> None:
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.HexColor("#7B8794"))
        canvas.drawRightString(page_width - 18 * mm, 10 * mm, f"Page {document.page}")

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)


def main() -> None:
    participants = parse_participants(HTML_PATH.read_text(encoding="utf-8"))
    build_pdf(participants)
    print(f"participants={len(participants)}")
    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
