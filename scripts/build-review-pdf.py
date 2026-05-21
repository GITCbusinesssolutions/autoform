from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "electrical-service-closeout-sample.pdf"


def section(title, rows):
    styles = getSampleStyleSheet()
    heading = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=15,
        textColor=colors.HexColor("#1f2937"),
        spaceBefore=8,
        spaceAfter=6,
    )
    label_style = ParagraphStyle(
        "Label",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=10,
        textColor=colors.HexColor("#334155"),
    )
    value_style = ParagraphStyle(
        "Value",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=10,
        textColor=colors.HexColor("#111827"),
    )
    data = [
        [Paragraph(label, label_style), Paragraph(value.replace("\n", "<br/>"), value_style)]
        for label, value in rows
    ]
    table = Table(data, colWidths=[62 * mm, 110 * mm], hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#cbd5e1")),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f8fafc")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return [Paragraph(title, heading), table, Spacer(1, 5 * mm)]


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=A4,
        rightMargin=16 * mm,
        leftMargin=16 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title="Electrical Service Closeout Sample",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "Title",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#0f172a"),
        spaceAfter=4,
    )
    subtitle = ParagraphStyle(
        "Subtitle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=10,
    )
    story = [
        Paragraph("Electrical Service Closeout", title),
        Paragraph("Sample report preview using test values for review before .sm8f import.", subtitle),
    ]
    sections = [
        (
            "Job Details",
            [
                ("Date", "21/05/2026"),
                ("Job Address", "42 Sample Street, Brisbane QLD"),
                ("Technician", "Alex Technician"),
                ("Job Status", "Completed"),
            ],
        ),
        (
            "Site Arrival",
            [
                ("Before Photos", "[2 photos captured before works commenced]"),
            ],
        ),
        (
            "Service Details",
            [
                ("Service Type", "Test & Tag"),
                ("Specify Service Type", ""),
                ("Job Category", "Scheduled Maintenance"),
            ],
        ),
        (
            "Labour and Attendance",
            [
                ("Technician Names", "Alex Technician, Morgan Electrician"),
                ("No of Technicians on Site", "2"),
                ("Start Time", "07:30"),
            ],
        ),
        (
            "Estimated Time and Variation Control",
            [
                ("Estimated Time to Complete Hours", "4"),
                ("Office Notified if Exceeding Estimate", "Yes"),
                ("Variation Notification Acknowledgement", ""),
                ("Time Notified", "10:45"),
                ("Who was Notified", "Service Coordinator"),
            ],
        ),
        (
            "Works Completed",
            [
                (
                    "Works Completed",
                    "- Isolated power to DB-1\n- Tested and tagged portable equipment\n- Replaced damaged plug top\n- Restored supply and checked operation",
                ),
            ],
        ),
        (
            "Site Delays",
            [
                ("Delay Type", "Access Issue"),
                ("Time Lost Hours", "0.5"),
                ("Delay Explanation", "Waiting for site contact to unlock switch room."),
                ("Office Notified of Delay", "Yes"),
            ],
        ),
        (
            "System Status on Exit",
            [
                ("System Status on Exit", "Operational - Minor Issues"),
                ("Action Required / Follow-Up Needed", "- Replace cracked GPO cover in plant room"),
            ],
        ),
        (
            "Additional Works Required",
            [
                ("Minor Works Required", "Yes"),
                ("Minor Works Details", "- Replace cracked GPO cover"),
                ("Major Works Required", "No"),
                ("Urgent Repair Required", "No"),
                ("ETA for Quote to Client", "24/05/2026 - minor works quote"),
            ],
        ),
        (
            "Test and Tag Summary",
            [
                ("Test Tag Completed", "Yes"),
                ("Total Items Tested", "37"),
                ("Test Tag Report Uploaded", "Yes"),
            ],
        ),
        (
            "Job Completion",
            [
                ("Job Completion Status", "Completed"),
                ("Finish Time", "12:15"),
                ("After Photos", "[3 photos captured after works completed]"),
                (
                    "Site Close Out Confirmed",
                    "SWMS Actioned; JSA Completed; Check In Completed; Check Out Completed; Site Left Safe; Office Notified Job Complete",
                ),
            ],
        ),
    ]
    for title_text, rows in sections:
        story.extend(section(title_text, rows))
    doc.build(story)
    print(OUT)


if __name__ == "__main__":
    main()
