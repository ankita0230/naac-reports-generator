import os
import sys
import math
import time
import pandas as pd
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from PIL import Image


def load_csv(csv_path):
    csv_path = os.path.abspath(csv_path)

    if not os.path.exists(csv_path):
        # try to find similarly named file in the same directory (ignore extension/trailing spaces)
        directory = os.path.dirname(csv_path) or "."
        base = os.path.splitext(os.path.basename(csv_path))[0].strip().lower()
        found = None
        try:
            for f in os.listdir(directory):
                name = os.path.splitext(f)[0].strip().lower()
                if name == base:
                    found = os.path.join(directory, f)
                    break
            if not found:
                # try substring match
                for f in os.listdir(directory):
                    name = os.path.splitext(f)[0].strip().lower()
                    if base in name or name in base:
                        found = os.path.join(directory, f)
                        break
        except FileNotFoundError:
            found = None

        if found and os.path.exists(found):
            print(f"ℹ️ Using found file instead: {found}")
            csv_path = os.path.abspath(found)
        else:
            raise FileNotFoundError(f"CSV file not found: {csv_path}")

    # detect if file is actually an Excel file (sometimes saved with .csv extension)
    is_excel_like = False
    try:
        with open(csv_path, "rb") as fh:
            head = fh.read(8)
            if head.startswith(b"PK") or b"<xml" in head:
                is_excel_like = True
    except Exception:
        is_excel_like = False

    # try CSV first, but if file looks like Excel, read as Excel
    if is_excel_like:
        # read all sheets and extract non-empty cells into a single-column DataFrame
        try:
            sheets = pd.read_excel(csv_path, sheet_name=None, header=None)
            rows = []
            for sname, sdf in sheets.items():
                for _, r in sdf.iterrows():
                    vals = [str(x).strip() for x in r.values if pd.notna(x) and str(x).strip() != "nan"]
                    if vals:
                        rows.append({"text": " | ".join(vals)})
            df = pd.DataFrame(rows)
        except Exception:
            df = pd.read_excel(csv_path)
    else:
        try:
            df = pd.read_csv(csv_path, encoding="utf-8", low_memory=False)
            # if CSV read produced no meaningful data, try Excel
            if df.empty:
                try:
                    # try reading all sheets and extracting non-empty rows
                    sheets = pd.read_excel(csv_path, sheet_name=None, header=None)
                    rows = []
                    for sname, sdf in sheets.items():
                        for _, r in sdf.iterrows():
                            vals = [str(x).strip() for x in r.values if pd.notna(x) and str(x).strip() != "nan"]
                            if vals:
                                rows.append({"text": " | ".join(vals)})
                    df = pd.DataFrame(rows)
                except Exception:
                    pass
        except Exception:
            df = pd.read_excel(csv_path)

    if df.empty:
        raise ValueError("CSV file is empty")

    print("✅ CSV loaded successfully")
    return df


def csv_to_text(df):
    # Concatenate all non-empty cells into a single text blob (fallback)
    text_data = []
    for _, row in df.iterrows():
        vals = [str(x).strip() for x in row.values if pd.notna(x) and str(x).strip() != "nan"]
        if vals:
            text_data.append(" | ".join(vals))
    return "\n".join(text_data)


def split_text(text, chunk_size=800, chunk_overlap=100):
    # Simple fallback splitter that slices by characters
    if not text:
        return []
    step = chunk_size - chunk_overlap
    chunks = [text[i:i+chunk_size] for i in range(0, max(1, len(text)), step)]
    return chunks


def create_vectorstore(text_chunks):
    # Placeholder: removed heavy dependency on langchain/FAISS.
    # Keep this function to preserve pipeline compatibility.
    if not text_chunks:
        raise ValueError("No text chunks created")
    print(f"ℹ️ Prepared {len(text_chunks)} text chunks (no vector store created)")
    return None


def generate_pdf(output_path, summary_text, df=None, images_dir=None):
    # Use Platypus for better layout and tables
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            rightMargin=2*cm, leftMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    story = []

    title_style = styles["Title"]
    normal = styles["Normal"]
    heading = ParagraphStyle("Heading", parent=styles["Heading2"], alignment=1)

    story.append(Paragraph("NAAC Accreditation Report", title_style))
    story.append(Spacer(1, 12))

    story.append(Paragraph(summary_text.replace("\n", "<br/>"), normal))
    story.append(Spacer(1, 12))

    # Add a simple metadata line
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["BodyText"]))
    story.append(Spacer(1, 12))

    # Table of first 30 rows with up to 6 columns (adjustable)
    if df is not None and not df.empty:
        max_cols = 6
        cols = list(df.columns[:max_cols])
        data = [cols]
        for _, row in df.head(30).iterrows():
            row_vals = [str(row.get(c, "")) for c in cols]
            # wrap long text in Paragraph for cell wrapping
            row_vals = [Paragraph(v.replace('\n', '<br/>'), normal) for v in row_vals]
            data.append(row_vals)

        tbl = Table(data, repeatRows=1)
        tbl_style = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d3d3d3")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ])
        tbl.setStyle(tbl_style)
        story.append(Spacer(1, 8))
        story.append(Paragraph("Sample Records (first 30)", styles["Heading3"]))
        story.append(Spacer(1, 6))
        story.append(tbl)
        story.append(Spacer(1, 12))

    # Images section
    if images_dir:
        images_dir = os.path.abspath(images_dir)
        if os.path.isdir(images_dir):
            img_files = [os.path.join(images_dir, f) for f in os.listdir(images_dir)
                         if f.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp"))]
            if img_files:
                story.append(Spacer(1, 12))
                story.append(Paragraph("Attached Images", styles["Heading3"]))
                story.append(Spacer(1, 6))
                for img_path in img_files:
                    try:
                        # resize image preserving aspect ratio to fit page width
                        im = Image.open(img_path)
                        iw, ih = im.size
                        max_w = A4[0] - 4*cm
                        max_h = A4[1] - 6*cm
                        scale = min(max_w / iw, max_h / ih, 1)
                        draw_w = iw * scale
                        draw_h = ih * scale
                        rlimg = RLImage(img_path, width=draw_w, height=draw_h)
                        story.append(rlimg)
                        story.append(Spacer(1, 6))
                    except Exception as ex:
                        print(f"⚠️ Could not include image {img_path}: {ex}")

    def add_page_number(canvas_obj, doc_obj):
        canvas_obj.saveState()
        page_num_text = f"Page {doc_obj.page}"
        canvas_obj.setFont('Helvetica', 8)
        canvas_obj.drawRightString(A4[0] - 2*cm, 1*cm, page_num_text)
        canvas_obj.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    print(f"✅ Report generated: {output_path}")


def main():
    if len(sys.argv) < 2:
        print("❌ Usage: py app.py <csv_file_path> [images_folder]")
        return

    csv_file = sys.argv[1]
    images_folder = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(csv_file), "images")

    try:
        df = load_csv(csv_file)
        text = csv_to_text(df)
        chunks = split_text(text)
        create_vectorstore(chunks)

        # Build a basic summary
        summary_lines = []
        summary_lines.append(f"Total Records: {len(df)}")
        # try to compute average CGPA if present
        if "CGPA" in df.columns:
            def to_float(x):
                try:
                    return float(str(x).strip())
                except Exception:
                    return None
            cgpas = [to_float(x) for x in df["CGPA"] if pd.notna(x)]
            cgpas = [x for x in cgpas if x is not None]
            if cgpas:
                summary_lines.append(f"Average CGPA: {sum(cgpas)/len(cgpas):.2f}")
        # top grades
        if "Grade" in df.columns:
            top = df["Grade"].value_counts().head(5).to_dict()
            summary_lines.append("Top Grades: " + ", ".join(f"{k}({v})" for k, v in top.items()))

        summary = "\n".join(["This report is generated automatically from NAAC accreditation data."] + summary_lines)

        os.makedirs("Output", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        pdf_path = os.path.join("Output", f"NAAC_Report_{timestamp}.pdf")

        # choose images folder if it exists
        images_dir_to_use = images_folder if os.path.isdir(images_folder) else os.path.join("Output", "images")

        generate_pdf(pdf_path, summary, df=df, images_dir=images_dir_to_use)

    except Exception as e:
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    main()
