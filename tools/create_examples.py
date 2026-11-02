"""Regenerate clearly labeled demo assets; no participant data."""
from pathlib import Path
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

ROOT=Path(__file__).resolve().parents[1]


def make_template(path,role):
    p=Presentation();p.slide_width=Inches(11.69);p.slide_height=Inches(8.27)
    s=p.slides.add_slide(p.slide_layouts[6])
    s.background.fill.solid();s.background.fill.fore_color.rgb=RGBColor.from_string('F5F3EC')
    border=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(.35),Inches(.35),Inches(10.99),Inches(7.57))
    border.fill.background();border.line.color.rgb=RGBColor.from_string('176B70');border.line.width=Pt(2)
    def label(y,height,content,size,color='244753'):
        box=s.shapes.add_textbox(Inches(.8),Inches(y),Inches(10.09),Inches(height))
        for i,line in enumerate(content.split('\n')):
            para=box.text_frame.paragraphs[0] if i==0 else box.text_frame.add_paragraph()
            para.alignment=PP_ALIGN.CENTER;r=para.add_run();r.text=line;r.font.name='Arial';r.font.size=Pt(size);r.font.color.rgb=RGBColor.from_string(color)
    label(.7,.5,'COMMUNITY FOODBANK',18,'176B70')
    label(1.55,.9,'CERTIFICATE OF APPRECIATION',29)
    label(2.65,.6,'Presented with gratitude to',18)
    label(3.3,.8,'{{NAME}}',36,'176B70')
    label(4.45,1.1,f'For your generous contribution as a {role}\nand your support for our community.',19)
    label(6.0,.5,'Date: {{DATE}}',16)
    label(7.15,.4,'DEMO TEMPLATE — replace with your organization’s approved design',11)
    p.save(path)


def main():
    target=ROOT/'examples';target.mkdir(exist_ok=True)
    for k in ('volunteer','donor'):make_template(target/f'{k}_sample.pptx',k)
    w=Workbook();s=w.active;s.title='Records';s.append(['Participant Name','Type','Phone','Email','Date'])
    s.append(['Example Volunteer','volunteer','','','2026-11-02'])
    s.append(['Example Donor','donor','','','2026-11-02'])
    s.append(['Example Family','beneficiary','','','2026-11-02'])
    s.freeze_panes='A2';s.column_dimensions['A'].width=27
    w.save(target/'sample_records.xlsx');w.close()


if __name__=='__main__':main()
