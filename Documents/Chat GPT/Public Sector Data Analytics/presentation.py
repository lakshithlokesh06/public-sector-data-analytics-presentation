from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from math import sin, cos, pi

OUT = 'Public-Sector-Data-Analytics-Slide-1.pptx'
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])

# Restrained night palette
BG = RGBColor(8, 16, 27)
PANEL = RGBColor(12, 24, 38)
GRID = RGBColor(23, 43, 58)
MUTED = RGBColor(119, 145, 159)
WHITE = RGBColor(236, 242, 243)
TEAL = RGBColor(77, 210, 191)
BLUE = RGBColor(67, 130, 156)

slide.background.fill.solid(); slide.background.fill.fore_color.rgb = BG

def rect(x,y,w,h,color,trans=0,line=None):
    s=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb=color; s.fill.transparency=trans
    s.line.fill.background() if line is None else None
    if line:
        s.line.color.rgb=line; s.line.width=Pt(.7)
    return s

def line(x1,y1,x2,y2,color,width=1,trans=0,dash=None):
    s=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    s.line.color.rgb=color; s.line.width=Pt(width); s.line.transparency=trans
    if dash: s.line.dash_style=dash
    return s

def circle(cx,cy,d,color,trans=0):
    s=slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx-d/2), Inches(cy-d/2), Inches(d), Inches(d))
    s.fill.solid(); s.fill.fore_color.rgb=color; s.fill.transparency=trans; s.line.fill.background()
    return s

def text(txt,x,y,w,h,size,color,font='Aptos',bold=False,spacing=None):
    box=slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf=box.text_frame; tf.clear(); tf.word_wrap=False
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tf.paragraphs[0]; p.alignment=PP_ALIGN.LEFT
    r=p.add_run(); r.text=txt; r.font.name=font; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color
    if spacing is not None: r.font.character_spacing=spacing
    return box

# Delicate technical grid, intentionally subordinate to title
for x in [0.72+i*.42 for i in range(30)]: line(x,0.72,x,6.82,GRID,.35,36)
for y in [0.72+i*.42 for i in range(15)]: line(.72,y,12.62,y,GRID,.35,36)
# Framing marks
line(.72,.72,1.34,.72,TEAL,1.2,5); line(.72,.72,.72,1.34,TEAL,1.2,5)
line(12.0,6.78,12.62,6.78,TEAL,1.2,5); line(12.62,6.16,12.62,6.78,TEAL,1.2,5)

# Right-side abstract city field
rect(7.45,1.05,5.17,5.55,PANEL,8)
# city skyline made from quiet architectural blocks
buildings=[(7.75,4.10,.57,1.48),(8.42,3.44,.78,2.14),(9.30,3.92,.61,1.66),(10.00,2.85,.83,2.73),(10.92,3.53,.67,2.05),(11.68,2.54,.70,3.04)]
for x,y,w,h in buildings:
    rect(x,y,w,h,RGBColor(15,35,49),4,GRID)
    # roof caps and window lights
    line(x+.05,y,x+w-.05,y,BLUE,.9,15)
    cols=max(2,int(w/.19)); rows=max(3,int(h/.33))
    for c in range(cols):
        for r in range(rows):
            wx=x+.11+c*.19; wy=y+.16+r*.31
            if wy < y+h-.08:
                rect(wx,wy,.035,.055,TEAL if (c+r+int(x*10))%9==0 else BLUE,45 if (c+r)%5 else 68)
# Ground datum
line(7.68,5.62,12.43,5.62,BLUE,.7,38)
# Network nodes, arcs, and paths layered across skyline
nodes=[(7.98,4.05),(8.70,3.28),(9.45,4.02),(10.42,2.65),(11.24,3.37),(12.03,2.34),(8.45,5.12),(9.58,4.92),(10.86,4.58),(11.85,5.12)]
edges=[(0,1),(1,2),(1,3),(2,3),(2,6),(3,4),(3,7),(4,5),(4,8),(5,9),(6,7),(7,8),(8,9)]
for a,b in edges:
    line(*nodes[a],*nodes[b],BLUE,.85,28)
for i,(x,y) in enumerate(nodes):
    circle(x,y,.085,TEAL if i in (1,3,7) else BLUE,5)
    circle(x,y,.18,TEAL if i in (1,3,7) else BLUE,84)
# A fine horizontal scanline through the field
line(7.65,4.57,12.43,4.57,TEAL,.65,52,MSO_LINE_DASH_STYLE.DASH)

# Left title lockup
text('AI  /  PUBLIC SYSTEMS',1.02,1.36,4.8,.25,10,MUTED,bold=True,spacing=1.4)
rect(1.02,1.87,.48,.035,TEAL)
text('Public Sector',1.00,2.14,6.45,.72,35,WHITE,'Aptos Display',True)
text('Data Analytics',1.00,2.89,6.5,.79,38,WHITE,'Aptos Display',True)
text('AI Ethics Case Study 6',1.04,3.98,4.3,.32,14,TEAL,bold=False,spacing=.35)
# minimal thematic line, supported by case study framing
text('Crime reduction  /  data  /  public trust',1.04,5.02,5.9,.28,11,MUTED,spacing=.35)
# footer ID
line(1.02,6.37,2.02,6.37,GRID,.8)
text('CASE STUDY 06',1.02,6.50,2.3,.22,9,MUTED,bold=True,spacing=1.1)
text('NEW LEVIATHAN',10.0,6.77,2.55,.17,8,MUTED,bold=True,spacing=1.15)

prs.save(OUT)
print(OUT)
