"""Compact publication layout of the conceptual relationships explored in FigJam."""
from pathlib import Path
from html import escape
import fitz

root=Path(__file__).resolve().parents[1]/'figures'
parts=['<svg xmlns="http://www.w3.org/2000/svg" width="1100" height="940" viewBox="0 0 1100 940"><rect width="1100" height="940" fill="white"/><defs><marker id="arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L8,3 L0,6" fill="none" stroke="#596575" stroke-width="1.5"/></marker></defs>']
def box(x,y,w,h,lines,fill='#FFFFFF'):
    parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="8" fill="{fill}" stroke="#596575" stroke-width="2"/>')
    cy=y+h/2-(len(lines)-1)*17
    for j,t in enumerate(lines):parts.append(f'<text x="{x+w/2}" y="{cy+j*34}" dominant-baseline="middle" text-anchor="middle" font-family="Arial" font-size="28" fill="#172332">{escape(t)}</text>')
def edge(points):
    parts.append('<polyline points="'+points+'" fill="none" stroke="#596575" stroke-width="2.5"/>')
    pts=[tuple(map(float,p.split(','))) for p in points.split()]
    (x0,y0),(x,y)=pts[-2:];dx=x-x0;dy=y-y0;length=(dx*dx+dy*dy)**.5;ux=dx/length;uy=dy/length
    coords=[(x,y),(x-12*ux+5*uy,y-12*uy-5*ux),(x-12*ux-5*uy,y-12*uy+5*ux)]
    parts.append('<polygon points="'+' '.join(f'{a},{b}' for a,b in coords)+'" fill="#596575"/>')
# Business services feed both the impact and supporting-resource representations.
box(360,15,380,80,['Business services'])
box(25,150,420,95,['Time dependent','impact assessments'])
box(520,150,550,95,['Approved criteria','and decision rules'],'#E7EFF8')
box(25,310,500,110,['Derived service criticality','MTPD and RTO'])
box(625,310,445,110,['Supporting dependency','relationships'])
box(625,505,445,110,['Documented criticality','inheritance rule'])
box(25,700,500,95,['Relative service priorities'])
box(625,700,445,95,['Dependency criticality'])
box(285,850,530,75,['Decision support dashboard'],'#E7F2ED')
edge('435,95 435,120 235,120 235,145')
edge('665,95 665,120 1090,120 1090,290 850,290 850,305')
edge('235,245 235,305')
edge('795,245 795,270 420,270 420,305')
edge('275,420 275,695')
edge('525,365 565,365 565,560 620,560')
edge('850,420 850,500')
edge('850,615 850,695')
edge('275,795 275,825 425,825 425,845')
edge('850,795 850,825 675,825 675,845')
parts.append('</svg>')
svg=root/'literature-concept-map.svg';svg.write_text(''.join(parts))
d=fitz.open(svg);d[0].get_pixmap(matrix=fitz.Matrix(2,2)).save(str(svg.with_suffix('.png')))
