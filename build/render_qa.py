from pathlib import Path
import fitz
from PIL import Image, ImageOps, ImageDraw
ROOT=Path(__file__).resolve().parents[1]
p=ROOT/'qa/pdf/PAP_RAD_HMV_Pathways_v2_all_slides.pdf'
doc=fitz.open(p)
thumbs=[]
for i,page in enumerate(doc):
    pix=page.get_pixmap(matrix=fitz.Matrix(1.4,1.4),alpha=False)
    out=ROOT/'qa'/f'slide_{i+1:02}.png';pix.save(out)
    im=Image.open(out).convert('RGB');im.thumbnail((640,360))
    tile=Image.new('RGB',(660,390),'#e3e7e7');tile.paste(im,((660-im.width)//2,10))
    ImageDraw.Draw(tile).text((14,369),f'{i+1:02}',fill='black');thumbs.append(tile)
for start in range(0,len(thumbs),12):
    chunk=thumbs[start:start+12];rows=(len(chunk)+2)//3
    canvas=Image.new('RGB',(1980,390*rows),'white')
    for j,im in enumerate(chunk):canvas.paste(im,((j%3)*660,(j//3)*390))
    canvas.save(ROOT/'qa'/f'montage_{start+1:02}_{start+len(chunk):02}.png')
print(len(doc),'pages rendered')
