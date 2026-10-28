"""Compare actual candidate paths and independent PDF glyphs with FontTools originals.
Usage: python check-geometry.py font.ttf candidate-output independent-output receipt.json
Dependencies: fonttools, pymupdf, pypdf (record their versions in the receipt).
"""
import io
import json
import math
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
import fontTools
from fontTools.ttLib import TTFont
from fontTools.pens.basePen import BasePen
from fontTools.svgLib.path import parse_path
import pymupdf
from pypdf import PdfReader

font_path, candidate, independent, receipt = map(Path, sys.argv[1:])
font = TTFont(font_path)
glyphs = font.getGlyphSet()
scale = 6 / font['head'].unitsPerEm
cases = {
 'second-only': ([36,37,40,39], [(3,6),(11,6),(3,22),(11,22)], 'AB\nCD'),
 'both-local-zero': ([40,37,41,39], [(3,6),(11,6),(3,22),(11,22)], 'AB\nCD'),
 'unequal-both': ([1653,36,37], [(3,6),(3,22),(12,22)], 'fl\nA'),
 'mixed-second': ([36,37,38,1653,39], [(3,6),(11,6),(3,22),(12,22),(22,22)], 'AB\nCflD')
}
class Pen(BasePen):
 def __init__(self, gs):
  super().__init__(gs); self.segments=[]; self.start=None
 def _moveTo(self,p): self.start=p
 def _lineTo(self,p): self.segments.append(('l',self._getCurrentPoint(),p))
 def _curveToOne(self,a,b,c): self.segments.append(('c',self._getCurrentPoint(),a,b,c))
 def _qCurveToOne(self,b,c):
  a=self._getCurrentPoint()
  self._curveToOne(tuple(a[i]+(b[i]-a[i])*2/3 for i in range(2)),
                  tuple(c[i]+(b[i]-c[i])*2/3 for i in range(2)),c)
 def _closePath(self):
  if self._getCurrentPoint()!=self.start: self._lineTo(self.start)
 def _endPath(self): pass

def shape(f,gid):
 gs=f.getGlyphSet(); p=Pen(gs);gs[f.getGlyphOrder()[gid]].draw(p);return p.segments

def placed(segments,origin):
 x,y=origin
 return [(s[0],*((x+p[0]*scale,y-p[1]*scale) for p in s[1:])) for s in segments]

def compare(actual,expected,tolerance):
 remaining=list(actual);maximum=0
 assert len(actual)==len(expected),(len(actual),len(expected))
 for wanted in expected:
  candidates=[]
  for i,s in enumerate(remaining):
   if s[0]!=wanted[0] or len(s)!=len(wanted):continue
   variants=[s]
   # MuPDF may represent a font's rectangular contour as a PDF rectangle.
   if s[0]=='l':variants.append((s[0],s[2],s[1]))
   err=min(max(abs(a-b) for p,q in zip(v[1:],wanted[1:]) for a,b in zip(p,q)) for v in variants)
   candidates.append((err,i))
  assert candidates, wanted
  err,i=min(candidates);assert err<=tolerance,(err,wanted,remaining[i])
  maximum=max(maximum,err);remaining.pop(i)
 return maximum

def pdf_segments(drawing):
 result=[]
 for s in drawing['items']:
  if s[0] in ('l','c'):result.append((s[0],*(tuple(v*25.4/72 for v in p) for p in s[1:])))
  elif s[0]=='re':
   r=s[1];points=[r.top_left,r.top_right,r.bottom_right,r.bottom_left]
   for a,b in zip(points,points[1:]+points[:1]):
    result.append(('l',tuple(v*25.4/72 for v in a),tuple(v*25.4/72 for v in b)))
  else:raise AssertionError(s)
 return result

rows=[]
for name,(gids,origins,logical) in cases.items():
 for prefix in ('grouped-','go-'):
  stem=prefix+name
  pdf_path=candidate/(stem+'.pdf')
  if not pdf_path.exists():
   assert stem=='go-unequal-both',stem
   continue
  tree=ET.parse(candidate/(stem+'.svg'))
  paths=[n for n in tree.iter() if 'data-original-gid' in n.attrib]
  assert [int(n.attrib['data-original-gid']) for n in paths]==gids
  doc=pymupdf.open(pdf_path);drawings=doc[0].get_drawings();assert len(drawings)==len(gids)
  svg_max=pdf_max=0
  for gid,origin,path,drawing in zip(gids,origins,paths,drawings):
   raw=shape(font,gid);p=Pen(None);parse_path(path.attrib['d'],p)
   svg_max=max(svg_max,compare(p.segments,placed(raw,origin),0.001))
   absolute=(origin[0]+10,origin[1]+10)
   pdf_max=max(pdf_max,compare(pdf_segments(drawing),placed(raw,absolute),0.001))
  other=independent/(stem+'-independent.pdf');go_doc=pymupdf.open(other)
  traces=[c for span in go_doc[0].get_texttrace() if span['opacity']>0 and span['type']==0 for c in span['chars']]
  assert len(traces)==len(gids),(stem,len(traces))
  fonts=go_doc[0].get_fonts(full=True);assert len(fonts)==1
  subset=TTFont(io.BytesIO(go_doc.extract_font(fonts[0][0])[3]))
  max_origin=0
  for char,gid,origin in zip(traces,gids,origins):
   compare(shape(subset,char[1]),shape(font,gid),1e-9)
   actual=tuple(v*25.4/72 for v in char[2]);expected=(origin[0]+10,origin[1]+10)
   error=max(abs(a-b) for a,b in zip(actual,expected));assert error<0.001,(stem,error)
   max_origin=max(max_origin,error)
  reader=PdfReader(other)
  actual_text=[str(args[1]['/ActualText']) for args,op in reader.pages[0].get_contents().operations
               if op==b'BDC' and len(args)>1 and '/ActualText' in args[1]]
  assert actual_text==[logical],(stem,actual_text,logical)
  rows.append(dict(stem=stem,originalGids=gids,originsObjectMm=origins,
   candidateSvgMaxCoordinateErrorMm=svg_max,candidatePdfMaxCoordinateErrorMm=pdf_max,
   independentPdfSubsetOutlinesEqualOriginal=True,independentPdfMaxOriginErrorMm=max_origin,
   independentActualText=actual_text,independentPypdfExtractedText=reader.pages[0].extract_text()))
receipt.write_text(json.dumps(dict(fontTools=fontTools.__version__,pymupdf=pymupdf.VersionBind,rows=rows),indent=2)+'\n')
print('Geometry checks passed:',len(rows),'input/output comparisons,',sum(len(r['originalGids']) for r in rows),'glyphs')
