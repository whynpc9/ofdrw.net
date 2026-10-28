package main
import (
 "context"
 "encoding/json"
 "os"
 "path/filepath"
 "fmt"
 "github.com/xiaoqidun/ofdgo"
)
func check(e error){if e!=nil{panic(e)}}
func main(){
 if len(os.Args)!=4 {panic("font input-dir output-dir")}
 font,err:=os.ReadFile(os.Args[1]);check(err); out:=os.Args[3];check(os.MkdirAll(out,0755))
 type sample struct {name,first,second,dx1,dx2 string;maps []ofdgo.CGTransform}
 cases:=[]sample{
 {"second-only","AB","CD","8","8",[]ofdgo.CGTransform{{CodePosition:2,CodeCount:1,GlyphCount:1,Glyphs:"40"}}},
 {"both-local-zero","AB","CD","8","8",[]ofdgo.CGTransform{{CodePosition:0,CodeCount:1,GlyphCount:1,Glyphs:"40"},{CodePosition:2,CodeCount:1,GlyphCount:1,Glyphs:"41"}}},
 {"unequal-both","fl","A","","9",[]ofdgo.CGTransform{{CodePosition:0,CodeCount:2,GlyphCount:1,Glyphs:"1653"},{CodePosition:2,CodeCount:1,GlyphCount:2,Glyphs:"36 37"}}},
 {"mixed-second","AB","CflD","8","9 10",[]ofdgo.CGTransform{{CodePosition:3,CodeCount:2,GlyphCount:1,Glyphs:"1653"}}},
 }
 for _,c:=range cases{
  editor:=ofdgo.NewEditor();page,err:=editor.AddPage(120,65);check(err);id,err:=editor.AddFont(ofdgo.FontFile{Data:font},0);check(err)
  text:=ofdgo.TextObject{Boundary:"10 10 90 40",Font:id,Size:6,CGTransform:c.maps,
   TextCode:[]ofdgo.TextCode{{X:"3",Y:"6",DeltaX:c.dx1,Value:c.first},{X:"3",Y:"22",DeltaX:c.dx2,Value:c.second}}}
  _,err=editor.AddObject(page,ofdgo.GraphicObject{Type:"TextObject",TextObject:text});check(err)
  target,err:=ofdgo.NewFileOutput(filepath.Join(out,"go-"+c.name+".ofd"));check(err)
  _,err=editor.SaveAs(context.Background(),target,ofdgo.WriteOptions{});check(err)
 }
 files,err:=filepath.Glob(filepath.Join(os.Args[2],"grouped-*.ofd"));check(err)
 produced,err:=filepath.Glob(filepath.Join(out,"go-*.ofd"));check(err);files=append(files,produced...)
 for _,path:=range files{
  input,err:=ofdgo.OpenFile(context.Background(),path,ofdgo.InputOptions{});check(err)
  reader,err:=ofdgo.NewReader(input);check(err);page,err:=reader.PageContentByIndex(0);check(err)
  data,err:=json.MarshalIndent(page,"","  ");check(err)
  stem:=filepath.Base(path[:len(path)-4]);check(os.WriteFile(filepath.Join(out,stem+"-decoded.json"),data,0644))
  renderer:=ofdgo.NewRenderer(reader)
  pdf,err:=os.Create(filepath.Join(out,stem+"-independent.pdf"));check(err);check(renderer.RenderToPDF(page,pdf));check(pdf.Close())
  svg,err:=os.Create(filepath.Join(out,stem+"-independent.svg"));check(err);check(renderer.RenderToSVG(page,svg));check(svg.Close())
  reader.Close();input.Close();fmt.Println(stem,"writer/read/PDF/SVG complete")
 }
}
