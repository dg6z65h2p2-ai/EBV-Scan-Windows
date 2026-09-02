using System.Drawing.Printing;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;

namespace EBVScan.Windows;

record Limit(string Name, string Unit, double?[] Values, double? Minimum = null);
record MaterialTable(string Title, string[] Classes, List<Limit> Limits);
record Reading(string Name, string Unit, string[] Texts, double?[] Values, bool[] Ignored);

static class EBVData
{
    static double?[] N(int count) => Enumerable.Repeat<double?>(null, count).ToArray();
    static IEnumerable<Limit> Common(int count) => new[] {
        "Arsen|mg/kg","Blei|mg/kg","Cadmium|mg/kg","Chrom gesamt|mg/kg","Kupfer|mg/kg","Nickel|mg/kg","Quecksilber|mg/kg","Thallium|mg/kg","Zink|mg/kg",
        "Cyanide|mg/kg","EOX|mg/kg","MKW C10-C40|mg/kg","MKW C10-C22|mg/kg","Σ PCB 7|mg/kg","Naphthalin|mg/kg","Acenaphthylen|mg/kg","Acenaphthen|mg/kg",
        "Fluoren|mg/kg","Phenanthren|mg/kg","Anthracen|mg/kg","Fluoranthen|mg/kg","Pyren|mg/kg","Benzo(a)anthracen|mg/kg","Chrysen|mg/kg",
        "Benzo(b)fluoranthen|mg/kg","Benzo(k)fluoranthen|mg/kg","Benzo(a)pyren|mg/kg","Dibenz(a,h)anthracen|mg/kg","Benzo(g,h,i)perylen|mg/kg","Indeno(1,2,3-c,d)pyren|mg/kg",
        "Messtemperatur pH-Wert|°C","Cyanide|mg/l","Antimon|µg/l","Arsen|µg/l","Blei|µg/l","Cadmium|µg/l","Chrom gesamt|µg/l","Kupfer|µg/l",
        "Molybdän|µg/l","Nickel|µg/l","Quecksilber|µg/l","Thallium|µg/l","Vanadium|µg/l","Zink|µg/l","Phenole|µg/l"
    }.Select(x => { var p=x.Split('|'); return new Limit(p[0],p[1],N(count)); });

    public static readonly Dictionary<string, MaterialTable> Tables = new()
    {
        ["Bauschutt"] = Make("Bauschutt / Recycling-Baustoff", new[]{"RC-1","RC-2","RC-3"}, new(){
            new("pH-Wert","",new double?[]{13,13,13},6), new("Elektrische Leitfähigkeit","µS/cm",new double?[]{2500,3200,10000}),
            new("Sulfat","mg/l",new double?[]{600,1000,3500}), new("PAK 15","µg/l",new double?[]{4,8,25}), new("PAK 16","mg/kg",new double?[]{10,15,20}),
            new("Chrom gesamt","µg/l",new double?[]{150,440,900}), new("Kupfer","µg/l",new double?[]{110,250,500}), new("Vanadium","µg/l",new double?[]{120,700,1350})
        }, 3),
        ["Gleisschotter"] = Make("Gleisschotter", new[]{"GS-0","GS-1","GS-2","GS-3"}, new(){
            new("pH-Wert","",new double?[]{10,10,10,12},5), new("Elektrische Leitfähigkeit","µS/cm",new double?[]{500,500,500,1000}),
            new("Atrazin","µg/l",new double?[]{.2,.7,3.5,14}), new("Bromacil","µg/l",new double?[]{.2,.4,1.2,5.3}), new("Diuron","µg/l",new double?[]{.1,.2,.8,4.6}),
            new("Glyphosat","µg/l",new double?[]{.2,1.7,17,27}), new("AMPA","µg/l",new double?[]{2.5,4.5,17,50}), new("Simazin","µg/l",new double?[]{.2,1.5,12,27}),
            new("Dimefuron","µg/l",new double?[]{.2,2.1,17,27}), new("Flazasulfuron","µg/l",new double?[]{.2,2.1,17,27}), new("Flumioxazin","µg/l",new double?[]{.2,2.1,17,27}),
            new("Ethidimuron","µg/l",new double?[]{.2,2.1,17,27}), new("Thiazafluron","µg/l",new double?[]{.2,2.1,17,27}), new("MKW","µg/l",new double?[]{150,160,310,500}), new("PAK 15","µg/l",new double?[]{.3,2.3,42,50})
        }, 4),
        ["Boden"] = new("Bodenmaterial",new[]{"BM-F0*","BM-F1","BM-F2","BM-F3"}, new List<Limit>{
            new("pH-Wert","",new double?[]{9.5,9.5,9.5,12},5.5), new("Elektrische Leitfähigkeit","µS/cm",new double?[]{350,500,500,2000}), new("Sulfat","mg/l",new double?[]{250,450,450,1000}),
            new("Arsen","mg/kg",new double?[]{40,40,40,150}),new("Arsen","µg/l",new double?[]{12,20,85,100}),new("Blei","mg/kg",new double?[]{140,140,140,700}),new("Blei","µg/l",new double?[]{35,90,250,470}),
            new("Cadmium","mg/kg",new double?[]{2,2,2,10}),new("Cadmium","µg/l",new double?[]{3,3,10,15}),new("Chrom gesamt","mg/kg",new double?[]{120,120,120,600}),new("Chrom gesamt","µg/l",new double?[]{15,150,290,530}),
            new("Kupfer","mg/kg",new double?[]{80,80,80,320}),new("Kupfer","µg/l",new double?[]{30,110,170,320}),new("Nickel","mg/kg",new double?[]{100,100,100,350}),new("Nickel","µg/l",new double?[]{30,30,150,280}),
            new("Quecksilber","mg/kg",new double?[]{.6,.6,.6,5}),new("Thallium","mg/kg",new double?[]{2,2,2,7}),new("Zink","mg/kg",new double?[]{300,300,300,1200}),new("Zink","µg/l",new double?[]{150,160,840,1600}),
            new("TOC","M%",new double?[]{5,5,5,5}),new("MKW C10-C22","mg/kg",new double?[]{300,300,300,1000}),new("MKW C10-C40","mg/kg",new double?[]{600,600,600,2000}),new("PAK 15","µg/l",new double?[]{.3,1.5,3.8,20}),new("PAK 16","mg/kg",new double?[]{6,6,9,30}),
            new("Antimon","µg/l",new double?[]{7.5,7.5,7.5,15}),new("Molybdän","µg/l",new double?[]{55,55,55,110}),new("Vanadium","µg/l",new double?[]{30,55,450,840}),new("EOX","mg/kg",new double?[]{3,3,3,10}),new("MKW","µg/l",new double?[]{150,160,160,310}),new("Cyanide","mg/kg",new double?[]{3,3,3,10}),new("Phenole","µg/l",new double?[]{12,60,60,2000})
        }.Concat(Common(4)).GroupBy(x=>x.Name+"|"+x.Unit).Select(x=>x.First()).ToList())
    };
    static MaterialTable Make(string title,string[] classes,List<Limit> rated,int n) => new(title,classes,rated.Concat(Common(n)).GroupBy(x=>x.Name+"|"+x.Unit).Select(x=>x.First()).ToList());
}

static class Analyzer
{
    static readonly Dictionary<string,string[]> Aliases=new(){
        ["pH-Wert"]=new[]{"ph-wert","ph "},["Messtemperatur pH-Wert"]=new[]{"messtemperatur ph-wert","messtemperatur ph wert"},
        ["Elektrische Leitfähigkeit"]=new[]{"el. leitfähigkeit","el leitfähigkeit","leitfähigkeit","leitfaehigkeit"},["Chrom gesamt"]=new[]{"chrom, ges","chrom ges","chrom gesamt","chrom (ges.)"},
        ["PAK 15"]=new[]{"σ pak15","∑ pak15","pak 15","pak15"},["PAK 16"]=new[]{"σ pak (epa)","∑ pak (epa)","pak (epa)","pak 16","pak16"},
        ["Σ PCB 7"]=new[]{"σ pcb 7","∑ pcb 7","pcb 7","pcb7"},["MKW C10-C22"]=new[]{"mkw (c10-c22)","mkw c10-c22"},["MKW C10-C40"]=new[]{"mkw (c10-c40)","mkw c10-c40"},
        ["Cyanide"]=new[]{"cyanide (ges.)","cyanide ges","cyanide"},["Phenole"]=new[]{"σ phenole","∑ phenole","phenole","phenol"},["Dibenz(a,h)anthracen"]=new[]{"dibenz(a,h)anthr.","dibenz(a,h)anthracen"},["Indeno(1,2,3-c,d)pyren"]=new[]{"indeno(1,2,3,c,d)pyren","indeno(1,2,3-c,d)pyren"}
    };
    static string Norm(string s)=>s.ToLowerInvariant().Replace("ä","a").Replace("ö","o").Replace("ü","u").Replace("gesamt","ges").Replace(".","");
    public static List<Reading> Parse(string text,MaterialTable table){
        var lines=text.Split('\n');var output=new List<Reading>();var used=new HashSet<string>();var section="";var unit="";
        for(int index=0;index<lines.Length;index++){
            var raw=lines[index].Trim();var n=Norm(raw).Replace('μ','µ');
            if(n.StartsWith("eluat")){section="Eluat";unit="";continue;}if(n.StartsWith("feststoff")){section="Feststoff";unit="";continue;}
            if(n.Contains("mg/kg")){unit="mg/kg";continue;}if(n.Contains("µg/l")||n.Contains("ug/l")){unit="µg/l";continue;}if(n.Contains("mg/l")){unit="mg/l";continue;}
            var matches=table.Limits.Select(l=>(l,len:(Aliases.TryGetValue(l.Name,out var a)?a:new[]{l.Name}).Select(Norm).Where(n.Contains).Select(x=>x.Length).DefaultIfEmpty(0).Max())).Where(x=>x.len>0).OrderByDescending(x=>x.len).Select(x=>x.l).ToList();
            if(matches.Count==0)continue;var line=raw;if(!Regex.IsMatch(line,@"<\s*BG",RegexOptions.IgnoreCase)&&!Regex.IsMatch(line,@"\d" )&&index+1<lines.Length)line+="  "+lines[index+1];
            string lu=Norm(line).Replace('μ','µ');Limit? selected=matches.FirstOrDefault(l=>l.Unit.Length==0||lu.Contains(Norm(l.Unit))||unit==Norm(l.Unit)||(unit=="mg/l"&&l.Unit=="µg/l")||(unit=="µg/l"&&l.Unit=="mg/l"));
            selected??=matches.FirstOrDefault(l=>section=="Feststoff"?(l.Unit.Contains("/kg")||l.Unit=="M%"):!(l.Unit.Contains("/kg")||l.Unit=="M%"));if(selected is null)continue;
            string scrub=line;foreach(var parameterAlias in Aliases.TryGetValue(selected.Name,out var aa)?aa:new[]{selected.Name})scrub=Regex.Replace(scrub,Regex.Escape(parameterAlias),"",RegexOptions.IgnoreCase);
            scrub=Regex.Replace(scrub,@"(?i)(PAK\s*1[56]|C10\s*-\s*C4[02]|mg\s*/\s*kg(?:\s*Ts\.?)?|[µμu]g\s*/\s*l|mg\s*/\s*l|[µμ]S\s*/\s*cm|\[\s*25\s*°?C\s*\]|°C)","");
            var tokens=Regex.Matches(scrub,@"(?i)<\s*BG|[-+]?\d+(?:[.,]\d+)?").Select(m=>m.Value).Take(2).ToArray();if(tokens.Length==0)continue;
            var texts=new[]{"–","–"};var values=new double?[2];var ignored=new bool[2];double factor=unit=="mg/l"&&selected.Unit=="µg/l"?1000:unit=="µg/l"&&selected.Unit=="mg/l"?.001:1;
            for(int i=0;i<tokens.Length;i++){ignored[i]=tokens[i].Contains("BG",StringComparison.OrdinalIgnoreCase);if(ignored[i])texts[i]="< BG";else if(double.TryParse(tokens[i].Replace(',','.'),NumberStyles.Float,CultureInfo.InvariantCulture,out var v)){values[i]=v*factor;texts[i]=values[i]!.Value.ToString("0.###",CultureInfo.GetCultureInfo("de-DE"));}}
            var key=selected.Name+"|"+selected.Unit;if(used.Add(key))output.Add(new(selected.Name,selected.Unit,texts,values,ignored));
        }return output;
    }
    public static string Stage(Reading r,int sample,MaterialTable t){if(r.Ignored[sample])return "ignoriert";if(r.Values[sample] is not double v)return "–";var l=t.Limits.First(x=>x.Name==r.Name&&x.Unit==r.Unit);if(l.Values.All(x=>x is null))return "ohne Grenzwert";for(int i=0;i<l.Values.Length;i++)if(l.Values[i] is double max&&v<=max&&(l.Minimum is null||v>=l.Minimum))return t.Classes[i];return "> "+t.Classes[^1];}
    public static string Overall(List<Reading> rs,int sample,MaterialTable t){var relevant=rs.Where(r=>r.Values[sample]!=null&&!r.Ignored[sample]&&r.Name!="pH-Wert"&&r.Name!="Elektrische Leitfähigkeit").ToList();if(relevant.Count==0)return "Keine Werte";for(int i=0;i<t.Classes.Length;i++)if(relevant.All(r=>{var l=t.Limits.First(x=>x.Name==r.Name&&x.Unit==r.Unit);return l.Values[i] is not double max||r.Values[sample]<=max;}))return t.Classes[i];return "> "+t.Classes[^1];}
}

class MainForm:Form
{
    readonly ComboBox material=new(){DropDownStyle=ComboBoxStyle.DropDownList,Width=180};readonly Button open=new(){Text="PDF auswählen",AutoSize=true};readonly Button reset=new(){Text="Nächste Bewertung",AutoSize=true,Enabled=false};readonly Button print=new(){Text="Drucken",AutoSize=true,Enabled=false};
    readonly Label status=new(){AutoSize=true,Text="PDF auswählen oder hier ablegen",Padding=new(8)};readonly Label result=new(){AutoSize=true,Font=new Font("Segoe UI",12,FontStyle.Bold),Padding=new(8)};readonly DataGridView grid=new(){Dock=DockStyle.Fill,AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill,AllowUserToAddRows=false,ReadOnly=true};
    List<Reading> readings=new();string fileName="";
    public MainForm(){Text="EBV Scan 1.5.0";Width=1180;Height=760;AllowDrop=true;material.Items.AddRange(EBVData.Tables.Keys.ToArray());material.SelectedItem="Boden";
        var top=new FlowLayoutPanel{Dock=DockStyle.Top,Height=52,Padding=new(8)};top.Controls.AddRange(new Control[]{new Label{Text="Material:",AutoSize=true,Padding=new(4,8,0,0)},material,open,reset,print,status});Controls.Add(grid);Controls.Add(result);result.Dock=DockStyle.Top;Controls.Add(top);
        open.Click+=async(_,_)=>{using var d=new OpenFileDialog{Filter="PDF-Dateien|*.pdf"};if(d.ShowDialog()==DialogResult.OK)await LoadPdf(d.FileName);};reset.Click+=(_,_)=>ResetEvaluation();print.Click+=(_,_)=>PrintReport();material.SelectedIndexChanged+=(_,_)=>{if(readings.Count>0)RefreshGrid();};DragEnter+=(_,e)=>{if(e.Data?.GetDataPresent(DataFormats.FileDrop)==true)e.Effect=DragDropEffects.Copy;};DragDrop+=async(_,e)=>{if(e.Data?.GetData(DataFormats.FileDrop) is string[] f&&f.Length>0&&Path.GetExtension(f[0]).Equals(".pdf",StringComparison.OrdinalIgnoreCase))await LoadPdf(f[0]);};
    }
    async Task LoadPdf(string path){try{UseWaitCursor=true;status.Text="PDF wird mit Windows OCR gelesen …";var text=await OcrPdf(path);fileName=Path.GetFileName(path);var lower=text.ToLowerInvariant();if(lower.Contains("gleisschotter")||lower.Contains("gs-"))material.SelectedItem="Gleisschotter";else if(lower.Contains("bauschutt")||lower.Contains("recycling")||lower.Contains("rc-"))material.SelectedItem="Bauschutt";else if(lower.Contains("boden"))material.SelectedItem="Boden";readings=Analyzer.Parse(text,EBVData.Tables[(string)material.SelectedItem!]);status.Text=$"{readings.Count} Messwerte erkannt · {fileName}";reset.Enabled=print.Enabled=true;RefreshGrid();}catch(Exception ex){MessageBox.Show(ex.Message,"PDF konnte nicht gelesen werden",MessageBoxButtons.OK,MessageBoxIcon.Error);}finally{UseWaitCursor=false;}}
    static async Task<string> OcrPdf(string path){var file=await StorageFile.GetFileFromPathAsync(path);var pdf=await PdfDocument.LoadFromFileAsync(file);var engine=OcrEngine.TryCreateFromUserProfileLanguages()??throw new InvalidOperationException("Windows OCR ist nicht verfügbar.");var sb=new StringBuilder();for(uint i=0;i<pdf.PageCount;i++){using var page=pdf.GetPage(i);using var stream=new InMemoryRandomAccessStream();await page.RenderToStreamAsync(stream,new PdfPageRenderOptions{DestinationWidth=2400});var decoder=await BitmapDecoder.CreateAsync(stream);var bitmap=await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8,BitmapAlphaMode.Premultiplied);var r=await engine.RecognizeAsync(bitmap);sb.AppendLine(r.Text);}return sb.ToString();}
    void RefreshGrid(){var t=EBVData.Tables[(string)material.SelectedItem!];grid.Columns.Clear();grid.Rows.Clear();foreach(var c in new[]{"Parameter","MP 1","Klasse MP 1","MP 2","Klasse MP 2"}.Concat(t.Classes))grid.Columns.Add(c,c);foreach(var r in readings){var row=new List<string>{r.Name+(r.Unit.Length>0?" ["+r.Unit+"]":""),r.Texts[0],r.Name is "pH-Wert" or "Elektrische Leitfähigkeit"?"Orientierung":Analyzer.Stage(r,0,t),r.Texts[1],r.Name is "pH-Wert" or "Elektrische Leitfähigkeit"?"Orientierung":Analyzer.Stage(r,1,t)};row.AddRange(t.Classes.Select((_,i)=>t.Limits.First(x=>x.Name==r.Name&&x.Unit==r.Unit).Values[i]?.ToString("0.###",CultureInfo.GetCultureInfo("de-DE"))??"–"));grid.Rows.Add(row.ToArray());}result.Text=$"MP 1: {Analyzer.Overall(readings,0,t)}     MP 2: {Analyzer.Overall(readings,1,t)}";}
    void ResetEvaluation(){readings.Clear();fileName="";grid.Columns.Clear();grid.Rows.Clear();result.Text="";status.Text="PDF auswählen oder hier ablegen";reset.Enabled=print.Enabled=false;}
    void PrintReport(){var doc=new PrintDocument{DocumentName="EBV Analyse"};doc.DefaultPageSettings.Landscape=true;doc.PrintPage+=(_,e)=>{var t=EBVData.Tables[(string)material.SelectedItem!];float y=40;e.Graphics!.DrawString("EBV Analyse – "+t.Title,new Font("Segoe UI",18,FontStyle.Bold),Brushes.Black,40,y);y+=38;e.Graphics.DrawString(fileName+"\nMP 1: "+Analyzer.Overall(readings,0,t)+"   MP 2: "+Analyzer.Overall(readings,1,t),new Font("Segoe UI",10),Brushes.Black,40,y);y+=48;foreach(var r in readings){e.Graphics.DrawString($"{r.Name} [{r.Unit}]   {r.Texts[0]} ({Analyzer.Stage(r,0,t)})   {r.Texts[1]} ({Analyzer.Stage(r,1,t)})",new Font("Segoe UI",8),Brushes.Black,40,y);y+=17;if(y>e.MarginBounds.Bottom){e.HasMorePages=true;return;}}};using var dlg=new PrintDialog{Document=doc};if(dlg.ShowDialog()==DialogResult.OK)doc.Print();}
}

static class Program
{
    [STAThread] static void Main(){ApplicationConfiguration.Initialize();Application.Run(new MainForm());}
}
