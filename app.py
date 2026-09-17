import os, sys, json, re, shutil, tempfile, zipfile, subprocess
from datetime import date, datetime
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageOps, ImageEnhance

APP = "CPK Report Generator"
NS_MAIN="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_XDR="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A="http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL="http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace('', NS_MAIN); ET.register_namespace('xdr',NS_XDR); ET.register_namespace('a',NS_A); ET.register_namespace('r',NS_R)

def base_dir():
    # Portable build: prefer resources placed next to the EXE.
    # Fall back to PyInstaller's internal bundle for compatibility.
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        if os.path.exists(os.path.join(exe_dir, 'config.json')):
            return exe_dir
        return getattr(sys, '_MEIPASS', exe_dir)
    return os.path.dirname(os.path.abspath(__file__))

def load_config():
    with open(os.path.join(base_dir(),'config.json'),encoding='utf-8') as f: return json.load(f)

def excel_serial(d):
    return (d - date(1899,12,30)).days

def parse_sysdata(path):
    rows=[]
    with open(path,'r',encoding='utf-8',errors='ignore') as f:
        for line in f:
            p=re.split(r'[\t,; ]+',line.strip())
            if len(p)<6: continue
            try: rows.append([float(x) for x in p[:6]])
            except ValueError: pass
    if len(rows)<100: raise ValueError(f"sysdata 有效資料只有 {len(rows)} 筆，至少需要 100 筆。")
    return rows[-100:]

def cell_ref_colrow(ref):
    m=re.match(r'([A-Z]+)(\d+)',ref); return m.group(1),int(m.group(2))

def set_cell(root, ref, value, numeric=False):
    ns={'m':NS_MAIN}; sheetdata=root.find('m:sheetData',ns)
    _, rownum=cell_ref_colrow(ref)
    row=None
    for r in sheetdata.findall('m:row',ns):
        if int(r.attrib['r'])==rownum: row=r; break
    if row is None:
        row=ET.SubElement(sheetdata,f'{{{NS_MAIN}}}row',{'r':str(rownum)})
    c=None
    for x in row.findall('m:c',ns):
        if x.attrib.get('r')==ref: c=x; break
    if c is None: c=ET.SubElement(row,f'{{{NS_MAIN}}}c',{'r':ref})
    for ch in list(c):
        if ch.tag in (f'{{{NS_MAIN}}}v',f'{{{NS_MAIN}}}is'): c.remove(ch)
    if numeric:
        c.attrib.pop('t',None); v=ET.SubElement(c,f'{{{NS_MAIN}}}v'); v.text=str(value)
    else:
        c.attrib['t']='inlineStr'; isel=ET.SubElement(c,f'{{{NS_MAIN}}}is'); t=ET.SubElement(isel,f'{{{NS_MAIN}}}t'); t.text=str(value)

def img_to_png_bytes(path):
    import io
    im=Image.open(path).convert('RGB'); out=io.BytesIO(); im.save(out,format='PNG',optimize=True); return out.getvalue()

def make_anchor(rid, idx, name, fr, to, ext):
    A=f'{{{NS_XDR}}}'; AA=f'{{{NS_A}}}'; RR=f'{{{NS_R}}}'
    an=ET.Element(A+'twoCellAnchor',{'editAs':'oneCell'})
    f=ET.SubElement(an,A+'from');
    for tag,val in zip(['col','colOff','row','rowOff'],fr): ET.SubElement(f,A+tag).text=str(val)
    t=ET.SubElement(an,A+'to');
    for tag,val in zip(['col','colOff','row','rowOff'],to): ET.SubElement(t,A+tag).text=str(val)
    pic=ET.SubElement(an,A+'pic'); nv=ET.SubElement(pic,A+'nvPicPr'); ET.SubElement(nv,A+'cNvPr',{'id':str(idx),'name':name,'descr':''}); ET.SubElement(nv,A+'cNvPicPr')
    bf=ET.SubElement(pic,A+'blipFill'); ET.SubElement(bf,AA+'blip',{RR+'embed':rid}); ET.SubElement(bf,AA+'stretch')
    sp=ET.SubElement(pic,A+'spPr'); xf=ET.SubElement(sp,AA+'xfrm'); ET.SubElement(xf,AA+'off',{'x':str(ext[0]),'y':str(ext[1])}); ET.SubElement(xf,AA+'ext',{'cx':str(ext[2]),'cy':str(ext[3])}); pg=ET.SubElement(sp,AA+'prstGeom',{'prst':'rect'}); ET.SubElement(pg,AA+'avLst'); ET.SubElement(sp,AA+'noFill'); ln=ET.SubElement(sp,AA+'ln',{'w':'0'}); ET.SubElement(ln,AA+'noFill'); ET.SubElement(an,A+'clientData'); return an

ANCHORS={
 'cpk1': ((1,0,133,0),(7,521280,150,161640),(765000,28835280,5878440,3562200)),
 'cpk2': ((1,0,154,0),(7,521280,171,162000),(765000,33045480,5878440,3562200)),
 'up':   ((1,0,190,0),(7,521280,207,161640),(765000,40265280,5878440,3562200)),
 'down': ((1,0,211,0),(7,521280,228,152280),(765000,44475480,5878440,3562200)),
}

def generate_report(template, outpath, fields, cpk, data, images):
    with zipfile.ZipFile(template,'r') as zin:
        parts={n:zin.read(n) for n in zin.namelist()}
    sheet='xl/worksheets/sheet1.xml'; root=ET.fromstring(parts[sheet])
    set_cell(root,'D4',fields['customer']); set_cell(root,'D5',fields['model']); set_cell(root,'G5',fields['serial'])
    eng=fields['eng1'] + ((" / "+fields['eng2']) if fields['eng2'] else '')
    set_cell(root,'D6',eng); set_cell(root,'G6',fields['date'].strftime('%Y/%m/%d'))
    for i,v in enumerate(cpk): set_cell(root,chr(ord('C')+i)+'14',f'{v:.3f}',numeric=True)
    for i,row in enumerate(data, start=20):
        for j,v in enumerate(row): set_cell(root,chr(ord('C')+j)+str(i),v,numeric=True)
    parts[sheet]=ET.tostring(root,encoding='utf-8',xml_declaration=True)

    draw='xl/drawings/drawing1.xml'; relp='xl/drawings/_rels/drawing1.xml.rels'
    droot=ET.fromstring(parts[draw]); rroot=ET.fromstring(parts[relp])
    # Keep company stamp topmost: locate existing rId2 anchor and append it last.
    stamp=None
    for an in list(droot):
        blip=an.find(f'.//{{{NS_A}}}blip')
        if blip is not None and blip.attrib.get(f'{{{NS_R}}}embed')=='rId2': stamp=an; droot.remove(an); break
    existing_ids=[]
    for rel in rroot: 
        m=re.match(r'rId(\d+)',rel.attrib.get('Id',''))
        if m: existing_ids.append(int(m.group(1)))
    nextid=max(existing_ids or [0])+1; picidx=100
    for key in ('cpk1','cpk2','up','down'):
        p=images.get(key)
        if not p: continue
        rid=f'rId{nextid}'; nextid+=1; media=f'image_auto_{key}.png'
        ET.SubElement(rroot,f'{{{NS_REL}}}Relationship',{'Id':rid,'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image','Target':'../media/'+media})
        parts['xl/media/'+media]=img_to_png_bytes(p)
        fr,to,ext=ANCHORS[key]; droot.append(make_anchor(rid,picidx,key,fr,to,ext)); picidx+=1
    if stamp is not None: droot.append(stamp)
    parts[draw]=ET.tostring(droot,encoding='utf-8',xml_declaration=True); parts[relp]=ET.tostring(rroot,encoding='utf-8',xml_declaration=True)
    with zipfile.ZipFile(outpath,'w',zipfile.ZIP_DEFLATED) as zout:
        for n,b in parts.items(): zout.writestr(n,b)

def find_tesseract():
    local=os.path.join(base_dir(),'ocr','tesseract.exe')
    if os.path.exists(local): return local
    return shutil.which('tesseract')

def ocr_cpk(path):
    exe=find_tesseract()
    if not exe: raise RuntimeError('找不到 Portable OCR 引擎 ocr/tesseract.exe。')
    tessdata=os.path.join(os.path.dirname(exe),'tessdata')
    src=ImageOps.autocontrast(Image.open(path).convert('L')); W,H=src.size
    vals=[]
    for row in range(3):
        for col in range(2):
            panel=src.crop((int(col*W/2),int(row*H/3),int((col+1)*W/2),int((row+1)*H/3)))
            roi=panel.crop((0,int(panel.height*0.60),panel.width,panel.height))
            roi=ImageEnhance.Contrast(roi).enhance(2.2).resize((roi.width*3,roi.height*3))
            tmp=tempfile.NamedTemporaryFile(suffix='.png',delete=False); tmp.close(); roi.save(tmp.name)
            try:
                cmd=[exe,tmp.name,'stdout','--psm','6','-l','eng','-c','tessedit_char_whitelist=CcpkPK=:.0123456789']
                if os.path.isdir(tessdata): cmd += ['--tessdata-dir',tessdata]
                cp=subprocess.run(cmd,capture_output=True,text=True,errors='ignore',timeout=20)
                txt=cp.stdout.replace(' ','')
            finally:
                try: os.unlink(tmp.name)
                except: pass
            m=re.search(r'Cpk[=:]?([0-9]+(?:\.[0-9]+)?)',txt,re.I)
            if m: vals.append(float(m.group(1)))
            else:
                nums=re.findall(r'([0-9]+\.[0-9]{2,4})',txt)
                vals.append(float(nums[-1]) if nums else None)
    if any(v is None for v in vals):
        raise RuntimeError(f'OCR 成功辨識 {sum(v is not None for v in vals)}/6 個 Cpk。請在預覽視窗確認並手動補正。')
    return vals

def ocr_cpk_panels(src, indices):
    exe=find_tesseract()
    if not exe: raise RuntimeError('找不到 Portable OCR 引擎 ocr/tesseract.exe。')
    tessdata=os.path.join(os.path.dirname(exe),'tessdata')
    gray=ImageOps.autocontrast(src.convert('L')); W,H=gray.size
    result={}
    for idx in indices:
        row, col = divmod(idx, 2)
        panel=gray.crop((int(col*W/2),int(row*H/3),int((col+1)*W/2),int((row+1)*H/3)))
        roi=panel.crop((0,int(panel.height*0.55),panel.width,panel.height))
        roi=ImageEnhance.Contrast(roi).enhance(2.5).resize((max(1,roi.width*4),max(1,roi.height*4)))
        tmp=tempfile.NamedTemporaryFile(suffix='.png',delete=False); tmp.close(); roi.save(tmp.name)
        try:
            cmd=[exe,tmp.name,'stdout','--psm','6','-l','eng','-c','tessedit_char_whitelist=CcpkPK=:.0123456789']
            if os.path.isdir(tessdata): cmd += ['--tessdata-dir',tessdata]
            cp=subprocess.run(cmd,capture_output=True,text=True,errors='ignore',timeout=20)
            txt=cp.stdout.replace(' ','')
        finally:
            try: os.unlink(tmp.name)
            except: pass
        m=re.search(r'Cpk[=:]?([0-9]+(?:\.[0-9]+)?)',txt,re.I)
        if m: result[idx]=float(m.group(1)); continue
        nums=re.findall(r'([0-9]+\.[0-9]{2,4})',txt)
        result[idx]=float(nums[-1]) if nums else None
    return result

def convert_excel_to_pdf(xlsx, pdf, outdir):
    # Prefer Microsoft Excel itself for faithful PDF rendering.
    ps1 = """param([string]$xlsx,[string]$pdf)
$ErrorActionPreference='Stop'
$excel=$null; $wb=$null
try {
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false
  $excel.DisplayAlerts=$false
  $excel.AskToUpdateLinks=$false
  $wb=$excel.Workbooks.Open($xlsx,0,$true)
  $wb.ExportAsFixedFormat(0,$pdf,0,$true,$false)
  $wb.Close($false)
} finally {
  if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
  if ($excel -ne $null) { try { $excel.Quit() } catch {} }
}
"""
    script=os.path.join(tempfile.gettempdir(),'cpk_export_pdf.ps1')
    with open(script,'w',encoding='utf-8-sig') as f: f.write(ps1)
    try:
        r=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',script,'-xlsx',os.path.abspath(xlsx),'-pdf',os.path.abspath(pdf)],capture_output=True,text=True,timeout=180)
        if r.returncode==0 and os.path.exists(pdf) and os.path.getsize(pdf)>0: return 'Microsoft Excel（原版面匯出）'
    except Exception: pass
    candidates=[os.path.join(base_dir(),'libreoffice','program','soffice.exe'),os.path.join(base_dir(),'libreoffice','soffice.exe')]
    soffice=next((p for p in candidates if os.path.exists(p)),None)
    if not soffice: raise RuntimeError('PDF 轉換失敗：此電腦找不到 Microsoft Excel，且 Portable LibreOffice 不存在。')
    profile=os.path.join(tempfile.gettempdir(),'CPK_Report_LO_Profile'); os.makedirs(profile,exist_ok=True)
    uri='file:///'+profile.replace(chr(92),'/')
    r=subprocess.run([soffice,'-env:UserInstallation='+uri,'--headless','--convert-to','pdf','--outdir',outdir,xlsx],capture_output=True,text=True,timeout=180)
    generated=os.path.join(outdir,os.path.splitext(os.path.basename(xlsx))[0]+'.pdf')
    if r.returncode!=0 or not os.path.exists(generated): raise RuntimeError('LibreOffice PDF 轉換失敗：'+(r.stderr or r.stdout)[-500:])
    if os.path.abspath(generated)!=os.path.abspath(pdf): shutil.move(generated,pdf)
    return 'LibreOffice Portable（相容模式，複雜版面可能略有差異）'
class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP); self.geometry('610x720'); self.resizable(False,False); self.cfg=load_config(); self.files={k:'' for k in ['cpk1','cpk2','up','down','sys']}; self.build()
    def build(self):
        f=ttk.Frame(self,padding=16); f.pack(fill='both',expand=True); r=0
        self.customer=self.entry(f,r,'客戶名稱'); r+=1
        self.model=self.combo(f,r,'機台型號',self.cfg['machine_models']); r+=1
        self.serial=self.entry(f,r,'機台序號'); r+=1
        self.eng1=self.combo(f,r,'工程師 1',self.cfg['engineers']); r+=1
        self.eng2=self.combo(f,r,'工程師 2（可選）',['']+self.cfg['engineers']); r+=1
        ttk.Label(f,text='選擇日期').grid(row=r,column=0,sticky='w',pady=5); self.datev=tk.StringVar(value=date.today().strftime('%Y/%m/%d')); ttk.Entry(f,textvariable=self.datev,width=30).grid(row=r,column=1,sticky='w'); r+=1
        for key,label in [('cpk1','圖片：CPK1'),('cpk2','圖片：CPK2'),('up','圖片：Camera Up'),('down','圖片：Camera Down'),('sys','sysdata.txt')]:
            ttk.Label(f,text=label).grid(row=r,column=0,sticky='w',pady=4); lab=ttk.Label(f,text='未選擇',width=28); lab.grid(row=r,column=1,sticky='w'); ttk.Button(f,text='選擇',command=lambda k=key,l=lab:self.choose(k,l)).grid(row=r,column=2); r+=1
        ttk.Separator(f).grid(row=r,column=0,columnspan=3,sticky='ew',pady=10); r+=1
        ttk.Label(f,text='CPK 數值（可手動修改）',font=('',10,'bold')).grid(row=r,column=0,columnspan=2,sticky='w'); r+=1
        self.cpk=[]
        for name in ['FormerX','FormerY','FTheta','LaterX','LaterY','LTheta']:
            ttk.Label(f,text=name).grid(row=r,column=0,sticky='e',padx=8,pady=3); v=tk.StringVar(); ttk.Entry(f,textvariable=v,width=18).grid(row=r,column=1,sticky='w'); self.cpk.append(v); r+=1
        ttk.Button(f,text='產出報告',command=self.generate).grid(row=r,column=0,columnspan=3,pady=(18,8),ipadx=55,ipady=5); r+=1
        ttk.Button(f,text='查看說明',command=self.help).grid(row=r,column=0,columnspan=3,pady=4)
    def entry(self,f,r,label):
        ttk.Label(f,text=label).grid(row=r,column=0,sticky='w',pady=5); v=tk.StringVar(); ttk.Entry(f,textvariable=v,width=32).grid(row=r,column=1,columnspan=2,sticky='w'); return v
    def combo(self,f,r,label,vals):
        ttk.Label(f,text=label).grid(row=r,column=0,sticky='w',pady=5); v=tk.StringVar(); ttk.Combobox(f,textvariable=v,values=vals,width=29).grid(row=r,column=1,columnspan=2,sticky='w'); return v
    def choose(self,key,label):
        ft=[('Text','*.txt')] if key=='sys' else [('Images','*.png *.jpg *.jpeg *.bmp'),('All','*.*')]
        p=filedialog.askopenfilename(filetypes=ft)
        if not p:return
        self.files[key]=p; label.config(text=os.path.basename(p)[:28])
        if key=='cpk1': self.preview_ocr(p)
    def preview_ocr(self,p):
        vals=None; err=''
        try: vals=ocr_cpk(p)
        except Exception as e: err=str(e)
        w=tk.Toplevel(self); w.title('CPK1 辨識結果預覽'); w.geometry('1120x780')
        left=ttk.Frame(w); left.pack(side='left',fill='both',expand=True,padx=8,pady=8)
        right=ttk.Frame(w,width=320); right.pack(side='right',fill='y',padx=8,pady=8)
        cv=tk.Canvas(left,bg='#202020',highlightthickness=0); cv.pack(fill='both',expand=True)
        src=Image.open(p).convert('RGB'); state={'scale':1.0,'photo':None}
        def redraw(keep_view=True):
            xv=cv.xview(); yv=cv.yview(); sc=state['scale']; im=src.resize((max(1,int(src.width*sc)),max(1,int(src.height*sc))))
            state['photo']=ImageTk.PhotoImage(im); cv.delete('all'); cv.create_image(10,10,image=state['photo'],anchor='nw',tags='img'); cv.config(scrollregion=(0,0,im.width+20,im.height+20))
            if keep_view and xv and yv:
                cv.xview_moveto(xv[0]); cv.yview_moveto(yv[0])
        def zoom(factor):
            state['scale']=max(0.15,min(6.0,state['scale']*factor)); redraw()
        def fit():
            w.update_idletasks(); aw=max(100,cv.winfo_width()-20); ah=max(100,cv.winfo_height()-20); state['scale']=min(aw/src.width,ah/src.height); redraw(False); cv.xview_moveto(0); cv.yview_moveto(0)
        cv.bind('<MouseWheel>',lambda e: zoom(1.15 if e.delta>0 else 1/1.15)); cv.bind('<Button-4>',lambda e: zoom(1.15)); cv.bind('<Button-5>',lambda e: zoom(1/1.15))
        cv.bind('<ButtonPress-1>',lambda e:cv.scan_mark(e.x,e.y)); cv.bind('<B1-Motion>',lambda e:cv.scan_dragto(e.x,e.y,gain=1))
        bar=ttk.Frame(right); bar.pack(fill='x',pady=(0,8))
        ttk.Button(bar,text='－',width=4,command=lambda:zoom(1/1.2)).pack(side='left'); ttk.Button(bar,text='＋',width=4,command=lambda:zoom(1.2)).pack(side='left',padx=4); ttk.Button(bar,text='100%',command=lambda:(state.update(scale=1.0),redraw(False))).pack(side='left'); ttk.Button(bar,text='適合視窗',command=fit).pack(side='left',padx=4)
        ttk.Label(right,text='滑鼠滾輪：縮放\n按住左鍵拖曳：移動畫面',justify='left').pack(anchor='w',pady=(0,8))
        status=tk.StringVar(value=err if err else '已完成第一次自動辨識。')
        ttk.Label(right,textvariable=status,foreground='firebrick' if err else 'black',wraplength=300).pack(anchor='w',pady=5)
        edits=[]; names=['FormerX','FormerY','FTheta','LaterX','LaterY','LTheta']
        for i,n in enumerate(names):
            row=ttk.Frame(right); row.pack(fill='x',pady=4); ttk.Label(row,text=n,width=10).pack(side='left'); v=tk.StringVar(value=(f'{vals[i]:.3f}' if vals and vals[i] is not None else self.cpk[i].get())); ttk.Entry(row,textvariable=v,width=12).pack(side='left'); edits.append(v)
        def set_results(res, label):
            ok=0
            for idx,val in res.items():
                if val is not None: edits[idx].set(f'{val:.3f}'); ok+=1
            status.set(f'{label}：成功辨識 {ok}/{len(res)} 個 Cpk。請確認數值。')
        def rerun_all():
            try: set_results(ocr_cpk_panels(src,range(6)),'重新辨識全部')
            except Exception as e: status.set(str(e))
        def rerun_visible():
            try:
                sc=state['scale']; x0=max(0,(cv.canvasx(0)-10)/sc); y0=max(0,(cv.canvasy(0)-10)/sc); x1=min(src.width,(cv.canvasx(cv.winfo_width())-10)/sc); y1=min(src.height,(cv.canvasy(cv.winfo_height())-10)/sc)
                indices=[]
                for idx in range(6):
                    rr,cc=divmod(idx,2); px0=cc*src.width/2; px1=(cc+1)*src.width/2; py0=rr*src.height/3; py1=(rr+1)*src.height/3
                    overlap=max(0,min(x1,px1)-max(x0,px0))*max(0,min(y1,py1)-max(y0,py0)); area=(px1-px0)*(py1-py0)
                    if area and overlap/area >= 0.20: indices.append(idx)
                if not indices: status.set('目前畫面沒有足夠完整的 CPK 區塊，請移動或縮小一點再辨識。'); return
                set_results(ocr_cpk_panels(src,indices),'辨識目前畫面')
            except Exception as e: status.set(str(e))
        btns=ttk.Frame(right); btns.pack(fill='x',pady=(10,4))
        ttk.Button(btns,text='重新辨識全部',command=rerun_all).pack(side='left',padx=(0,6)); ttk.Button(btns,text='辨識目前畫面',command=rerun_visible).pack(side='left')
        ttk.Label(right,text='「辨識目前畫面」會依你目前放大/拖曳後可見的 CPK 區塊重新辨識，並只更新那些欄位。',wraplength=300).pack(anchor='w',pady=(4,10))
        def apply():
            try:
                for i,v in enumerate(edits): self.cpk[i].set(f'{float(v.get()):.3f}')
                w.destroy()
            except: messagebox.showerror('錯誤','六個 CPK 都必須是數字。',parent=w)
        ttk.Button(right,text='確認套用',command=apply).pack(pady=12,ipadx=30)
        w.after(150,fit)

    def _safe_name(self, text):
        return re.sub(r'[<>:"/\\|?*]+','_',text).strip().rstrip('.')

    def _unique_folder(self, root, name):
        p=os.path.join(root,name); n=2
        while os.path.exists(p): p=os.path.join(root,f'{name}_{n}'); n+=1
        os.makedirs(p,exist_ok=False); return p

    def _copy_materials(self, folder):
        names={'cpk1':'CPK1','cpk2':'CPK2','up':'up','down':'down','sys':'sysdata'}
        for k,src in self.files.items():
            if not src: continue
            ext=os.path.splitext(src)[1] or ('.txt' if k=='sys' else '')
            shutil.copy2(src,os.path.join(folder,names[k]+ext.lower()))


    def generate(self):
        try:
            if not self.customer.get().strip() or not self.model.get().strip() or not self.serial.get().strip() or not self.eng1.get().strip(): raise ValueError('客戶名稱、機台型號、機台序號、工程師 1 為必填。')
            if not self.files['sys']: raise ValueError('請選擇 sysdata.txt。')
            vals=[float(v.get()) for v in self.cpk]; d=datetime.strptime(self.datev.get().strip(),'%Y/%m/%d').date(); data=parse_sysdata(self.files['sys'])
            has_ccd=bool(self.files['up'] or self.files['down']); tname='sample.xlsx' if has_ccd else 'sample_without_ccd.xlsx'; template=os.path.join(base_dir(),'templates',tname)
            fields={'customer':self.customer.get().strip(),'model':self.model.get().strip(),'serial':self.serial.get().strip(),'eng1':self.eng1.get().strip(),'eng2':self.eng2.get().strip(),'date':d}
            base=self._safe_name(f"{fields['customer']}-{fields['model']}-{fields['serial']}-{d.isoformat()}")
            parent=os.path.join(base_dir(),'Report History'); os.makedirs(parent,exist_ok=True)
            folder=self._unique_folder(parent,base); xlsx=os.path.join(folder,base+'.xlsx'); pdf=os.path.join(folder,base+'.pdf')
            generate_report(template,xlsx,fields,vals,data,self.files); self._copy_materials(folder)
            pdf_engine=convert_excel_to_pdf(xlsx,pdf,folder)
            # Open PDF and report folder on Windows.
            try: os.startfile(pdf)
            except Exception: pass
            try: os.startfile(folder)
            except Exception: pass
            note='' if pdf_engine.startswith('Microsoft Excel') else '\n\n注意：本機未使用 Microsoft Excel 匯出 PDF，目前為 LibreOffice 相容模式；若版面與 Excel 不一致，請以 Excel 檔為準。'
            messagebox.showinfo('完成',f'報告已產生：\n{folder}\n\nExcel、PDF、照片與 sysdata 已整理在同一資料夾。\nPDF 引擎：{pdf_engine}{note}')
        except Exception as e: messagebox.showerror('無法產出報告',str(e))

    def help(self):
        messagebox.showinfo('使用說明','1. 輸入客戶/機型/序號/工程師/日期（YYYY/MM/DD）。\n2. 選 CPK1 後會自動 OCR；預覽可縮放/拖曳，並可按「重新辨識全部」或「辨識目前畫面」。\n3. 選 sysdata，程式取最後 100 筆有效資料的前 6 欄。\n4. Camera Up/Down 任一有提供時自動使用 CCD 範本。\n5. 報告固定儲存在程式旁的 Report History\\客戶-機型-序號-日期。\n6. PDF 直接由完成後的 Excel 轉換；完成後自動開啟 PDF 與該次報告資料夾。')
if __name__=='__main__': App().mainloop()
