import os, sys, json, re, shutil, tempfile, zipfile, subprocess
from datetime import date, datetime
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageOps, ImageEnhance
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

APP = "CPK Report Generator"
NS_MAIN="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_XDR="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A="http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL="http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace('', NS_MAIN); ET.register_namespace('xdr',NS_XDR); ET.register_namespace('a',NS_A); ET.register_namespace('r',NS_R)

def base_dir():
    return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

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
    if not exe: raise RuntimeError('Portable OCR 引擎尚未放入 ocr/tesseract.exe。可先手動輸入 CPK。')
    im=Image.open(path).convert('L'); im=ImageOps.autocontrast(im); im=ImageEnhance.Contrast(im).enhance(1.8); im=im.resize((im.width*2,im.height*2))
    tmp=tempfile.NamedTemporaryFile(suffix='.png',delete=False); tmp.close(); im.save(tmp.name)
    try:
        cp=subprocess.run([exe,tmp.name,'stdout','--psm','6','-l','eng'],capture_output=True,text=True,errors='ignore',timeout=30)
        txt=cp.stdout
    finally:
        try: os.unlink(tmp.name)
        except: pass
    vals=[float(x) for x in re.findall(r'\bCpk\s*[=:]\s*([0-9]+(?:\.[0-9]+)?)',txt,re.I)]
    if len(vals)<6: raise RuntimeError(f'OCR 找到 {len(vals)} 個 Cpk，未達 6 個。請在預覽視窗手動修正。')
    return vals[:6]

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
        w=tk.Toplevel(self); w.title('CPK1 辨識結果預覽'); w.geometry('1100x760')
        left=ttk.Frame(w); left.pack(side='left',fill='both',expand=True,padx=8,pady=8)
        right=ttk.Frame(w,width=300); right.pack(side='right',fill='y',padx=8,pady=8)
        cv=tk.Canvas(left,bg='#202020',highlightthickness=0); cv.pack(fill='both',expand=True)
        src=Image.open(p).convert('RGB'); state={'scale':1.0,'photo':None,'x':0,'y':0}
        def redraw():
            sc=state['scale']; im=src.resize((max(1,int(src.width*sc)),max(1,int(src.height*sc))))
            state['photo']=ImageTk.PhotoImage(im); cv.delete('all'); cv.create_image(10,10,image=state['photo'],anchor='nw',tags='img'); cv.config(scrollregion=(0,0,im.width+20,im.height+20))
        def zoom(factor):
            state['scale']=max(0.15,min(5.0,state['scale']*factor)); redraw()
        def fit():
            w.update_idletasks(); aw=max(100,cv.winfo_width()-20); ah=max(100,cv.winfo_height()-20); state['scale']=min(aw/src.width,ah/src.height); redraw()
        cv.bind('<MouseWheel>',lambda e: zoom(1.15 if e.delta>0 else 1/1.15))
        cv.bind('<Button-4>',lambda e: zoom(1.15)); cv.bind('<Button-5>',lambda e: zoom(1/1.15))
        drag={'x':0,'y':0}
        cv.bind('<ButtonPress-1>',lambda e:(cv.scan_mark(e.x,e.y),drag.update(x=e.x,y=e.y)))
        cv.bind('<B1-Motion>',lambda e:cv.scan_dragto(e.x,e.y,gain=1))
        bar=ttk.Frame(right); bar.pack(fill='x',pady=(0,10))
        ttk.Button(bar,text='－',width=4,command=lambda:zoom(1/1.2)).pack(side='left'); ttk.Button(bar,text='＋',width=4,command=lambda:zoom(1.2)).pack(side='left',padx=4); ttk.Button(bar,text='100%',command=lambda:(state.update(scale=1.0),redraw())).pack(side='left'); ttk.Button(bar,text='適合視窗',command=fit).pack(side='left',padx=4)
        ttk.Label(right,text='滑鼠滾輪：縮放\n按住左鍵拖曳：移動畫面').pack(anchor='w',pady=(0,10))
        if err: ttk.Label(right,text=err,foreground='firebrick',wraplength=280).pack(anchor='w',pady=5)
        edits=[]; names=['FormerX','FormerY','FTheta','LaterX','LaterY','LTheta']
        for i,n in enumerate(names):
            row=ttk.Frame(right); row.pack(fill='x',pady=4); ttk.Label(row,text=n,width=10).pack(side='left'); v=tk.StringVar(value=(f'{vals[i]:.3f}' if vals else self.cpk[i].get())); ttk.Entry(row,textvariable=v,width=12).pack(side='left'); edits.append(v)
        def apply():
            try:
                for i,v in enumerate(edits): self.cpk[i].set(f'{float(v.get()):.3f}')
                w.destroy()
            except: messagebox.showerror('錯誤','六個 CPK 都必須是數字。',parent=w)
        ttk.Button(right,text='確認套用',command=apply).pack(pady=18,ipadx=30)
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

    def _make_pdf(self, path, fields, cpk, data):
        # Portable PDF preview: generated without requiring Excel/Office.
        c=pdfcanvas.Canvas(path,pagesize=A4); W,H=A4
        c.setFont('Helvetica-Bold',15); c.drawString(36,H-42,'CPK Report')
        c.setFont('Helvetica',9); y=H-62
        info=[('Customer',fields['customer']),('Model',fields['model']),('Serial',fields['serial']),('Engineer',fields['eng1']+((' / '+fields['eng2']) if fields['eng2'] else '')),('Date',fields['date'].strftime('%Y/%m/%d'))]
        for k,v in info: c.drawString(36,y,f'{k}: {v}'); y-=14
        names=['FormerX','FormerY','FTheta','LaterX','LaterY','LTheta']; y-=4
        c.setFont('Helvetica-Bold',9); c.drawString(36,y,'CPK'); y-=14; c.setFont('Helvetica',8)
        for i,n in enumerate(names): c.drawString(36+(i%3)*170,y-(i//3)*14,f'{n}: {cpk[i]:.3f}')
        y-=42
        # Image pages keep originals readable.
        c.setFont('Helvetica',6.5); colx=[36,112,188,264,340,416]; c.drawString(18,y,'No.')
        for j,n in enumerate(names): c.drawString(colx[j],y,n)
        y-=10
        for idx,row in enumerate(data,1):
            if y<40: c.showPage(); y=H-40; c.setFont('Helvetica',6.5)
            c.drawRightString(30,y,str(idx))
            for j,v in enumerate(row): c.drawRightString(colx[j]+45,y,f'{v:.3f}')
            y-=7
        for key,title in [('cpk1','CPK1'),('cpk2','CPK2'),('up','Camera Up'),('down','Camera Down')]:
            src=self.files.get(key)
            if not src: continue
            c.showPage(); c.setFont('Helvetica-Bold',12); c.drawString(36,H-38,title)
            try:
                im=Image.open(src); iw,ih=im.size; maxw,maxh=W-72,H-90; sc=min(maxw/iw,maxh/ih); dw,dh=iw*sc,ih*sc
                c.drawImage(ImageReader(im),36,H-60-dh,width=dw,height=dh,preserveAspectRatio=True,mask='auto')
            except Exception: pass
        c.save()

    def generate(self):
        try:
            if not self.customer.get().strip() or not self.model.get().strip() or not self.serial.get().strip() or not self.eng1.get().strip(): raise ValueError('客戶名稱、機台型號、機台序號、工程師 1 為必填。')
            if not self.files['sys']: raise ValueError('請選擇 sysdata.txt。')
            vals=[float(v.get()) for v in self.cpk]; d=datetime.strptime(self.datev.get().strip(),'%Y/%m/%d').date(); data=parse_sysdata(self.files['sys'])
            has_ccd=bool(self.files['up'] or self.files['down']); tname='sample.xlsx' if has_ccd else 'sample_without_ccd.xlsx'; template=os.path.join(base_dir(),'templates',tname)
            fields={'customer':self.customer.get().strip(),'model':self.model.get().strip(),'serial':self.serial.get().strip(),'eng1':self.eng1.get().strip(),'eng2':self.eng2.get().strip(),'date':d}
            base=self._safe_name(f"{fields['customer']}-{fields['model']}-{fields['serial']}-{d.isoformat()}")
            parent=filedialog.askdirectory(title='選擇報告儲存位置');
            if not parent:return
            folder=self._unique_folder(parent,base); xlsx=os.path.join(folder,base+'.xlsx'); pdf=os.path.join(folder,base+'.pdf')
            generate_report(template,xlsx,fields,vals,data,self.files); self._copy_materials(folder); self._make_pdf(pdf,fields,vals,data)
            # Open PDF and report folder on Windows.
            try: os.startfile(pdf)
            except Exception: pass
            try: os.startfile(folder)
            except Exception: pass
            messagebox.showinfo('完成',f'報告已產生：\n{folder}\n\nExcel、PDF、照片與 sysdata 已整理在同一資料夾。')
        except Exception as e: messagebox.showerror('無法產出報告',str(e))

    def help(self):
        messagebox.showinfo('使用說明','1. 輸入客戶/機型/序號/工程師/日期（YYYY/MM/DD）。\n2. 選 CPK1 後會開啟可縮放/拖曳的 OCR 預覽。\n3. 選 sysdata，程式取最後 100 筆有效資料的前 6 欄。\n4. Camera Up/Down 任一有提供時自動使用 CCD 範本。\n5. 產出時選擇父資料夾，程式自動建立「客戶-機型-序號-日期」資料夾，並放入 Excel、PDF、照片與 sysdata。\n6. 完成後自動開啟 PDF 與儲存資料夾。')
if __name__=='__main__': App().mainloop()
