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
    set_cell(root,'D6',eng); set_cell(root,'G6',excel_serial(fields['date']),numeric=True)
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
        w=tk.Toplevel(self); w.title('CPK1 辨識結果預覽'); w.geometry('900x680');
        im=Image.open(p); im.thumbnail((820,430)); photo=ImageTk.PhotoImage(im); il=ttk.Label(w,image=photo); il.image=photo; il.pack(pady=10)
        if err: ttk.Label(w,text=err,foreground='firebrick').pack()
        frame=ttk.Frame(w); frame.pack(pady=8); edits=[]
        names=['FormerX','FormerY','FTheta','LaterX','LaterY','LTheta']
        for i,n in enumerate(names):
            ttk.Label(frame,text=n).grid(row=i//3,column=(i%3)*2,padx=5,pady=4); v=tk.StringVar(value=(f'{vals[i]:.3f}' if vals else self.cpk[i].get())); ttk.Entry(frame,textvariable=v,width=10).grid(row=i//3,column=(i%3)*2+1); edits.append(v)
        def apply():
            try:
                for i,v in enumerate(edits): float(v.get()); self.cpk[i].set(f'{float(v.get()):.3f}')
                w.destroy()
            except: messagebox.showerror('錯誤','六個 CPK 都必須是數字。',parent=w)
        ttk.Button(w,text='確認套用',command=apply).pack(pady=10,ipadx=30)
    def generate(self):
        try:
            if not self.customer.get().strip() or not self.model.get().strip() or not self.serial.get().strip() or not self.eng1.get().strip(): raise ValueError('客戶名稱、機台型號、機台序號、工程師 1 為必填。')
            if not self.files['sys']: raise ValueError('請選擇 sysdata.txt。')
            vals=[float(v.get()) for v in self.cpk]
            d=datetime.strptime(self.datev.get().strip(),'%Y/%m/%d').date(); data=parse_sysdata(self.files['sys'])
            has_ccd=bool(self.files['up'] or self.files['down']); tname='sample.xlsx' if has_ccd else 'sample_without_ccd.xlsx'; template=os.path.join(base_dir(),'templates',tname)
            default=f"{self.customer.get().strip()}-{self.model.get().strip()}-{self.serial.get().strip()}-{d.isoformat()}.xlsx"
            out=filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile=default,filetypes=[('Excel Workbook','*.xlsx')]);
            if not out:return
            fields={'customer':self.customer.get().strip(),'model':self.model.get().strip(),'serial':self.serial.get().strip(),'eng1':self.eng1.get().strip(),'eng2':self.eng2.get().strip(),'date':d}
            generate_report(template,out,fields,vals,data,self.files); messagebox.showinfo('完成',f'報告已產生：\n{out}')
        except Exception as e: messagebox.showerror('無法產出報告',str(e))
    def help(self):
        messagebox.showinfo('使用說明','1. 輸入客戶/機型/序號/工程師/日期。\n2. 選 CPK1 後會開啟 OCR/人工確認預覽。\n3. 選 sysdata，程式取最後 100 筆有效資料的前 6 欄。\n4. Camera Up/Down 任一有提供時自動使用 CCD 範本。\n5. 確認六個 CPK 後按「產出報告」。')
if __name__=='__main__': App().mainloop()
