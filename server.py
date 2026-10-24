import io, os, re, shutil, subprocess, tempfile, zipfile
from pathlib import Path
from contextlib import asynccontextmanager
import fitz
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook, load_workbook
from pptx import Presentation
from pptx.util import Inches as PInches
from PIL import Image

ROOT=Path(__file__).resolve().parent
app=FastAPI(title='咚咚文件工具箱')
MAX=40*1024*1024
EXT={0:{'.doc','.docx'},1:{'.doc','.docx'},2:{'.xls','.xlsx'},3:{'.xls','.xlsx'},4:{'.xls','.xlsx'},5:{'.xls','.xlsx'},6:{'.ppt','.pptx'},**{i:{'.pdf'} for i in range(7,19)}}

def safe_name(s):return re.sub(r'[^\w.\-\u4e00-\u9fff]+','_',Path(s).name)[:100]
def lo(src,out,fmt):
    profile=out/'profile';profile.mkdir(exist_ok=True)
    cmd=['libreoffice',f'-env:UserInstallation=file://{profile}','--headless','--convert-to',fmt,'--outdir',str(out),str(src)]
    try:p=subprocess.run(cmd,capture_output=True,text=True,timeout=100)
    except subprocess.TimeoutExpired:raise ValueError('轉換逾時')
    result=out/(src.stem+'.'+fmt.split(':')[0])
    if not result.exists():raise ValueError('LibreOffice 轉換失敗：'+(p.stderr or p.stdout)[-350:])
    return result

def pdf_doc(src,password=''):
    d=fitz.open(src)
    if d.needs_pass and not d.authenticate(password):
        d.close();raise ValueError('PDF 已加密，請輸入正確密碼')
    return d

def render_pdf(src,out,ext='jpg',password=''):
    d=pdf_doc(src,password);paths=[]
    try:
        if len(d)>100:raise ValueError('最多處理 100 頁')
        for n,page in enumerate(d):
            pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
            p=out/f'{src.stem}_第{n+1}頁.{ext}'
            pix.save(str(p),output='jpeg' if ext=='jpg' else 'png');paths.append(p)
    finally:d.close()
    return paths

def bundle(paths,out,name):
    if len(paths)==1:return paths[0]
    dest=out/name
    with zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as z:
        for p in paths:z.write(p,p.name)
    return dest

def office_pdf(src,out):return lo(src,out,'pdf')

def pdf_to_word(src,out,password):
    d=pdf_doc(src,password);doc=Document()
    for i,page in enumerate(d):
        if i:doc.add_page_break()
        text=page.get_text('text')
        for line in text.splitlines():doc.add_paragraph(line)
    d.close();p=out/(src.stem+'.docx');doc.save(p);return p

def pdf_to_excel(src,out,password):
    d=pdf_doc(src,password);w=Workbook();w.remove(w.active)
    for i,page in enumerate(d):
        s=w.create_sheet(f'第{i+1}頁')
        tabs=page.find_tables()
        if tabs.tables:
            row=1
            for table in tabs.tables:
                for values in table.extract():
                    for j,v in enumerate(values,1):s.cell(row,j,str(v or '')[:32767])
                    row+=1
                row+=1
        else:
            for j,line in enumerate(page.get_text('text').splitlines()[:10000],1):s.cell(j,1,line[:32767])
    d.close();p=out/(src.stem+'.xlsx');w.save(p);return p

def pdf_to_ppt(src,out,password):
    d=pdf_doc(src,password);prs=Presentation();prs.slide_width=PInches(10);prs.slide_height=PInches(7.5)
    for page in d:
        slide=prs.slides.add_slide(prs.slide_layouts[6]);pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False)
        img=io.BytesIO(pix.tobytes('png'));w,h=page.rect.width,page.rect.height
        ratio=min(prs.slide_width/w,prs.slide_height/h);pw=int(w*ratio);ph=int(h*ratio)
        slide.shapes.add_picture(img,(prs.slide_width-pw)//2,(prs.slide_height-ph)//2,width=pw,height=ph)
    d.close();p=out/(src.stem+'.pptx');prs.save(p);return p

def split_excel(src,out,as_pdf):
    # LibreOffice can import legacy .xls, normalize it first.
    if src.suffix.lower()=='.xls':src=lo(src,out,'xlsx')
    wb=load_workbook(src,read_only=False,data_only=False)
    if len(wb.sheetnames)>100:raise ValueError('最多處理 100 張工作表')
    paths=[]
    for idx,sheet in enumerate(wb.worksheets,1):
        single=Workbook();target=single.active;target.title=sheet.title
        for row in sheet:
            for c in row:
                if c.value is not None:target.cell(c.row,c.column,c.value)
        for key,dim in sheet.column_dimensions.items():target.column_dimensions[key].width=dim.width
        for key,dim in sheet.row_dimensions.items():target.row_dimensions[key].height=dim.height
        target.sheet_format=sheet.sheet_format
        p=out/f'{safe_name(src.stem)}_{idx}_{safe_name(sheet.title)}.xlsx';single.save(p)
        if as_pdf:
            converted=office_pdf(p,out);paths.append(converted)
        else:paths.append(p)
    wb.close();return bundle(paths,out,src.stem+('_工作表PDF.zip' if as_pdf else '_工作表.zip'))

def process(index,files,out,option,password):
    src=files[0];stem=src.stem
    if index in (0,2,6):return office_pdf(src,out)
    if index in (1,3):return bundle(render_pdf(office_pdf(src,out),out,'jpg'),out,stem+'_圖片.zip')
    if index in (4,5):return split_excel(src,out,index==4)
    if index==7:return pdf_to_ppt(src,out,password)
    if index==12:return pdf_to_word(src,out,password)
    if index==13:return pdf_to_excel(src,out,password)
    if index in (14,15):return bundle(render_pdf(src,out,'jpg' if index==14 else 'png',password),out,stem+'_圖片.zip')
    if index==8:
        if len(files)<2:raise ValueError('至少需要兩個 PDF')
        result=fitz.open()
        for f in files:
            d=pdf_doc(f,password);result.insert_pdf(d);d.close()
        p=out/'合併文件.pdf';result.save(p,garbage=4,deflate=True);result.close();return p
    d=pdf_doc(src,password)
    try:
        if index==9:
            if not d.needs_pass:raise ValueError('這份 PDF 沒有密碼保護')
            p=out/(stem+'_解鎖.pdf');d.save(p,encryption=fitz.PDF_ENCRYPT_NONE,garbage=4);return p
        if index==10:
            angle=int(option or '90')
            if angle not in (90,180,270):raise ValueError('旋轉角度不正確')
            for page in d:page.set_rotation((page.rotation+angle)%360)
        elif index in (11,16):
            indices=parse_pages(option,len(d))
            if index==11:
                paths=[]
                for i in indices:
                    part=fitz.open();part.insert_pdf(d,from_page=i,to_page=i)
                    p=out/f'{stem}_第{i+1}頁.pdf';part.save(p);part.close();paths.append(p)
                return bundle(paths,out,stem+'_拆分.zip')
            if len(indices)==len(d):raise ValueError('不可刪除全部頁面')
            d.delete_pages(indices)
        elif index==17:
            for page in d:
                for widget in list(page.widgets() or []):
                    try:widget.update();page.delete_widget(widget)
                    except Exception:pass
                # rasterize each page to truly flatten visual content and annotations
            p=out/(stem+'_壓平.pdf');flat=fitz.open()
            for page in d:
                pix=page.get_pixmap(matrix=fitz.Matrix(1.5,1.5),alpha=False,annots=True)
                rect=page.rect;new=flat.new_page(width=rect.width,height=rect.height)
                new.insert_image(rect,stream=pix.tobytes('png'))
            flat.save(p,garbage=4,deflate=True);flat.close();return p
        elif index==18:pass # rewrite cross-reference, streams and object structure
        else:raise ValueError('未知工具')
        p=out/(stem+{10:'_旋轉',16:'_刪除頁面',18:'_修復'}.get(index,'_處理')+'.pdf')
        d.save(p,garbage=4,deflate=True,clean=True);return p
    finally:d.close()

def parse_pages(value,n):
    if not value.strip():raise ValueError('請輸入頁碼')
    indices=set()
    for token in value.split(','):
        m=re.fullmatch(r'\s*(\d+)\s*(?:-\s*(\d+)\s*)?',token)
        if not m:raise ValueError('頁碼格式應為 1-3,5')
        a=int(m[1]);b=int(m[2] or a)
        if a<1 or b>n or a>b:raise ValueError(f'頁碼必須在 1 至 {n} 之間')
        indices.update(range(a-1,b))
    return sorted(indices)

@app.post('/api/process')
async def api_process(tool:int=Form(...),option:str=Form(''),password:str=Form(''),files:list[UploadFile]=File(...)):
    if tool not in EXT:raise HTTPException(400,'不支援的工具')
    if len(files)>30:raise HTTPException(400,'最多 30 個檔案')
    if tool != 8 and len(files) != 1:raise HTTPException(400,'此工具一次只能處理一個檔案')
    with tempfile.TemporaryDirectory(prefix='dongdong_') as td:
        out=Path(td);paths=[];total=0
        for i,f in enumerate(files):
            ext=Path(f.filename or '').suffix.lower()
            if ext not in EXT[tool]:raise HTTPException(400,'檔案格式不符')
            data=await f.read(MAX+1)
            if len(data)>MAX:raise HTTPException(413,'單一檔案不得超過 40 MB')
            total+=len(data)
            if total>120*1024*1024:raise HTTPException(413,'單次上傳總量不得超過 120 MB')
            p=out/(str(i)+'_'+safe_name(f.filename or 'file'+ext));p.write_bytes(data);paths.append(p)
        try:result=process(tool,paths,out,option,password)
        except (ValueError,fitz.FileDataError,fitz.FileNotFoundError) as e:raise HTTPException(422,str(e))
        except Exception as e:raise HTTPException(500,'處理失敗，請確認來源檔案完整及格式受支援')
        content=result.read_bytes();name=result.name
        from urllib.parse import quote
        return Response(content,media_type='application/zip' if result.suffix=='.zip' else 'application/octet-stream',headers={'Content-Disposition':"attachment; filename*=UTF-8''"+quote(name),'X-Filename':quote(name)})

@app.get('/api/health')
def health():return {'ok':True,'tools':19,'libreoffice':bool(shutil.which('libreoffice'))}
@app.get('/')
def homepage():return FileResponse(ROOT/'index.html',media_type='text/html')
