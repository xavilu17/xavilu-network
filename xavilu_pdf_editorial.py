"""Editorial PDF renderer shared by generated documents and study materials."""
from io import BytesIO
from pathlib import Path
from html import escape
import re
import threading

_font_lock = threading.Lock()


def generar_pdf_editorial(titulo, contenido, logo_bytes=None, familia_fuente='Arial', fuente_personalizada=None):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, CondPageBreak, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image, Preformatted

    if familia_fuente == 'Angelina' and not fuente_personalizada:
        raise ValueError('La fuente Angelina requiere su archivo TTF para el diseño editorial.')
    with _font_lock:
        family = 'Georgia' if familia_fuente == 'Georgia' else 'Courier' if familia_fuente == 'Courier New' else 'Times' if familia_fuente == 'Times New Roman' else 'Arial'
        choices = {'Arial':('arial.ttf','arialbd.ttf'), 'Georgia':('georgia.ttf','georgiab.ttf'),
                   'Times':('times.ttf','timesbd.ttf'), 'Courier':('cour.ttf','courbd.ttf')}
        font_paths = [Path('C:/Windows/Fonts'), Path('/usr/share/fonts/truetype/dejavu')]
        regular = next((p/choices[family][0] for p in font_paths if (p/choices[family][0]).exists()), None)
        bold = next((p/choices[family][1] for p in font_paths if (p/choices[family][1]).exists()), None)
        if regular is None:
            regular = next((p/'DejaVuSans.ttf' for p in font_paths if (p/'DejaVuSans.ttf').exists()), None)
            bold = next((p/'DejaVuSans-Bold.ttf' for p in font_paths if (p/'DejaVuSans-Bold.ttf').exists()), None)
        font, font_bold = 'Helvetica', 'Helvetica-Bold'
        if regular:
            font, font_bold = 'XaviluEditorial'+family, 'XaviluEditorial'+family+'Bold'
            if font not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(font, str(regular)))
                pdfmetrics.registerFont(TTFont(font_bold, str(bold or regular)))
                pdfmetrics.registerFontFamily(font, normal=font, bold=font_bold, italic=font, boldItalic=font_bold)
        if fuente_personalizada:
            import hashlib
            font = 'XaviluCustom'+hashlib.sha256(fuente_personalizada).hexdigest()[:12]
            if font not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(font, BytesIO(fuente_personalizada)))
                pdfmetrics.registerFontFamily(font, normal=font, bold=font, italic=font, boldItalic=font)
            font_bold = font
    navy, teal = colors.HexColor('#132C43'), colors.HexColor('#087E8B')
    muted, pale = colors.HexColor('#546777'), colors.HexColor('#F0F5F8')
    styles = getSampleStyleSheet()
    body = ParagraphStyle('XBody', fontName=font, fontSize=9.6, leading=14.5, textColor=navy, spaceAfter=7, splitLongWords=True)
    title_style = ParagraphStyle('XTitle', parent=body, fontName=font_bold, fontSize=23, leading=28, spaceAfter=14, keepWithNext=True)
    headings = {1: ParagraphStyle('XH1', parent=body, fontName=font_bold, fontSize=15, leading=20, textColor=teal, spaceBefore=14, spaceAfter=7, keepWithNext=False),
                2: ParagraphStyle('XH2', parent=body, fontName=font_bold, fontSize=12, leading=17, spaceBefore=11, spaceAfter=6, keepWithNext=False)}
    cell = ParagraphStyle('XCell', parent=body, fontSize=8.2, leading=11, spaceAfter=0)
    header_cell = ParagraphStyle('XCellHead', parent=cell, fontName=font_bold, textColor=colors.white)
    code_style = ParagraphStyle('XCode', parent=body, fontName='Courier', fontSize=8, leading=11, backColor=pale, borderPadding=8)
    face = pdfmetrics.getFont(font).face
    supported = getattr(face, 'charWidths', None)
    def clean(text):
        text = str(text).replace('\t','    ')
        text = ''.join(c for c in text if c in '\n\r' or ord(c)>=32)
        if supported:
            text = ''.join(c if c in '\n\r' or ord(c) in supported else ' ' for c in text)
        return text
    def inline(text):
        text = escape(clean(text), quote=False)
        text = re.sub(r'\[([^\]]+)\]\((https?://[^\s)]+)\)', lambda m: '<link href="'+escape(m[2],quote=True)+'" color="#087E8B">'+m[1]+'</link>', text)
        text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
        text = re.sub(r'`([^`]+)`', r'<font name="Courier">\1</font>', text)
        return text.replace('\n','<br/>')
    buffer = BytesIO()
    page_width, page_height = A4
    usable = page_width - 40*mm
    doc = BaseDocTemplate(buffer, pagesize=A4, rightMargin=20*mm, leftMargin=20*mm, topMargin=25*mm, bottomMargin=22*mm, title=clean(titulo or 'Documento Xavilu'), author='Xavilu Network', pageCompression=1)
    story=[]
    if logo_bytes:
        image = Image(BytesIO(logo_bytes))
        image._restrictSize(42*mm, 19*mm)
        image.hAlign='LEFT'
        story.extend([image,Spacer(1,5*mm)])
    if titulo:
        story.append(Paragraph(inline(titulo),title_style))
    story.extend([HRFlowable(width='100%', thickness=1.2, color=teal), Spacer(1,5*mm)])
    lines = clean(contenido).splitlines()
    paragraph=[]
    def flush():
        if paragraph:
            story.append(Paragraph(inline(' '.join(paragraph)),body))
            paragraph.clear()
    i=0
    while i<len(lines):
        line=lines[i].rstrip()
        if line.strip().startswith('```'):
            flush();i+=1;code=[]
            while i<len(lines) and not lines[i].strip().startswith('```'):
                code.extend([lines[i][j:j+88] for j in range(0,max(1,len(lines[i])),88)])
                i+=1
            story.append(Preformatted('\n'.join(code),code_style));story.append(Spacer(1,3*mm));i+=1;continue
        if '|' in line and i+1<len(lines) and re.match(r'^\s*\|?\s*:?-{3,}',lines[i+1]):
            flush();rows=[]
            def cells(row):return [v.strip() for v in row.strip().strip('|').split('|')]
            header=cells(line);i+=2
            while i<len(lines) and '|' in lines[i] and lines[i].strip():
                rows.append(cells(lines[i]));i+=1
            count=max([len(header)]+[len(row) for row in rows]);header += ['']*(count-len(header))
            data=[[Paragraph(inline(v),header_cell) for v in header]]
            for row in rows:
                row += ['']*(count-len(row));data.append([Paragraph(inline(v),cell) for v in row])
            table=Table(data,colWidths=[usable/count]*count,repeatRows=1,hAlign='LEFT',splitByRow=1,splitInRow=1)
            table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),navy),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,pale]),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),('LINEBELOW',(0,0),(-1,0),.8,teal),('LINEBELOW',(0,1),(-1,-1),.3,colors.HexColor('#DCE5EA'))]))
            story.extend([table,Spacer(1,4*mm)]);continue
        heading = re.match(r'^(#{1,6})\s+(.+)',line)
        if heading:
            flush();story.append(CondPageBreak(75));story.append(Paragraph(inline(heading[2]),headings[min(len(heading[1]),2)]))
        elif re.match(r'^\s*[-*•]\s+',line):
            flush();text=re.sub(r'^\s*[-*•]\s+','',line);story.append(Paragraph('• '+inline(text),body))
        elif not line.strip():
            flush()
        else:
            paragraph.append(line.strip())
        i+=1
    flush()
    def page(canvas, document):
        canvas.saveState()
        canvas.setFont(font_bold,8)
        canvas.setFillColor(navy)
        canvas.drawString(20*mm,page_height-14*mm,'XAVILU NETWORK')
        canvas.setStrokeColor(teal)
        canvas.setLineWidth(.6)
        canvas.line(20*mm,17*mm,page_width-20*mm,17*mm)
        canvas.setFont(font,7.5)
        canvas.setFillColor(muted)
        canvas.drawString(20*mm,12*mm,'Documento generado en Xavilu')
        canvas.drawRightString(page_width-20*mm,12*mm,f'Página {document.page}')
        canvas.restoreState()
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="contenido")
    doc.addPageTemplates([PageTemplate(id="Xavilu", frames=[frame], onPageEnd=page)])
    doc.build(story)
    return buffer.getvalue()
