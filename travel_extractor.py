#!/usr/bin/env python3
"""Travel document extractor v4 (6-column output): ZIP/folder/single file -> normalized Excel.
Supports PDF, PNG, JPG/JPEG, WEBP; local OCR/rules and optional OpenAI-compatible AI vision/text extraction.
"""
from __future__ import annotations
import argparse, base64, io, json, os, re, shutil, tempfile, zipfile, sys, traceback
from pathlib import Path
from typing import Any
import fitz, pandas as pd, pytesseract, requests
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SUPPORTED={'.pdf','.png','.jpg','.jpeg','.webp','.tif','.tiff','.bmp'}
COLUMNS=['Vendor','Amount Charged','Commission','CGST','SGST','TDS','Passenger Name','Travel Date',
         'From City / Station','To City / Station','Arrival Date','Departure Date','Airline Name',
         'Hotel City','Visa Country','Insurance Plan','PNR / Booking Confirmation','Booking Date',
         'Vendor Invoice Number','Type of Document','Currency','Source File','Extraction Method','Review Needed','Hotel Name']
DOC_TYPES=['Air Ticket','Hotel Booking','Visa','Insurance','Train Ticket','Bus Ticket','Other']

def configure_tesseract():
    """Find Tesseract automatically, especially on Windows."""
    env=os.getenv('TESSERACT_CMD','').strip()
    candidates=[env, shutil.which('tesseract'),
                r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe',
                r'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe']
    for c in candidates:
        if c and Path(c).exists():
            pytesseract.pytesseract.tesseract_cmd=str(c); return str(c)
    return None

TESSERACT_PATH=configure_tesseract()

def pdf_text(path):
    with fitz.open(path) as d: return '\n'.join(p.get_text('text') for p in d)

def render_pdf_pages(path,dpi=250):
    with fitz.open(path) as d:
        for p in d:
            pix=p.get_pixmap(matrix=fitz.Matrix(dpi/72,dpi/72),alpha=False)
            yield Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB')

def prep_images(img):
    """Multiple OCR-friendly variants; fixes phone photos, low contrast and handwriting as far as Tesseract permits."""
    img=ImageOps.exif_transpose(img).convert('RGB')
    # upscale small/medium photos
    if max(img.size)<2600:
        scale=min(3.0,2600/max(img.size)); img=img.resize((int(img.width*scale),int(img.height*scale)))
    gray=ImageOps.grayscale(img)
    gray=ImageEnhance.Contrast(gray).enhance(1.8).filter(ImageFilter.SHARPEN)
    auto=ImageOps.autocontrast(gray)
    bw=auto.point(lambda x: 255 if x>165 else 0)
    return [img,auto,bw]

def ocr_image(img):
    ocr_image.last_variants=[]
    if not TESSERACT_PATH:
        raise RuntimeError('Tesseract OCR was not found. Install Tesseract or set TESSERACT_CMD in .env / environment. See README.md.')
    texts=[]
    for v in prep_images(img):
        for psm in (6,11):
            try: texts.append(pytesseract.image_to_string(v,config=f'--oem 3 --psm {psm}'))
            except Exception: pass
    ocr_image.last_variants=list(texts)
    return max(texts,key=len,default='')

def extract_local_text(path):
    if path.suffix.lower()=='.pdf':
        txt=pdf_text(path)
        # OCR image-only/scanned PDFs; also append OCR when embedded text is suspiciously short
        if len(txt.strip())>=180: return txt,'pdf-text'
        ocr='\n'.join(ocr_image(x) for x in render_pdf_pages(path))
        return (txt+'\n'+ocr).strip(),'pdf-ocr'
    try:
        with Image.open(path) as im: return ocr_image(im.copy()),'image-ocr'
    except Exception as e: raise RuntimeError(f'Cannot open image: {e}')

def first(patterns,text):
    for p in patterns:
        m=re.search(p,text,re.I|re.M|re.S)
        if m: return re.sub(r'\s+',' ',m.group(1)).strip(' :-')
    return ''
def num(s):
    if not s:return None
    try:return float(re.sub(r'[^0-9.\-]','',s.replace(',','')))
    except:return None

def classify(t):
    x=t.lower()
    # Strong, transaction-specific signals are checked before incidental words in T&C.
    if any(k in x for k in ['train no','irctc','railway','boarding station']): return 'Train Ticket'
    if any(k in x for k in ['bus ticket','boarding point','dropping point','redbus','bus charges','bus operator']): return 'Bus Ticket'
    if 'no. of nights' in x and 'hotel' in x and ('check in' in x or 'check-in' in x): return 'Hotel Booking'
    air_score=sum(k in x for k in ['pnr','departure flight','flight fare','flight details','airways','airline','airport','ticket no','airfare'])
    if air_score >= 2: return 'Air Ticket'
    if any(k in x for k in ['hotel','check-in','check in','check-out','room type']) and any(k in x for k in ['room','guest','stay','property']): return 'Hotel Booking'
    if any(k in x for k in ['insurance','travel assist','policy no','certificate no']) and not air_score: return 'Insurance'
    if any(k in x for k in ['cost of visa','visa type','visa application','destination:','ukvi','uk visas','vfs.global','vfs global']) and 'visa' in x: return 'Visa'
    if any(k in x for k in ['travel agent','travels','class-','class -']) and any(k in x for k in [' to ','dated','amount','total']): return 'Air Ticket'
    if air_score: return 'Air Ticket'
    return 'Other'

def vendor(text):
    mp={'travomint':'Travomint','makemytrip':'MakeMyTrip','cleartrip':'Cleartrip','atlys':'Atlys','indigo':'IndiGo',
        'asego':'Asego Global Assistance','expedia':'Expedia','agarwal travels':'Agarwal Travels','ibis':'ibis','british airways':'British Airways'}
    low=text.lower()
    for k,v in mp.items():
        if k in low:return v
    return first([r'(?:invoice from|vendor|company)\s*[:\-]?\s*([A-Za-z][A-Za-z &.]{2,50})'],text)

def money_field(label,text):
    m=re.search(rf'{label}\s*@\s*[\d.]+\s*%\s*\(([\-\d,]+(?:\.\d+)?)\)',text,re.I)     # "CGST @ 9.00% (135.00)"
    if m: return num(m.group(1))
    return num(first([rf'{label}\s*(?:\(\+\)|\(\-\))?\s*[:₹INR ]*([\-\d,]+(?:\.\d+)?)(?!\s*%)',rf'{label}[^\d\-]{{0,20}}([\-\d,]+(?:\.\d+)?)(?!\s*%)'],text))

def date_field(labels,text):
    labs='|'.join(labels)
    return first([rf'(?:{labs})\s*[:\-]?\s*((?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)?[,]?\s*\d{{1,2}}[ /\-](?:[A-Za-z]{{3,9}}|\d{{1,2}})[ /\-]\d{{2,4}}(?:\s+\d{{1,2}}:\d{{2}}(?:\s*[AP]M)?)?)',
                  rf'(?:{labs})\s*[:\-]?\s*(\d{{4}}[-/]\d{{1,2}}[-/]\d{{1,2}}(?:\s+\d{{1,2}}:\d{{2}})?)'],text)

BAD_NAME_WORDS={'information','info','details','detail','name','names','type','list','summary','adult','child','infant','adt','chd','inf',
    'sector','seat','services','service','number','no','contact','email','ticketing','ticket','pnr','booking','base','fare','tax','corporate',
    'value','economy','refundable','flexi','saver','class','gender','dob','age','mobile','phone','status','passenger','passengers','mon','tue','wed','thu','fri','sat','sun','jan','feb','apr','jun','jul','aug','sep','oct','nov','dec'}
TITLE=r'(?:MR|Mr|MS|Ms|MRS|Mrs|MSTR|Mstr|MISS|Miss|DR|Dr)'
MONTHS=r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'
DATE_RX=r'\d{1,2}[ \-/]'+MONTHS+r'[a-z]*[ \-/]\d{4}'

def clean_name(n):
    n=re.sub(r'\s+',' ',n or '').strip(' :-|.')
    out=[]
    for w in n.split(' '):
        if w.lower().strip('.|,') in BAD_NAME_WORDS: break   # stop at Adult / ADT / Seat / Ticketing ...
        out.append(w)
    n=' '.join(out).strip(' :-|.,')
    return n if len(n)>=3 else ''

def get_passenger(text):
    """All passenger names (multiple passengers -> 'A, B'). Never returns header words like Information/Ticketing."""
    m=re.search(r'Passengers? Details.*?(?:QR Code|Ticket No)\s+(?:Mr|Ms|Mrs)?\.?\s*([A-Za-z]{2,25}\s+[A-Za-z]{2,25})',text,re.I|re.S)
    if m:
        n=clean_name(m.group(1))
        if n: return n
    # labelled single name - a colon/dash is REQUIRED so table headers are not mistaken for names
    for p in [r'Passenger\s+Name\s*[:\-]\s*(?:'+TITLE+r')?\.?\s*([A-Za-z][A-Za-z .]{2,45})',
              r'Passenger\s*[:\-]\s*(?:'+TITLE+r')?\.?\s*([A-Za-z][A-Za-z .]{2,45})',
              r'Primary Guest\)?\s*(?:'+TITLE+r')?\.?\s*([A-Za-z][A-Za-z .]{2,45})',
              r'Name of Traveller\s*[:\-]?\s*([A-Za-z][A-Za-z .]{2,45})',
              r'Customer Name\s*\n\s*([A-Za-z][A-Za-z .]{2,45})',                      # MakeMyTrip tax invoice
              r'Pax\s*[:\-]\s*([A-Za-z][A-Za-z .]{2,45})(?=\s*[|\n]\s*PNR)',          # Cleartrip eReceipt "Pax: NAME | PNR:"
              r'Name\s*\n+\s*([A-Z][A-Za-z .]{2,45}(?:\n[A-Za-z .]{1,30})?)\s*\n+\s*(?:UAN|Visa)\b']:  # VFS / UKVI "Name\n<NAME>\nVisa|UAN" (name may wrap)
        for m in re.finditer(p,text,re.I):
            n=clean_name(m.group(1))
            if n: return n
    m=re.search(r'^Name\s*[:\-]\s*([A-Za-z][A-Za-z .]{2,45})\s*$',text,re.M)            # "Name: NAME" own line (UKVI email)
    if m:
        n=clean_name(m.group(1))
        if n: return n
    n=hotel_guest(text)
    if n: return n
    # ALL-CAPS "LASTNAME, First Middle" passenger-list lines (e.g. visa agency invoices)
    names=[]
    for m in re.finditer(r'^([A-Z][A-Z]+),\s*([A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+)?)\s*$',text,re.M):
        last,first=m.group(1).title(),m.group(2).title()
        n=f'{first} {last}'
        if n.lower() not in [x.lower() for x in names]: names.append(n)
    if names: return ', '.join(names)
    # US visa scheduling: primary applicant + each additional "Applicant Name:"
    m=re.search(r'Appointment\(s\)\s+Made\s+By\s*:\s*\n\s*([A-Za-z][A-Za-z .]{2,45})',text,re.I)
    if m:
        names=[clean_name(m.group(1))]
        for a in re.finditer(r'Applicant\s+Name\s*:\s*\n\s*([A-Za-z][A-Za-z .]{2,45})',text,re.I):
            n=clean_name(a.group(1))
            if n and n.lower() not in [x.lower() for x in names]: names.append(n)
        names=[n for n in names if n]
        if names: return ', '.join(names)
    # titled names anywhere: "MR Prakash Poddar", "Mr Sunil\nAgarwal ADT", "Mrs Sonali Agarwal"
    names=[]
    for m in re.finditer(r'\b'+TITLE+r'\.?[ \t]+([A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+){0,2})',text):
        n=clean_name(m.group(1))
        if n and n.lower() not in [x.lower() for x in names]: names.append(n)
    return ', '.join(names)

def get_pnr(text):
    for m in re.finditer(r'(?i:PNR)(?:\s*(?i:NUMBER|No\.?)|\s*/\s*(?i:Booking)\s*(?i:Ref(?:erence)?|ID|No\.?))?\s*[-:#]?\s*([A-Z0-9]{5,10})\b',text):
        v=m.group(1)
        if v in ('NUMBER','DETAILS','BOOKING','STATUS','TICKET') or (not re.search(r'\d',v) and len(v)!=6): continue
        return v
    m=re.search(r'\b([A-Z0-9]{6})\s*/\s*\1\b',text)          # table cell like  U5332P/U5332P
    if m: return m.group(1)
    m=re.search(r'(?i:Confirmation)(?:\s*(?i:number|no\.?|code|#))?\s*[:#\-]?\s*([A-Z0-9\-]{5,20})\b',text)
    if m and re.search(r'\d',m.group(1)): return m.group(1)
    v=(grid_or_cell(text,'bookref') or '').split(' ')[0]
    if re.fullmatch(r'[A-Z0-9\-]{5,20}',v or '') and re.search(r'\d',v): return v
    m=re.search(r'\bREZ[A-Z0-9]{5,14}\b',text)
    if m: return m.group(0)
    m=re.search(r'booking reference(?: number)? is\s*([A-Z0-9\-]+)',text,re.I)
    if m: return m.group(1)
    m=re.search(r'Itinerary\s*(?:number|no\.?)\s*[:#\-]?\s*(\d{8,20})',text,re.I)
    if m: return m.group(1)
    m=re.search(r'AIRLINE\s*PNR\s*\nTICKET\s*NO\.?\s*\n(?:[^\n]*\n){0,6}?\s*(?=[A-Z0-9]{6}\b)(?=[A-Z0-9]*[A-Z])[A-Z0-9]{6}\b',text,re.I)
    if m: return re.search(r'[A-Z0-9]{6}$',m.group(0)).group(0)
    m=re.search(r'Reference\s*Number\s*\n\s*([A-Z0-9]{6,15})',text,re.I)
    if m: return m.group(1)
    m=re.search(r'DS-160\s*Confirmation\s*Number\s*:?\s*\n?\s*([A-Z0-9]{5,15})',text,re.I)
    if m: return m.group(1)
    m=re.search(r'Reference(?:\s*(?:Number|No\.?))?\s*:?\s*\n*\s*([A-Za-z0-9/_\-]{4,25})',text,re.I)
    if m and re.search(r'\d',m.group(1)): return m.group(1).strip('/')
    m=re.search(r'Booking\s*ID\s*:?\s*\n?\s*([A-Z0-9]{6,25})',text,re.I)
    if m: return m.group(1)
    return ''

AIRLINE_NAMES={'british airways':'British Airways','air india express':'Air India Express','air india':'Air India','indigo':'IndiGo','vistara':'Vistara',
    'akasa air':'Akasa Air','emirates':'Emirates','qatar airways':'Qatar Airways','etihad airways':'Etihad Airways','lufthansa':'Lufthansa',
    'spicejet':'SpiceJet','airasia':'AirAsia','singapore airlines':'Singapore Airlines','thai airways':'Thai Airways','turkish airlines':'Turkish Airlines'}
AIRLINE_CODES={'6E':'IndiGo','IX':'Air India Express','AI':'Air India','UK':'Vistara','QP':'Akasa Air','SG':'SpiceJet','I5':'AirAsia India','G8':'Go First',
    'EK':'Emirates','QR':'Qatar Airways','EY':'Etihad Airways','LH':'Lufthansa','BA':'British Airways','SQ':'Singapore Airlines','TG':'Thai Airways','TK':'Turkish Airlines'}

def get_airline(text):
    names='|'.join(re.escape(k) for k in AIRLINE_NAMES)
    m=re.search(r'\b('+names+r')\b',text,re.I)
    if m: return AIRLINE_NAMES[m.group(1).lower()]
    for p in [r'\(([A-Z0-9]{2})\s*[/\-]?\s*\d{2,4}\)', r'\b('+'|'.join(AIRLINE_CODES)+r')[ \-/]?\d{3,4}\b']:   # (IX/1547), 6E 536
        for m in re.finditer(p,text):
            if m.group(1) in AIRLINE_CODES: return AIRLINE_CODES[m.group(1)]
    return ''

def find_travel_date(text,route,booking):
    """Travel date when it has no label: after 'Departing', else right after the route, else first date that is not the booking/invoice date."""
    m=re.search(r'Departing.{0,600}?('+DATE_RX+')',text,re.S|re.I)
    if m: return m.group(1)
    if route:
        m=re.search(DATE_RX,text[route.end():route.end()+250],re.I)
        if m: return m.group(0)
    bd=re.search(DATE_RX,booking or '',re.I); bd=bd.group(0).lower() if bd else None
    for m in re.finditer(DATE_RX,text,re.I):
        if m.group(0).lower()!=bd: return m.group(0)
    return ''

# ---------------- Hotel (Expedia-style) helpers - only used for Hotel Booking documents ----------------
def hotel_name(text):
    m=re.search(r'\d+[\-\s]Nights?\s+Stay\s*\n(.+?)\n\s*Booking\s+Id',text,re.I|re.S)
    if m:
        n=re.sub(r'\s+',' ',m.group(1)).strip(' :-|')
        if n: return n
    m=re.search(r'Stay at[ \t]+([^\n]+)',text,re.I)
    if m:
        n=re.sub(r'[ \t]+\S{0,2}[ \t]*Edit\b.*$','',m.group(1)).strip(' :-|')
        toks=n.split()
        # drop trailing OCR junk from the pencil icon / Edit link: single characters or tokens with odd symbols
        while toks and (len(toks[-1])<=1 or re.search(r"[^\w\-'.,&]",toks[-1]) or toks[-1].lower()=='edit'): toks.pop()
        n=' '.join(toks).strip(' :-|')
        if n: return n
    n=grid_fields(text).get('hotel') or ''
    if n: return n
    m=re.search(r'Hotel(?: Name)?[ \t]*[:\-][ \t]*([^\n]+)',text,re.I)
    if m:
        v=m.group(1).strip(' :-|')
        if v and not any(re.match(l+r'\b',v,re.I) for k,l,c in GRID_LABELS): return v
    return cell_value(text,'hotel')

def hotel_guest(text):
    m=re.search(r'Room\s*\d+\s*:\s*([A-Z][a-z]+(?:[ \t]+[A-Z][a-z]+){1,3})\b.{0,12}?\d+\s*adults?',text,re.S)
    if m:
        n=clean_name(m.group(1))
        if n: return n
    m=re.search(r'Primary\s+travell?er(.{0,200})',text,re.I|re.S)
    if m:
        for c in re.finditer(r'\b([A-Z][a-z]+(?:[ \t]+[A-Z][a-z]+){1,2})\b',m.group(1)):
            n=clean_name(c.group(1))
            if n: return n
    n=re.sub(r'^(?:'+TITLE+r')\.?\s+','',grid_or_cell(text,'leader'))
    return clean_name(n)

MONEY=r'[\d,]*\d[.,]\d{2}'          # 4,499.10   or OCR'd comma-decimal 224,96

def _amt(s):
    s=s.strip()
    if '.' not in s and re.search(r',\d{2}$',s): s=s[:-3].replace(',','')+'.'+s[-2:]   # 224,96 -> 224.96
    return num(s)

def _variants(s):
    """[as read, with first digit dropped] - the 2nd covers a currency symbol (Rs) mis-read as a digit."""
    v=[_amt(s)]
    t=s.strip(); d=re.sub(r'^\d','',t,count=1).lstrip(',') if (re.match(r'\d',t) and len(re.sub(r'\D','',t))>3) else None
    v.append(_amt(d) if d else None)
    return v

def hotel_total(text):
    tot=[m.group(1) for m in re.finditer(r'\bTotal\s+for\s+\d+\s+nights?\b\D{0,4}?('+MONEY+')',text,re.I)]
    tot+=[m.group(1) for m in re.finditer(r'(?<!Sub )(?<!Sub-)\bTotal\b(?!\s+(?:for|earnings?|trip))\D{0,4}?('+MONEY+')',text,re.I)]
    room=re.search(r'\bRoom\b\D{0,4}?('+MONEY+')',text,re.I)
    tax=re.search(r'\bMiscellaneous\s+Tax\D{0,4}?('+MONEY+')',text,re.I)
    if room and tax:
        # Room + Tax must equal Total. Each number may independently have a mis-read currency symbol glued on,
        # so try every combination; prefer "as read", otherwise the smallest matching total.
        rv=_variants(room.group(1)); xv=_variants(tax.group(1))
        pools=[_variants(t) for t in tot] or [[None,None]]
        allnums=[_variants(m.group(0)) for m in re.finditer(MONEY,text)]
        for pool in (pools,allnums):
            hits=[]
            for tv in pool:
                for i in (0,1):
                    for j in (0,1):
                        for k in (0,1):
                            if None in (tv[i],rv[j],xv[k]): continue
                            if abs(rv[j]+xv[k]-tv[i])<0.02: hits.append((i+j+k,tv[i]))
            if hits:
                hits.sort(key=lambda h:(h[0]>0,h[1]))       # no correction first; else smallest total
                return hits[0][1]
    return _amt(tot[0]) if tot else None

def agent_earnings(text,total=None):
    m=re.search(r'Total\s+earnings?\D{0,40}?('+MONEY+')',text,re.I)
    if not m: return None
    s=m.group(1); v=_amt(s)
    if total and v and v>0.3*total:            # glued currency symbol -> drop first digit
        alt=_variants(s)[1]
        if alt: v=alt
    return v

# ---------------- "Label row / value row" grids (Rezlive-style tax invoices) ----------------
INDIAN_STATES=('Andhra Pradesh|Arunachal Pradesh|Assam|Bihar|Chhattisgarh|Goa|Gujarat|Haryana|Himachal Pradesh|Jharkhand|Karnataka|Kerala|'
    'Madhya Pradesh|Maharashtra|Manipur|Meghalaya|Mizoram|Nagaland|Odisha|Punjab|Rajasthan|Sikkim|Tamil Nadu|Telangana|Tripura|Uttar Pradesh|'
    'Uttarakhand|West Bengal|Delhi|Jammu and Kashmir|Ladakh|Chandigarh|Puducherry|Lakshadweep|Andaman and Nicobar Islands')
_TEXT='TEXT'
GRID_LABELS=[  # key, label phrase, shape of the value
    ('tel',r'TEL\.?\s*NO',r'\+?[\d\-]{8,}'),
    ('bookref',r'BOOKING\s*REF(?:ERENCE)?(?:\s*NO\.?)?',r'\S+'),
    ('voucherdate',r'VOUCHER\s*DATE',DATE_RX),
    ('voucherid',r'VOUCHER\s*ID',r'\S+'),
    ('checkin',r'CHECK[\s\-]*IN(?:\s*/\s*SUPPLY\s*DATE)?',DATE_RX),
    ('checkout',r'CHECK[\s\-]*OUT',DATE_RX),
    ('bookingdate',r'BOOKING\s*DATE',DATE_RX),
    ('agentcode',r'AGENT\s*CODE',r'\S+'),
    ('agentref',r'AGENT\s*REF\.?\s*NO',r'\S+'),
    ('leader',r'LEADER',_TEXT),
    ('nights',r'NO\.?\s*OF\s*NIGHTS',r'\d{1,3}'),
    ('agentname',r'AGENT\s*NAME',_TEXT),
    ('invoicedate',r'INVOICE\s*DATE',DATE_RX),
    ('state',r'STATE\s*/\s*PLACE\s*OF\s*SUPPLY',r'(?:'+INDIAN_STATES+r'|\S+)'),
    ('hotel',r'HOTEL',_TEXT),
    ('destination',r'DESTINATION',_TEXT),
    ('statecode',r'STATE\s*CODE',r'\d{1,3}'),
    ('gst',r'GST\s*NO\.?',r'\S+'),
    ('service',r'SERVICE\s*TYPE',_TEXT),
    ('sac',r'SAC',r'\d{4,8}'),
]

def split_hotel_dest(blob):
    """Row-wise layout gives 'Hotel name Destination' in one string; drop the trailing destination."""
    toks=blob.split()
    if len(toks)<2: return blob.strip()
    for k in (2,1):
        if len(toks)>k:
            dest=' '.join(toks[-k:]).lower(); rest=' '.join(toks[:-k]).lower()
            if dest in rest: return ' '.join(toks[:-k])
    return ' '.join(toks[:-1])

def grid_fields(text):
    """Rows of labels followed by a row of values ->  {key: value}. Empty dict when the document has no such grid."""
    out={}; lines=[l.strip() for l in text.splitlines()]
    for i,l in enumerate(lines):
        found=[]
        for key,lab,cons in GRID_LABELS:
            m=re.search(r'(?<![A-Za-z])'+lab+r'(?![A-Za-z])',l,re.I)
            if m: found.append((m.start(),m.end(),key,cons))
        found.sort()
        # remove overlapping hits (e.g. 'CHECK IN' inside a longer label)
        clean=[]
        for f in found:
            if clean and f[0]<clean[-1][1]: continue
            clean.append(f)
        if len(clean)<2: continue
        j=i+1
        while j<len(lines) and not lines[j]: j+=1
        if j>=len(lines): continue
        keys=[f[2] for f in clean]; conss=[f[3] for f in clean]
        merged_hotel=False
        if 'hotel' in keys and 'destination' in keys and keys.index('destination')==keys.index('hotel')+1:
            d=keys.index('destination'); keys.pop(d); conss.pop(d); merged_hotel=True
        parts=[]
        for n,c in enumerate(conss):
            last=n==len(conss)-1
            parts.append(r'(.+)' if (c==_TEXT and last) else (r'(.+?)' if c==_TEXT else '('+c+')'))
        m=re.match(r'^\W*'+r'\s+'.join(parts)+r'\s*$',lines[j],re.I)
        if not m: continue
        for k,v in zip(keys,m.groups()):
            v=v.strip(' :;|')
            if k=='hotel' and merged_hotel: v=split_hotel_dest(v)
            out.setdefault(k,v)
    return out

def cell_value(text,key):
    """Layout where each label is followed by its own value (label on one line, value on the next / same line)."""
    lab=[l for k,l,c in GRID_LABELS if k==key]
    if not lab: return ''
    m=re.search(r'(?<![A-Za-z])'+lab[0]+r'\s*[:;.\u00ab]?[ \t]*\n?[ \t]*([^\n]+)',text,re.I)
    if not m: return ''
    v=m.group(1).strip(' :;|')
    for k,l,c in GRID_LABELS:                       # value must not itself be another label
        if re.match(l+r'\b',v,re.I): return ''
    return v

def grid_or_cell(text,key):
    return grid_fields(text).get(key) or cell_value(text,key)

def vfs_grand_total_sum(text):
    vals=[_amt(x) for x in re.findall(r'Grand\s+Total\s*:?\s*INR\s*('+MONEY+')',text,re.I)]
    vals=[v for v in vals if v]
    return round(sum(vals),2) if len(vals)>=2 else None

def us_visa_fee_sum(text):
    """US visa scheduling 'MRV FEE PAYMENTS': one 'Amount:\n<INR> (<USD> USD)' block per applicant - sum them all."""
    if 'mrv fee payments' not in text.lower(): return None
    vals=[num(x) for x in re.findall(r'Amount\s*:?\s*\n\s*([\d,]+)\s*\(\s*[\d.]+\s*USD\)',text,re.I)]
    vals=[v for v in vals if v]
    return round(sum(vals),2) if vals else None

def total_amount_label(text):
    m=re.search(r'(?<!SUB )(?<!SUB-)\bTOTAL\s+AMOUNT\b(?!\s+IN\b)\s*[:;.\-]?\s*(?:INR\s*)?('+MONEY+')',text,re.I)
    if m: return _amt(m.group(1))
    m=re.search('('+MONEY+r')\s*\n\s*INR\s+[A-Za-z ]+\n\s*Net\s+Total',text,re.I)        # value BEFORE the label (Prominent Visa Services)
    if m: return _amt(m.group(1))
    m=re.search(r'Net\s+Total\s*(?:\(INR\))?\s*[:;.\-]?\s*(?:INR\s*)?('+MONEY+')',text,re.I)
    if m: return _amt(m.group(1))
    m=re.search(r'Paid\s+INR\s*([\d,]+(?:\.\d+)?)\s+with',text,re.I)                          # MakeMyTrip hotel voucher: "Paid INR 11,335 with ..."
    if m: return num(m.group(1))
    vals=[num(x) for x in re.findall(r'Net amount charged[^\n]*\n\s*([\d,]+(?:\.\d+)?)',text,re.I)]  # Cleartrip eReceipt: last "Net amount charged" wins
    if vals: return vals[-1]
    vals=[_amt(x) for x in re.findall('('+MONEY+r')\s*INR\b',text)]      # label lost in OCR: biggest '<amount> INR' is the total
    vals=[v for v in vals if v]
    return max(vals) if vals else None

def hotel_dates(text,booking):
    """(check-in, check-out): from a label grid if present, else independent 'Check-in'/'Check-out' searches, else the first two dates after 'Check-in'."""
    g=grid_fields(text)
    if g.get('checkin'): return g['checkin'],g.get('checkout','')
    ci=re.search(r'Check[\s\-]*in\b.{0,60}?('+DATE_RX+')',text,re.I|re.S)
    co=re.search(r'Check[\s\-]*out\b.{0,60}?('+DATE_RX+')',text,re.I|re.S)
    ci_v=ci.group(1) if ci else ''; co_v=co.group(1) if co else ''
    if ci_v and co_v and co_v.lower()==ci_v.lower():
        # "Check-in Check-out" on one line, both labels land on the same first date -> the real
        # check-out date is the next date that follows it (e.g. "... Wed, 12 Aug 2026 Thu, 13 Aug 2026")
        dates=re.findall(DATE_RX,text[ci.end():ci.end()+150],re.I)
        nxt=next((d for d in dates if d.lower()!=ci_v.lower()),None)
        if nxt: co_v=nxt
    if ci_v or co_v: return ci_v,co_v
    i=text.lower().find('check-in')
    if i<0:
        m=re.search(r'check[\s\-]*in',text,re.I); i=m.start() if m else -1
    seg=text[i:] if i>=0 else text
    bd=re.search(DATE_RX,booking or '',re.I); bd=bd.group(0).lower() if bd else None
    out=[]
    for m in re.finditer(DATE_RX,seg,re.I):
        d=m.group(0)
        if d.lower()!=bd and d not in out: out.append(d)
        if len(out)==2: break
    return (out[0] if out else ''),(out[1] if len(out)>1 else '')

def cleartrip_ticket_total(text):
    """Cleartrip 'ticket' PDFs: the fare numbers are extracted as a standalone block, in the same
    order as the (label-less, further down) FARE BREAKUP rows. The last number in that block is the Total Fare."""
    if 'FARE BREAKUP' not in text or 'Total Fare' not in text: return None
    lines=text.splitlines()
    best=[]
    run=[]
    for l in lines:
        l=l.strip()
        if re.fullmatch(r'-?[\d,]+(?:\.\d+)?',l or ''):
            run.append(l)
        else:
            if len(run)>len(best): best=run
            run=[]
    if len(run)>len(best): best=run
    return _amt(best[-1]) if len(best)>=2 else None

def visa_details(text):
    """Visa class/category (e.g. 'B1/B2 - Business/Tourism', 'Visitor visa'). Best-effort: different visa
    agencies/portals label this differently, so several shapes are tried in priority order."""
    cls=None; cat=None
    m=re.search(r'Visa\s*Class\s*:?\s*\n?\s*([A-Za-z0-9/]{2,15})',text,re.I)
    if m: cls=m.group(1).strip()
    m=re.search(r'Visa\s*Category\s*:?\s*\n\s*([A-Za-z][A-Za-z/ ]{2,40})',text,re.I)
    if m: cat=m.group(1).strip()
    if not cat:
        m=re.search(r'Visa\s*Type\s*:?\s*\n\s*([A-Za-z][A-Za-z/ ]{2,40})',text,re.I)
        if m and 'paid by pax' not in m.group(1).lower(): cat=m.group(1).strip()
    if not cat:
        m=re.search(r'\bVisa\b\s*:?\s*\n*\s*([A-Za-z][A-Za-z ]{2,30}?)\s*(?=\n|$)',text,re.I)
        if m and (m.group(1).strip().upper()==m.group(1).strip() or 'application' in m.group(1).lower()):
            m=None                                                              # reject shouting company-name text, e.g. "PROMINENT VISA SERVICES"
        if m and 'application' not in m.group(1).lower(): cat=m.group(1).strip()
    if not cls and not cat:
        m=re.search(r'Visa\s+Information\s*\n\s*[A-Z][A-Za-z ]*\n\s*([A-Za-z0-9/]{2,10})\s*\n\s*([^\n]+(?:\n[^\n]+)?)',text,re.I)
        if m: cls=m.group(1).strip(); cat=re.sub(r'\s+',' ',m.group(2)).strip()
    if not cls and not cat:
        m=re.search(r'\bCountry\s*:?\s*\n?\s*([A-Za-z][A-Za-z ]{2,20})',text,re.I)     # last resort: at least name the country
        if m: cat=f'Visa ({m.group(1).strip()})'
    return ' - '.join(dict.fromkeys(p for p in (cls,cat) if p))

def local_parse(text,filename):
    dtype=classify(text); low=text.lower()
    route=(re.search(r'\b([A-Z]{3})\s*[-–→]\s*([A-Z]{3})\b',text)
        or re.search(r'from\s+[A-Za-z ]+\(([A-Z]{3})\)\s+to\s+[A-Za-z ]+\n?\(([A-Z]{3})\)',text,re.I)
        or re.search(r'\b([A-Z]{3})\s+\d{1,2}:\d{2}\b[\s\S]{1,80}?\d{1,2}:\d{2}\s+([A-Z]{3})\b',text)
        or re.search(r'From\s+City(?:\s*/\s*Station)?\s*:?\s*\n\s*([A-Za-z][A-Za-z .]{2,30}?)\s*\n\s*To\s+City(?:\s*/\s*Station)?\s*:?\s*\n\s*([A-Za-z][A-Za-z .]{2,30})',text,re.I))
    passenger=get_passenger(text)
    amount=first([r'(?:Grand Total|Total Trip Cost|Net Amount(?:\(INR\))?|Paid Amount|Total fare|Fee paid|Fee Payment|(?<!Sub[- ])\bTOTAL\b)\s*:?\s*(?:INR|₹|£|€|\$|USD|GBP|EUR|Rs\.?)?\s*([\d,]+(?:\.\d+)?)'],text)
    inv=first([r'Invoice\s*(?:No\.?|Number)?\s*[:#\-]?\s*([A-Z0-9\-/]+)',r'Reference\s+No\s*[:\-]?\s*([A-Z0-9\-]+)'],text)
    pnr=get_pnr(text)
    booking=date_field(['Booking Date','Date of booking','Invoice Date','Bill Date'],text)
    depart=date_field(['Departing','Departure Date','Check-in','Check in'],text)
    arrive=date_field(['Arrival Date','Arriving','Check-out','Check out'],text)
    travel=depart or date_field(['Travel Date'],text) or find_travel_date(text,route,booking)
    airline=get_airline(text)
    if dtype=='Hotel Booking':
        ci,co=hotel_dates(text,booking)
        if ci: travel=ci; depart=ci; arrive=co or arrive
    hotelcity=''; visacountry=''; plan=''
    if dtype=='Visa': visacountry=visa_details(text)
    if dtype=='Insurance': plan=first([r'(?:Plan|Product|Policy Type)\s*[:\-]\s*([^\n]+)'],text)
    curr='INR' if ('₹' in text or re.search(r'\bINR\b',text,re.I)) else ('GBP' if '£' in text else ('EUR' if '€' in text else ('USD' if '$' in text else '')))
    amt=num(amount)
    vt=vfs_grand_total_sum(text)
    if vt is not None: amt=vt
    vt2=us_visa_fee_sum(text)
    if vt2 is not None: amt=vt2
    comm=money_field('Commission(?: Earned)?',text)
    if dtype=='Hotel Booking':
        ht=hotel_total(text)
        if ht: amt=ht
    if amt is None: amt=cleartrip_ticket_total(text)
    if amt is None: amt=total_amount_label(text)
    if comm is None: comm=agent_earnings(text,amt)
    return {'Hotel Name':hotel_name(text) if dtype=='Hotel Booking' else '','Vendor':vendor(text),'Amount Charged':amt,'Commission':comm,'CGST':money_field('CGST',text),'SGST':money_field('SGST',text),'TDS':money_field('TDS(?: Deducted)?',text),
      'Passenger Name':passenger,'Travel Date':travel,'From City / Station':route.group(1) if route else '','To City / Station':route.group(2) if route else '',
      'Arrival Date':arrive,'Departure Date':depart,'Airline Name':airline,'Hotel City':hotelcity,'Visa Country':visacountry,'Insurance Plan':plan,
      'PNR / Booking Confirmation':pnr,'Booking Date':booking,'Vendor Invoice Number':inv,'Type of Document':dtype,'Currency':curr,'Source File':filename}

AI_SCHEMA={k:'' for k in COLUMNS[:21]}
AI_SCHEMA.update({'Amount Charged':None,'Commission':None,'CGST':None,'SGST':None,'TDS':None})

def data_url_for(path):
    ext=path.suffix.lower()
    if ext=='.pdf':
        imgs=list(render_pdf_pages(path,dpi=180)); img=imgs[0] if imgs else None
    else:
        img=Image.open(path).convert('RGB')
    if img is None:return None
    img.thumbnail((1800,1800)); b=io.BytesIO(); img.save(b,'JPEG',quality=85)
    return 'data:image/jpeg;base64,'+base64.b64encode(b.getvalue()).decode()

def ai_parse(text,filename,path):
    base=os.getenv('AI_BASE_URL','https://api.openai.com/v1').rstrip('/'); key=os.getenv('AI_API_KEY'); model=os.getenv('AI_MODEL','gpt-4.1-mini')
    if not key:raise RuntimeError('AI_API_KEY is not set')
    rules=f'''Extract this travel document into exactly this JSON schema: {json.dumps(AI_SCHEMA)}.
Type of Document MUST be one of {DOC_TYPES}. Numeric money fields must be numbers or null. Never invent missing data.
For air tickets: From/To are journey endpoints; Airline Name is carrier; Travel Date is departure travel date.
For hotel: Hotel City, Arrival Date=check-in, Departure Date=check-out.
For visa: Visa Country is destination country. For insurance: Insurance Plan is plan/product name.
For train/bus use From City / Station and To City / Station. Booking Date means bill/booking/invoice date. Vendor Invoice Number is invoice number, not PNR.
Return JSON only. Filename: {filename}\nOCR/TEXT:\n{text[:40000]}'''
    content=[{'type':'text','text':rules}]
    try:
        du=data_url_for(path)
        if du:content.append({'type':'image_url','image_url':{'url':du,'detail':'high'}})
    except Exception:pass
    payload={'model':model,'messages':[{'role':'system','content':'You are a precise travel-accounting document extractor. Read printed and handwritten content.'},{'role':'user','content':content}], 'temperature':0}
    r=requests.post(base+'/chat/completions',headers={'Authorization':f'Bearer {key}','Content-Type':'application/json'},json=payload,timeout=180);r.raise_for_status()
    raw=r.json()['choices'][0]['message']['content'].strip();raw=re.sub(r'^```(?:json)?\s*|\s*```$','',raw,flags=re.I);x=json.loads(raw)
    out={k:x.get(k,'') for k in COLUMNS[:21]};out['Source File']=filename;return out

HOTEL_VOTE_FIELDS=['Amount Charged','Commission','Passenger Name','PNR / Booking Confirmation','Travel Date','Arrival Date','Hotel Name']
def hotel_consensus(rec,variants,filename):
    """OCR of a screenshot differs slightly per preprocessing variant. For hotel documents parse every variant and keep the most common value per field."""
    from collections import Counter
    recs=[rec]
    for t in variants:
        try:
            r=local_parse(t,filename)
            if r.get('Type of Document')=='Hotel Booking': recs.append(r)
        except Exception: pass
    if len(recs)<3: return rec
    for f in HOTEL_VOTE_FIELDS:
        vals=[r.get(f) for r in recs if r.get(f) not in ('',None)]
        if not vals: continue
        top=Counter(vals).most_common()
        best=top[0][1]
        cands=[v for v,c in top if c==best]
        rec[f]=rec[f] if rec.get(f) in cands else cands[0]     # tie -> keep the originally chosen text's value
    return rec

def collect_input(inp,work):
    if inp.is_dir():root=inp
    elif inp.suffix.lower()=='.zip':
        root=work/'unzipped';root.mkdir()
        with zipfile.ZipFile(inp) as z:
            for m in z.infolist():
                dest=(root/m.filename).resolve()
                if not str(dest).startswith(str(root.resolve())):raise ValueError('Unsafe ZIP path')
            z.extractall(root)
    else:root=work/'single';root.mkdir();shutil.copy2(inp,root/inp.name)
    return sorted(p for p in root.rglob('*') if p.is_file() and p.suffix.lower() in SUPPORTED)

# ---- Output layout (matches the required template exactly) ----
OUT_COLUMNS=['Amount','Comission','TDS','CGST','SGST','Discription']

def build_description(r):
    """Passenger name, from - to, PNR, airline name, doc type, travelling date -> one cell. (Hotels: hotel name, confirmation, check-in/out)"""
    if r.get('Extraction Method') in ('error','fatal error'):
        return f"ERROR: could not read {r.get('Source File','')} (see output/extraction_log.txt)"
    def g(k):
        v=r.get(k)
        return '' if v is None or (isinstance(v,float) and v!=v) else v
    hotel=g('Type of Document')=='Hotel Booking'; visa=g('Type of Document')=='Visa'
    frm=str(g('From City / Station')).strip(); to=str(g('To City / Station')).strip()
    route=f'{frm} - {to}' if frm and to else (frm or to)
    parts=[('Passenger',g('Passenger Name')),('Route',route),('Hotel',g('Hotel Name') if hotel else ''),
           ('Confirmation' if hotel else 'PNR',g('PNR / Booking Confirmation')),
           ('Airline',g('Airline Name')),('Doc Type',g('Type of Document')),
           ('Visa Type',g('Visa Country') if visa else ''),('Travel Date',g('Travel Date'))]
    if hotel: parts.append(('Check-out',g('Arrival Date')))
    return ' | '.join(f'{k}: {str(v).strip()}' for k,v in parts if v not in ('',None))

def write_excel(rows,out):
    full=pd.DataFrame(rows,columns=COLUMNS)
    sort=pd.to_datetime(full['Travel Date'].replace('',pd.NA),errors='coerce',dayfirst=True)
    full=full.assign(_sort=sort).sort_values(['_sort','Passenger Name','Type of Document'],na_position='last').drop(columns='_sort')
    df=pd.DataFrame({'Amount':full['Amount Charged'],'Comission':full['Commission'],'TDS':full['TDS'],
                     'CGST':full['CGST'],'SGST':full['SGST'],
                     'Discription':[build_description(r) for r in full.to_dict('records')]},columns=OUT_COLUMNS)
    with pd.ExcelWriter(out,engine='openpyxl') as w:
        df.to_excel(w,index=False,sheet_name='Travel Documents');ws=w.book['Travel Documents'];ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for c in ws[1]:c.font=Font(bold=True,size=12);c.alignment=Alignment(horizontal='center',vertical='center')
        for i,x in enumerate([14,14,12,12,12,110],1):ws.column_dimensions[get_column_letter(i)].width=x
        for row in ws.iter_rows(min_row=2):
            for c in row:c.alignment=Alignment(vertical='top',wrap_text=True)
        for col in range(1,6):
            for c in ws.iter_cols(min_col=col,max_col=col,min_row=2):
                for x in c:x.number_format='#,##0.00'

def _clean(v):
    """NaN / numpy types / empty strings -> JSON-friendly python values."""
    if v is None: return None
    try:
        if v!=v: return None            # NaN
    except Exception: pass
    if hasattr(v,'item'):
        try: v=v.item()
        except Exception: pass
    if isinstance(v,str):
        v=v.strip()
        return v if v else None
    return v

def write_json(rows,out):
    """Same 6 fields as the Excel template + a structured 'details' block per record."""
    full=pd.DataFrame(rows,columns=COLUMNS)
    sort=pd.to_datetime(full['Travel Date'].replace('',pd.NA),errors='coerce',dayfirst=True)
    full=full.assign(_sort=sort).sort_values(['_sort','Passenger Name','Type of Document'],na_position='last').drop(columns='_sort')
    records=[]
    for r in full.to_dict('records'):
        details={k:_clean(r.get(k)) for k in COLUMNS if k not in ('Amount Charged','Commission','TDS','CGST','SGST')}
        records.append({
            'Amount':_clean(r.get('Amount Charged')),
            'Comission':_clean(r.get('Commission')),
            'TDS':_clean(r.get('TDS')),
            'CGST':_clean(r.get('CGST')),
            'SGST':_clean(r.get('SGST')),
            'Discription':build_description(r),
            'Details':details,
        })
    out.write_text(json.dumps({'count':len(records),'records':records},indent=2,ensure_ascii=False),encoding='utf-8')

def write_output(rows,out,fmt):
    """fmt: 'json' | 'excel' | 'both'. Returns list of written paths."""
    written=[]
    if fmt in ('json','both'):
        jp=out.with_suffix('.json'); write_json(rows,jp); written.append(jp)
    if fmt in ('excel','both'):
        xp=out.with_suffix('.xlsx'); write_excel(rows,xp); written.append(xp)
    return written

def load_dotenv_simple(path):
    if not path.exists(): return
    for line in path.read_text(encoding='utf-8',errors='ignore').splitlines():
        line=line.strip()
        if not line or line.startswith('#') or '=' not in line: continue
        k,v=line.split('=',1); os.environ.setdefault(k.strip(),v.strip().strip('\"').strip("'"))

def main():
    print('=== travel_extractor v4.6 (MakeMyTrip, Cleartrip ticket, VFS/UKVI visa, Prominent invoice support) | file:',Path(__file__).resolve(),'===')
    base_dir=Path(__file__).resolve().parent
    load_dotenv_simple(base_dir/'.env')
    # Re-run configuration after .env was loaded
    global TESSERACT_PATH
    TESSERACT_PATH=configure_tesseract()

    ap=argparse.ArgumentParser(description='Travel documents -> JSON / Excel')
    ap.add_argument('input',nargs='?',default=str(base_dir/'input'),help='PDF/image, ZIP, or folder. Default: ./input')
    ap.add_argument('-o','--output',default=str(base_dir/'output'/'travel_documents.json'),help='Output file path (extension is set from --format). Default: output/travel_documents.json')
    ap.add_argument('--format',choices=['json','excel','both'],default='json',help='Output format. Default: json')
    ap.add_argument('--mode',choices=['local','ai','auto'],default='auto')
    a=ap.parse_args()
    inp=Path(a.input).expanduser().resolve(); out=Path(a.output).expanduser().resolve()
    out.parent.mkdir(parents=True,exist_ok=True)
    log_path=out.parent/'extraction_log.txt'
    rows=[]; logs=[]

    try:
        if not inp.exists():
            if str(inp)==str((base_dir/'input').resolve()):
                inp.mkdir(parents=True,exist_ok=True)
            raise FileNotFoundError(f'Input not found or input folder is empty: {inp}')
        use_ai=a.mode=='ai' or (a.mode=='auto' and bool(os.getenv('AI_API_KEY')))
        with tempfile.TemporaryDirectory() as td:
            files=collect_input(inp,Path(td))
            if not files:
                raise RuntimeError(f'No supported files found in {inp}. Supported: {sorted(SUPPORTED)}')
            for p in files:
                try:
                    text,method=extract_local_text(p)
                    rec=ai_parse(text,p.name,p) if use_ai else local_parse(text,p.name)
                    if (not use_ai) and rec.get('Type of Document')=='Hotel Booking' and method=='image-ocr':
                        rec=hotel_consensus(rec,getattr(ocr_image,'last_variants',[]),p.name)
                    rec['Extraction Method']=('AI vision + '+method) if use_ai else method
                    essentials=['Vendor','Amount Charged','Passenger Name','Type of Document']
                    rec['Review Needed']='Yes' if any(rec.get(k) in ('',None) for k in essentials) else 'No'
                    rows.append(rec); msg=f'OK  {p.name} -> {rec["Type of Document"]} ({method})'; print(msg); logs.append(msg)
                    # Save OCR text for troubleshooting
                    dbg=out.parent/'debug_ocr'; dbg.mkdir(exist_ok=True)
                    (dbg/(p.stem+'.txt')).write_text(text,encoding='utf-8',errors='ignore')
                    print('    passenger =',repr(rec.get('Passenger Name')),'| pnr =',repr(rec.get('PNR / Booking Confirmation')))
                except Exception as e:
                    msg=f'ERR {p.name}: {type(e).__name__}: {e}'; print(msg); logs.append(msg)
                    rec={k:'' for k in COLUMNS};rec.update({'Source File':p.name,'Extraction Method':'error','Review Needed':'Yes'});rows.append(rec)
        for wp in write_output(rows,out,a.format):
            msg=f'OUTPUT: {wp} ({len(rows)} record(s))'; print(msg); logs.append(msg)
    except Exception as e:
        msg=f'FATAL: {type(e).__name__}: {e}'; print(msg); logs.append(msg); logs.append(traceback.format_exc())
        # Always create an Excel file, even when setup/input fails, so the failure is visible.
        rec={k:'' for k in COLUMNS}; rec.update({'Source File':'','Extraction Method':'fatal error','Review Needed':'Yes'})
        try: write_output([rec],out,a.format)
        except Exception as ee: logs.append(f'Could not create fallback Excel: {ee}')
    finally:
        log_path.write_text('\n'.join(logs),encoding='utf-8')
        print(f'LOG: {log_path}')

if __name__=='__main__':
    main()