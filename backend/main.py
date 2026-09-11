import os,re,subprocess,tempfile,sqlite3,secrets,json
from pathlib import Path
from datetime import datetime,timedelta
from email.message import EmailMessage
from io import BytesIO
import smtplib, zipfile, socket, urllib.parse, shutil, threading, time
import pandas as pd
from openpyxl.styles import Font
from flask import Flask,render_template,request,redirect,url_for,flash,send_file,jsonify,Response,make_response
from mailmerge import MailMerge
from dotenv import load_dotenv
from cryptography.fernet import Fernet
from werkzeug.utils import secure_filename

BASE=Path(__file__).resolve().parent
ROOT=Path(__file__).resolve().parent.parent
# Load optional server defaults from the project-root .env file.
load_dotenv(ROOT / ".env", override=False)
UPLOADS,OUTPUTS,TMP=ROOT/"storage"/"uploads",ROOT/"storage"/"outputs",ROOT/"storage"/"temp"
CRED_DIR=ROOT/"storage"/"credentials"
for p in (UPLOADS,OUTPUTS,TMP,CRED_DIR): p.mkdir(parents=True,exist_ok=True)
DB=BASE/"mail_log.db"
CRED_KEY_FILE=CRED_DIR/"credentials.key"
CRED_DATA_FILE=CRED_DIR/"credentials.enc"
app=Flask(__name__, template_folder=str(ROOT/"frontend"/"templates")); app.secret_key=os.getenv("FLASK_SECRET_KEY",secrets.token_hex(24))
app.config["MAX_CONTENT_LENGTH"]=int(os.getenv("MAX_UPLOAD_MB","25"))*1024*1024
CANCEL_EVENTS = {}
BATCH_STATE = {}
BATCH_LOCK = threading.Lock()

def db():
    c=sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS logs(
    id INTEGER PRIMARY KEY,
    batch,
    row_no INTEGER,
    email,
    employee_id,
    employee_name,
    status,
    message,
    pdf_name,
    created_at,
    attachments
    )""")
    # Upgrade older BO-ID logs databases in place.
    cols={r[1] for r in c.execute("PRAGMA table_info(logs)").fetchall()}
    if "employee_id" not in cols:
        c.execute("ALTER TABLE logs ADD COLUMN employee_id")
    if "employee_name" not in cols:
        c.execute("ALTER TABLE logs ADD COLUMN employee_name")
    if "attachments" not in cols:
        c.execute("ALTER TABLE logs ADD COLUMN attachments")
    c.commit()
    return c
def ok_email(x): return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$",str(x).strip()))

def norm_field(x):
    # Makes Word/Excel headers tolerant of spaces, underscores, hyphens and punctuation.
    return re.sub(r"[^a-z0-9]", "", str(x).strip().lower())

def cmap(df):
    return {norm_field(c): c for c in df.columns}

def find_col(cm, *names):
    for name in names:
        c=cm.get(norm_field(name))
        if c is not None: return c
    return None

def encrypt(src,dst,pw):
    if shutil.which("qpdf") is None:
        raise RuntimeError("qpdf is required for Protected PDF mode. Install qpdf and add it to Windows PATH.")
    r=subprocess.run(["qpdf","--encrypt",pw,pw,"256","--",str(src),str(dst)],capture_output=True,text=True)
    if r.returncode: raise RuntimeError(r.stderr.strip() or "qpdf encryption failed")

def to_pdf(docx,out):
    r=subprocess.run(["soffice","--headless","--convert-to","pdf","--outdir",str(out),str(docx)],capture_output=True,text=True)
    if r.returncode: raise RuntimeError(r.stderr.strip() or "LibreOffice conversion failed")
    p=out/(docx.stem+".pdf")
    if not p.exists(): raise RuntimeError("PDF was not created")
    return p

def _credential_cipher():
    CRED_DIR.mkdir(parents=True, exist_ok=True)
    if not CRED_KEY_FILE.exists():
        CRED_KEY_FILE.write_bytes(Fernet.generate_key())
    return Fernet(CRED_KEY_FILE.read_bytes())

def _load_saved_credentials():
    if not CRED_DATA_FILE.exists():
        return {}
    try:
        raw=_credential_cipher().decrypt(CRED_DATA_FILE.read_bytes())
        data=json.loads(raw.decode("utf-8"))
        return data if isinstance(data,dict) else {}
    except Exception:
        return {}

def _save_saved_credentials(data):
    payload=json.dumps(data,ensure_ascii=False).encode("utf-8")
    tmp=CRED_DATA_FILE.with_suffix(".tmp")
    tmp.write_bytes(_credential_cipher().encrypt(payload))
    tmp.replace(CRED_DATA_FILE)

def _client_id():
    cid=request.cookies.get("mm_client_id", "").strip()
    if not re.fullmatch(r"[A-Za-z0-9_-]{32,100}", cid):
        cid=secrets.token_urlsafe(32)
    return cid

def _credential_record(provider, host, user, password, port=587, ssl=False):
    now=datetime.now()
    return {
        "provider":provider,
        "host":host,
        "user":user,
        "password":password,
        "port":int(port),
        "ssl":bool(ssl),
        "saved_at":now.isoformat(timespec="seconds"),
        "expires_at":(now.replace(hour=0,minute=0,second=0,microsecond=0)+timedelta(days=1)).isoformat(timespec="seconds"),
    }

def _get_saved_for_client(provider):
    cid=_client_id()
    data=_load_saved_credentials()
    rec=data.get(cid,{}).get(provider)
    if not rec:
        return None
    try:
        exp=datetime.fromisoformat(rec["expires_at"])
    except Exception:
        return None
    if exp <= datetime.now():
        data.get(cid,{}).pop(provider,None)
        if not data.get(cid): data.pop(cid,None)
        _save_saved_credentials(data)
        return None
    return rec

def _save_credentials_for_client(provider, host, user, password, port=587, ssl=False):
    cid=_client_id(); data=_load_saved_credentials(); existing=data.get(cid,{}).get(provider)
    if existing:
        try:
            same=(existing.get("host")==host and existing.get("user")==user and existing.get("password")==password and int(existing.get("port",587))==int(port) and bool(existing.get("ssl",False))==bool(ssl) and datetime.fromisoformat(existing.get("expires_at"))>datetime.now())
        except Exception:
            same=False
        if same:
            return cid, existing
    data.setdefault(cid,{})[provider]=_credential_record(provider,host,user,password,port,ssl)
    _save_saved_credentials(data)
    return cid, data[cid][provider]

def _forget_credentials_for_client(provider):
    cid=_client_id(); data=_load_saved_credentials()
    if cid in data:
        data[cid].pop(provider,None)
        if not data[cid]: data.pop(cid,None)
        _save_saved_credentials(data)
    return cid

def smtp_settings(credentials=None):
    if credentials:
        host=str(credentials.get("host","")).strip()
        user=str(credentials.get("user","")).strip()
        pw=str(credentials.get("password","")).strip()
        try: port=int(credentials.get("port",587))
        except (TypeError,ValueError): raise ValueError("SMTP_PORT must be a number, normally 587.")
        ssl=bool(credentials.get("ssl",False))
    else:
        host=os.getenv("SMTP_HOST","").strip(); user=os.getenv("SMTP_USER","").strip(); pw=os.getenv("SMTP_PASSWORD","").strip()
        if "://" in host: host=urllib.parse.urlparse(host).hostname or host
        if host.count(":")==1 and not host.startswith("["): host=host.split(":",1)[0]
        try: port=int(os.getenv("SMTP_PORT","587"))
        except ValueError: raise ValueError("SMTP_PORT must be a number, normally 587.")
        ssl=os.getenv("SMTP_SSL","0").strip()=="1"
    if not host or not user or not pw:
        missing=[k for k,v in (("SMTP_HOST",host),("SMTP Email",user),("SMTP Password",pw)) if not v]
        raise ValueError("SMTP credentials are incomplete. Missing: "+", ".join(missing))
    return host,port,user,pw,ssl

def mail(to,att,employee_id,employee_name,subject,body,extra_attachment=None,smtp_credentials=None):
    host,port,user,pw,ssl=smtp_settings(smtp_credentials)
    replacements={
        "{EMPLOYEE_ID}": employee_id,
        "{EMPLOYEE_NAME}": employee_name,
        "{EMAIL}": to,
    }
    rendered_subject=subject
    rendered_body=body
    for token,value in replacements.items():
        rendered_subject=rendered_subject.replace(token,value)
        rendered_body=rendered_body.replace(token,value)
    m=EmailMessage(); m["Subject"]=rendered_subject; m["From"]=os.getenv("MAIL_FROM",user).strip(); m["To"]=to
    m.set_content(rendered_body)
    m.add_attachment(att.read_bytes(),maintype="application",subtype="pdf",filename=att.name)
    if extra_attachment:
        for extra_bytes,extra_name,extra_maintype,extra_subtype in extra_attachment:
            m.add_attachment(
                extra_bytes,
                maintype=extra_maintype,
                subtype=extra_subtype,
                filename=extra_name
            )
    try:
        if ssl:
            with smtplib.SMTP_SSL(host,port,timeout=20) as s: s.login(user,pw); s.send_message(m)
        else:
            with smtplib.SMTP(host,port,timeout=20) as s: s.starttls(); s.login(user,pw); s.send_message(m)
    except socket.gaierror as e:
        raise RuntimeError(f"SMTP server cannot be found: '{host}'. Check SMTP_HOST. Do not include https://, :port, or spaces. For Gmail use smtp.gmail.com with port 587.") from e
    except TimeoutError as e:
        raise RuntimeError(f"Could not connect to SMTP server '{host}:{port}'. Check the host, port, firewall, and internet connection.") from e
    except smtplib.SMTPAuthenticationError as e:
        raise RuntimeError("SMTP login failed. Check SMTP_USER and use a Google App Password (not your normal Gmail password) for Gmail.") from e
    except smtplib.SMTPException as e:
        raise RuntimeError(f"SMTP error from {host}:{port}: {e}") from e

def merge_data(row, merge_fields):
    raw={str(k):("" if pd.isna(v) else str(v)) for k,v in row.to_dict().items()}
    normalized={norm_field(k):v for k,v in raw.items()}
    aliases={
        "employeeid":["employeeid","employee_id","employee id","empid","emp_id","emp id","employee code","employee_code","employee no","employee number","emp no","emp number"],
        "employeename":["employeename","employee_name","employee name","name","employee fullname","employee full name","full name","accounttitle","account_title","account title"],
        "accounttitle":["accounttitle","account_title","account title"],
        "boid":["boid","bo_id","bo id"],
        "balanc1":["balanc1"],
        "balance1":["balance1","balance_1","balance 1","balanc_1","balanc 1"],
    }
    out={}
    missing=[]
    for field in merge_fields:
        key=norm_field(field)
        value=None
        if key in normalized:
            value=normalized[key]
        else:
            # Common naming variants, including Address_Line_1 / Address Line 1 / Address1.
            candidates=[key]
            if key.startswith("addressline"):
                n=key[len("addressline"):]
                candidates += ["address"+n, "address_line"+n, "address line"+n]
            if key.startswith("balance"):
                n=key[len("balance"):]
                candidates += ["balanc"+n, "balance"+n]
            if key in aliases: candidates += aliases[key]
            for candidate in candidates:
                if norm_field(candidate) in normalized:
                    value=normalized[norm_field(candidate)]
                    break
        if value is None:
            # BO_ID is no longer part of the Employee-ID workflow. If an older
            # Word template still contains a BO_ID field but the Excel has no
            # BO_ID column, leave that field blank instead of blocking the merge.
            if key == "boid":
                out[field] = ""
            else:
                missing.append(field)
        else:
            out[field]=value
    return out, missing

def safe_filename(value):
    # Keep Unicode names while removing Windows/path-breaking characters.
    value=re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value))
    return value.strip().rstrip(".") or "document"

def one(row,template,employee_id,employee_name,pdf_name_template="Payslip - {EMPLOYEE_ID} - {EMPLOYEE_NAME}",protect_pdf=True):
    with tempfile.TemporaryDirectory(dir=TMP) as td:
        td=Path(td); doc=td/"merged.docx"; pd=td/"pdf"; pd.mkdir()
        with MailMerge(str(template)) as m:
            fields=m.get_merge_fields()
            data,missing=merge_data(row,fields)
            if missing:
                raise ValueError("Missing Word fields: "+", ".join(missing))
            m.merge(**data); m.write(str(doc))
        plain=to_pdf(doc,pd)
        replacements={"{EMPLOYEE_ID}": employee_id, "{EMPLOYEE_NAME}": employee_name, "{EMAIL}": str(row.get("Email", "")) if hasattr(row, "get") else ""}
        rendered_name=str(pdf_name_template or "Payslip - {EMPLOYEE_ID} - {EMPLOYEE_NAME}")
        for token,value in replacements.items():
            rendered_name=rendered_name.replace(token, value)
        base=safe_filename(rendered_name)
        final=OUTPUTS/(base+".pdf")
        if protect_pdf:
            encrypt(plain,final,employee_id)
        else:
            shutil.copy2(plain,final)
        return final

def process(excel,template,subject,body,batch,selected=None,on_result=None,protect_pdf=True,pdf_name_template="Payslip - {EMPLOYEE_ID} - {EMPLOYEE_NAME}"):
    df=pd.read_excel(excel,dtype=str).fillna(""); cm=cmap(df)
    ec=find_col(cm,"Email","Email Address","Email_Address")
    ic=find_col(cm,"Employee_ID","Employee ID","EmployeeID","Emp_ID","Emp ID")
    nc=find_col(cm,"Employee_Name","Employee Name","EmployeeName","Name","Account Title","Account_Title")
    if not ec or not ic or not nc:
        raise ValueError("Excel must contain Email, Employee ID and Employee Name (or Name/Full Name) columns. BO_ID is not required.")
    vals=df[ic].astype(str).str.strip(); dup=set(vals[vals.duplicated(keep=False)])
    chosen=set(selected) if selected else None; con=db(); out=[]
    total=sum(1 for i in df.index if not chosen or i+2 in chosen); done=0
    for i,row in df.iterrows():
        rn=i+2
        if chosen and rn not in chosen: continue
        email=str(row[ec]).strip(); employee_id=str(row[ic]).strip(); employee_name=str(row[nc]).strip(); status="FAILED"; msg=""; pdf=""
        try:
            if not email or not employee_id or not employee_name: raise ValueError("Missing Email, Employee ID or Employee Name")
            if not ok_email(email): raise ValueError("Invalid email")
            if employee_id in dup: raise ValueError("Duplicate Employee ID")
            p=one(row,template,employee_id,employee_name,pdf_name_template,protect_pdf); pdf=str(p.relative_to(OUTPUTS)); mail(email,p,employee_id,employee_name,subject,body,extra_attachment,saved_cred); status="SENT"; msg="Email sent"
        except Exception as e: msg=str(e)
        con.execute("INSERT INTO logs(batch,row_no,email,employee_id,employee_name,status,message,pdf_name,created_at,attachments) VALUES(?,?,?,?,?,?,?,?,?,?)",(batch,rn,email,employee_id,employee_name,status,msg,pdf,datetime.now().isoformat(timespec="seconds"),", ".join(x.name for x in xps))); con.commit()
        done+=1
        if on_result: on_result(result,done,total)
    con.close(); return out

@app.route("/",methods=["GET"])
def index():
    provider=request.args.get("provider","outlook").strip().lower()
    if provider not in ("outlook","gmail","others"): provider="outlook"
    rec=_get_saved_for_client(provider)
    response=make_response(render_template("index.html",results=None,auto_shutdown=False,saved_credentials=rec,selected_provider=provider))
    if not request.cookies.get("mm_client_id"):
        response.set_cookie("mm_client_id",_client_id(),max_age=60*60*24*365,httponly=True,samesite="Lax")
    return response

@app.route("/credentials",methods=["GET","DELETE"])
def credentials_api():
    provider=request.args.get("provider","outlook").strip().lower()
    if provider not in ("outlook","gmail","others"): provider="others"
    if request.method=="DELETE":
        cid=_forget_credentials_for_client(provider)
        response=make_response(jsonify(ok=True))
        if not request.cookies.get("mm_client_id"):
            response.set_cookie("mm_client_id",cid,max_age=60*60*24*365,httponly=True,samesite="Lax")
        return response
    rec=_get_saved_for_client(provider)
    response=make_response(jsonify(ok=True,saved=bool(rec),credential=rec or {}))
    if not request.cookies.get("mm_client_id"):
        response.set_cookie("mm_client_id",_client_id(),max_age=60*60*24*365,httponly=True,samesite="Lax")
    return response

@app.route("/send",methods=["POST"])
def send_stream():
    """Start mail merge in a background worker and return immediately.

    SMTP credentials entered in the form are saved encrypted on the server for
    24 hours, per browser, and are reused until they expire or are forgotten.
    """
    provider=request.form.get("email_provider","outlook").strip().lower()
    if provider not in ("outlook","gmail","others"): provider="others"
    smtp_host=request.form.get("smtp_host","").strip()
    smtp_email=request.form.get("smtp_email","").strip()
    smtp_password=request.form.get("smtp_password","").strip()
    smtp_port=request.form.get("smtp_port","587").strip() or "587"
    smtp_ssl=request.form.get("smtp_ssl","0").strip()=="1"
    if not smtp_host or not smtp_email or not smtp_password:
        return jsonify(ok=False, message="SMTP Host, SMTP Email and SMTP Password are required."), 400
    if not ok_email(smtp_email):
        return jsonify(ok=False, message="Please enter a valid SMTP email address."), 400
    try: int(smtp_port)
    except ValueError: return jsonify(ok=False, message="SMTP Port must be a number."), 400
    client_id, saved_cred=_save_credentials_for_client(provider,smtp_host,smtp_email,smtp_password,int(smtp_port),smtp_ssl)
    e=request.files.get("excel"); t=request.files.get("template")
    if not e or not t:
        return jsonify(ok=False, message="Excel and Word files are required."), 400
    if Path(e.filename).suffix.lower()!=".xlsx" or Path(t.filename).suffix.lower()!=".docx":
        return jsonify(ok=False, message="Only .xlsx and .docx are accepted."), 400
    extra_files=request.files.getlist("extra_attachment")
    xps=[]

    for x in extra_files:
        if x and x.filename:
            if Path(x.filename).suffix.lower() not in (".docx",".pdf"):
                return jsonify(ok=False, message="Additional attachment must be a .docx or .pdf file."), 400
            xp=UPLOADS/safe_filename(x.filename)
            x.save(xp)
            xps.append(xp)

    ep=UPLOADS/secure_filename(e.filename); tp=UPLOADS/secure_filename(t.filename)
    e.save(ep); t.save(tp)
    batch=datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    subject=request.form.get("subject","Your Employee Document")
    pdf_mode=request.form.get("pdf_mode","protected").strip().lower()
    if pdf_mode not in ("plain","protected"): pdf_mode="protected"
    body=request.form.get("body","Dear recipient,\n\nPlease find your PDF attached.\n\nRegards")
    pdf_name_template=request.form.get("pdf_name", "Payslip - {EMPLOYEE_ID} - {EMPLOYEE_NAME}").strip() or "Payslip - {EMPLOYEE_ID} - {EMPLOYEE_NAME}"
    row_mode=request.form.get("row_mode","all").strip().lower()
    start_row_raw=request.form.get("start_row","").strip()
    end_row_raw=request.form.get("end_row","").strip()
    protect_pdf=(pdf_mode=="protected")
    cancel_event=threading.Event()
    with BATCH_LOCK:
        CANCEL_EVENTS[batch]=cancel_event
        BATCH_STATE[batch]={"type":"start","batch":batch,"status":"starting","done":0,"total":0,"sent":0,"failed":0,"processed":0,"message":"Starting…"}

    def worker():
        try:
            extra_attachment=[]
            for xp in xps:
                if xp.suffix.lower()==".docx":
                    extra_maintype,extra_subtype="application","vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    extra_maintype,extra_subtype="application","pdf"
                extra_attachment.append((xp.read_bytes(),xp.name,extra_maintype,extra_subtype))
            df=pd.read_excel(ep,dtype=str).fillna(""); cm=cmap(df)
            ec=find_col(cm,"Email","Email Address","Email_Address")
            ic=find_col(cm,"Employee_ID","Employee ID","EmployeeID","Emp_ID","Emp ID")
            nc=find_col(cm,"Employee_Name","Employee Name","EmployeeName","Name","Account Title","Account_Title")
            if not ec or not ic or not nc:
                raise ValueError("Excel must contain Email, Employee ID and Employee Name (or Name/Full Name) columns. BO_ID is not required.")
            data_first_row=2; data_last_row=len(df)+1
            if row_mode=="range":
                try:
                    start_row=int(start_row_raw); end_row=int(end_row_raw)
                except ValueError:
                    raise ValueError("Please enter valid Start Row and End Row numbers.")
                if start_row<data_first_row or end_row<data_first_row:
                    raise ValueError(f"Row range must start from Excel row {data_first_row} or later because row 1 is the header.")
                if start_row>end_row: raise ValueError("Start Row cannot be greater than End Row.")
                if end_row>data_last_row: raise ValueError(f"End Row cannot be greater than {data_last_row}. This Excel file has {len(df)} data rows.")
                selected_rows=set(range(start_row,end_row+1))
            elif row_mode=="all": selected_rows=None
            else: raise ValueError("Invalid row selection mode.")
            check_df=df if selected_rows is None else df.loc[[r-2 for r in sorted(selected_rows)]]
            check_vals=check_df[ic].astype(str).str.strip(); dup=set(check_vals[check_vals.duplicated(keep=False)])
            con=db(); total=len(df) if selected_rows is None else len(selected_rows); done=0
            with BATCH_LOCK:
                BATCH_STATE[batch].update({"status":"running","total":total,"message":"Processing…"})
            for i,row in df.iterrows():
                rn=i+2
                if selected_rows is not None and rn not in selected_rows: continue
                if cancel_event.is_set(): break
                email=str(row[ec]).strip(); employee_id=str(row[ic]).strip(); employee_name=str(row[nc]).strip(); status="FAILED"; msg=""; pdf=""
                try:
                    if not email or not employee_id or not employee_name: raise ValueError("Missing Email, Employee ID or Employee Name")
                    if not ok_email(email): raise ValueError("Invalid email")
                    if employee_id in dup: raise ValueError("Duplicate Employee ID")
                    p=one(row,tp,employee_id,employee_name,pdf_name_template,protect_pdf); pdf=str(p.relative_to(OUTPUTS)); mail(email,p,employee_id,employee_name,subject,body,extra_attachment,saved_cred); status="SENT"; msg="Email sent"
                except Exception as ex: msg=str(ex)
                con.execute("INSERT INTO logs(batch,row_no,email,employee_id,employee_name,status,message,pdf_name,created_at,attachments) VALUES(?,?,?,?,?,?,?,?,?,?)",(batch,rn,email,employee_id,employee_name,status,msg,pdf,datetime.now().isoformat(timespec="seconds"),", ".join(x.name for x in xps))); con.commit()
                done+=1
                with BATCH_LOCK:
                    sent=con.execute("SELECT COUNT(*) FROM logs WHERE batch=? AND status='SENT'",(batch,)).fetchone()[0]
                    failed=con.execute("SELECT COUNT(*) FROM logs WHERE batch=? AND status='FAILED'",(batch,)).fetchone()[0]
                    BATCH_STATE[batch].update({"type":"row","status":"running","done":done,"total":total,"row":rn,"email":email,"employee_id":employee_id,"employee_name":employee_name,"status_row":status,"message":msg,"pdf":pdf,"sent":sent,"failed":failed,"processed":done})
            con.close()
            cancelled=cancel_event.is_set()
            with BATCH_LOCK:
                sent=BATCH_STATE[batch].get("sent",0); failed=BATCH_STATE[batch].get("failed",0)
                BATCH_STATE[batch].update({"type":"cancelled" if cancelled else "done","status":"cancelled" if cancelled else "done","done":done,"total":total,"sent":sent,"failed":failed,"processed":done,"message":("Cancelled" if cancelled else "Completed")})
        except Exception as ex:
            with BATCH_LOCK:
                BATCH_STATE[batch].update({"type":"error","status":"error","message":str(ex)})
        finally:
            with BATCH_LOCK: CANCEL_EVENTS.pop(batch,None)

    threading.Thread(target=worker,name=f"mail-merge-{batch}",daemon=True).start()
    return jsonify(ok=True,batch=batch,message="Mail merge started in background.")

@app.route("/status/<batch>", methods=["GET"])
def batch_status(batch):
    with BATCH_LOCK:
        state=BATCH_STATE.get(batch)
        if not state: return jsonify(ok=False,message="Batch not found."),404
        return jsonify(ok=True,**state)

@app.route("/cancel/<batch>", methods=["POST"])
def cancel_batch(batch):
    with BATCH_LOCK:
        event = CANCEL_EVENTS.get(batch)
        if event is None:
            return jsonify(ok=False, message="Batch is not running or has already finished."), 404
        event.set()
        if batch in BATCH_STATE:
            BATCH_STATE[batch]["message"]="Cancellation requested…"
    return jsonify(ok=True, message="Cancellation requested.")

@app.route("/excel_info", methods=["POST"])
def excel_info():
    e=request.files.get("excel")
    if not e or Path(e.filename).suffix.lower() != ".xlsx":
        return jsonify(ok=False, message="Please select a valid .xlsx Excel file."), 400
    try:
        df=pd.read_excel(e,dtype=str)
        return jsonify(ok=True, data_rows=len(df), first_row=2, last_row=len(df)+1)
    except Exception as ex:
        return jsonify(ok=False, message=f"Could not read Excel file: {ex}"), 400

@app.route("/logs")
def logs():
    c=db(); rows=c.execute("SELECT batch,row_no,email,employee_id,employee_name,status,message,pdf_name,attachments,created_at FROM logs ORDER BY id DESC LIMIT 500").fetchall(); c.close()
    return render_template("logs.html",rows=rows,auto_shutdown=False)

@app.route("/logs/export")
def export_logs():
    """Excel export of the same rows shown on the View Logs page (up to 500, newest first)."""
    c=db(); rows=c.execute("SELECT batch,row_no,email,employee_id,employee_name,status,message,pdf_name,created_at,attachments FROM logs ORDER BY id DESC LIMIT 500").fetchall(); c.close()
    columns=["Batch","Row","Email","Employee ID","Name","Status","Message","PDF","Created","Additional Attachments"]
    df=pd.DataFrame(rows,columns=columns)
    buf=BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as writer:
        df.to_excel(writer,index=False,sheet_name="Logs")
        ws=writer.sheets["Logs"]
        ws.freeze_panes="A2"
        ws.auto_filter.ref=ws.dimensions
        for cell in ws[1]:
            cell.font=Font(bold=True)
        for col_cells in ws.columns:
            width=max((len(str(cell.value)) if cell.value is not None else 0) for cell in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width=min(max(width+2,10),60)
    buf.seek(0)
    filename="mail_logs_"+datetime.now().strftime("%Y%m%d_%H%M%S")+".xlsx"
    return send_file(buf,as_attachment=True,download_name=filename,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/download/<path:name>")
def download(name):
    p=(OUTPUTS/name).resolve()
    if not p.is_relative_to(OUTPUTS.resolve()) or not p.exists(): return "Not found",404
    return send_file(p,as_attachment=True)

@app.route("/zip/<path:employee_key>")
def zip_employee(employee_key):
    key=safe_filename(employee_key)
    f=OUTPUTS/key
    if not f.exists(): return "Not found",404
    z=TMP/(safe_filename(key)+".zip")
    with zipfile.ZipFile(z,"w",zipfile.ZIP_DEFLATED) as q:
        for p in f.glob("*.pdf"): q.write(p,p.name)
    return send_file(z,as_attachment=True)

@app.route("/health")
def health(): return jsonify(ok=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=False, threaded=True)
