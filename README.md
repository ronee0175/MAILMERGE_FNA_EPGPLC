# Mail Merge PDF Mailer — Phase 1 Server

This Phase 1 build preserves the existing DOCX → PDF → optional protected PDF → SMTP workflow, but moves runtime storage and templates into a server-friendly structure. LibreOffice and qpdf remain server-side dependencies.

## Run on Windows

1. Create/activate `.venv` with a normal Windows Python installation.
2. Install `requirements.txt`.
3. Install LibreOffice and qpdf and ensure `soffice` and `qpdf` are on PATH.
4. Copy `.env.example` to `.env` and configure SMTP.
5. From the project root run:

   `python backend\main.py`

6. Open `http://127.0.0.1:5000` on the server.
7. From another PC on the same LAN, open `http://SERVER-IP:5000` after allowing TCP 5000 through Windows Firewall.

## Phase 1 notes

- Browser PCs do not need Python, LibreOffice, or qpdf.
- Auto-shutdown/heartbeat behavior from the desktop build is intentionally removed.
- Background jobs are still in-process threads in this phase; persistent job queue/database migration comes later.
