from fasthtml.common import *  # FastHTML + Starlette + HTMX helpers
import pandas as pd
from datetime import datetime
from unicodedata import normalize
import re, io
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from starlette.responses import StreamingResponse

# -------------------------------------------------
# Helpers
# -------------------------------------------------
def _only_digits(s: str) -> str:
    return re.sub(r"\D+", "", str(s) if s is not None else "")

def _norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    return re.sub(r"\s+", " ", s.strip()).lower()

def _pick_col(df, candidates):
    m = {_norm(c): c for c in df.columns}
    for c in candidates:
        k = _norm(c)
        if k in m: return m[k]
    for c in candidates:
        k = re.sub(r"[^a-z0-9 ]+", "", _norm(c))
        for kk, orig in m.items():
            if re.sub(r"[^a-z0-9 ]+", "", kk) == k:
                return orig
    return None

def _status(venc_dt, today=None):
    if today is None: today = datetime.today()
    if pd.isna(venc_dt): return "Sem data"
    d = (venc_dt - today).days
    if d < 0: return "Vencido"
    if d <= 30: return "A vencer"
    return "No prazo"

def gerar_relatorio(df_sieg, df_cert):
    cnpj_sieg = _pick_col(df_sieg, ["CPF_CNPJ", "CNPJ", "CPF/CNPJ", "CPF CNPJ"])
    cnpj_cert = _pick_col(df_cert, ["CPF_CNPJ", "CNPJ", "CPF/CNPJ", "CPF CNPJ"])
    if not cnpj_sieg or not cnpj_cert:
        raise ValueError("Não encontrei a coluna CPF_CNPJ/CNPJ em uma das planilhas.")

    col_resp = _pick_col(df_sieg, ["Responsável", "Responsavel"])
    col_emp  = _pick_col(df_sieg, ["Empresa", "Razão Social", "Razao Social", "Cliente", "Nome do Cliente"])
    col_venc = _pick_col(df_cert, ["Vencimento", "Validade", "Data de Vencimento", "Data Vencimento"])

    df_sieg["_CPF_CNPJ_"] = df_sieg[cnpj_sieg].map(_only_digits)
    df_cert["_CPF_CNPJ_"] = df_cert[cnpj_cert].map(_only_digits)

    keep = ["_CPF_CNPJ_"] + ([col_venc] if col_venc else [])
    df_cert_small = df_cert[keep].copy()

    merged = pd.merge(df_sieg, df_cert_small, on="_CPF_CNPJ_", how="left")

    out = pd.DataFrame()
    out["Responsavel"] = merged[col_resp] if col_resp else ""
    out["Empresa"]     = merged[col_emp] if col_emp else ""
    out["CPF_CNPJ"]    = merged["_CPF_CNPJ_"]

    if col_venc:
        venc = pd.to_datetime(merged[col_venc], errors="coerce", dayfirst=True)
        out["Vencimento"] = venc.dt.strftime("%d/%m/%Y").fillna("")
        out["Status"]     = [_status(d) for d in venc]
    else:
        out["Vencimento"] = ""
        out["Status"]     = "Sem data"

    return out

def make_excel_bytes(df: pd.DataFrame, sheet_name="Relatorio") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        # Cabeçalho
        header_fill = PatternFill(start_color="222222", end_color="222222", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for col in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            ws.column_dimensions[get_column_letter(col)].width = max(12, len(str(df.columns[col-1])) + 4)

        # Cores por Status
        status_col_idx = None
        for i, name in enumerate(df.columns, start=1):
            if name == "Status":
                status_col_idx = i
                break
        if status_col_idx:
            fill_vencido  = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            fill_avencer  = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            fill_noprazo  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            fill_semdata  = PatternFill(start_color="ECEFF1", end_color="ECEFF1", fill_type="solid")
            for r in range(2, len(df) + 2):
                c = ws.cell(row=r, column=status_col_idx)
                val = (c.value or "").strip()
                if val == "Vencido":
                    c.fill = fill_vencido
                elif val == "A vencer":
                    c.fill = fill_avencer
                elif val == "No prazo":
                    c.fill = fill_noprazo
                else:
                    c.fill = fill_semdata
    buf.seek(0)
    return buf.getvalue()

# -------------------------------------------------
# FastHTML app (sem HTMX no submit – download direto)
# -------------------------------------------------
app, rt = fast_app()

def global_css() -> FT:
    return Style("""
      body{min-height:100vh;background:#0b1020;color:#e5f2ff;font-family:system-ui,Segoe UI,Roboto}
      .container{max-width:980px;margin:0 auto;padding:22px}
      .glass{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,.25);padding:18px}
      .lbl{font-weight:700}
      .filebox{width:100%;padding:12px;border:2px dashed #0ea5e9;border-radius:12px;background:rgba(255,255,255,.08);color:#e5f2ff}
      .btn{background:linear-gradient(90deg,#0ea5e9,#38bdf8);color:#001018;border:none;border-radius:12px;font-weight:800;padding:.8rem 1.1rem;cursor:pointer}
    """)

def page() -> FT:
    return (
        global_css(),
        Main(
            Section(
                Article(
                    H1("SIEG x Certificados — Comparador por CNPJ"),
                    P("Anexe as duas planilhas e baixe o Excel gerado imediatamente."),
                    # >>> Formulário normal (POST) que baixa o arquivo direto <<<
                    Form(action=baixar.to(), method="post", enctype="multipart/form-data")(
                        Div(Label("Planilha SIEG (xlsx/xls)", cls="lbl"),
                            Input(type="file", name="file_sieg", accept=".xlsx,.xls", required=True, cls="filebox")),
                        Div(Label("Planilha Certificados (xlsx/xls)", cls="lbl", style="margin-top:10px"),
                            Input(type="file", name="file_cert", accept=".xlsx,.xls", required=True, cls="filebox")),
                        Div(Button("📥 Baixar Excel", type="submit", cls="btn", style="margin-top:14px"))
                    ),
                    cls="glass"
                ),
                cls="container"
            ),
        )
    )

@rt   # GET /
def index():
    return Titled("SIEG x Certificados — Download Direto", page())

@rt   # POST /baixar  -> retorna o .xlsx diretamente
async def baixar(file_sieg: UploadFile, file_cert: UploadFile):
    try:
        df_sieg = pd.read_excel(io.BytesIO(await file_sieg.read()), dtype=str)
        df_cert = pd.read_excel(io.BytesIO(await file_cert.read()), dtype=str)
        df = gerar_relatorio(df_sieg, df_cert)
        xbytes = make_excel_bytes(df)

        fname = f"Relatorio_SIEG_Certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        headers = {
            "Content-Disposition": f'attachment; filename="{fname}"',
            "X-Content-Type-Options": "nosniff",
            "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
            "Pragma": "no-cache",
            "Expires": "0",
        }
        return StreamingResponse(
            io.BytesIO(xbytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )
    except Exception as e:
        # em caso de erro, mostra uma página simples com a mensagem
        return Titled("Erro", Main(Section(Article(P(f"❌ Erro: {e}"), cls="glass"), cls="container")))

# start
serve()  # http://localhost:5001


