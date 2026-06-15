# -*- coding: utf-8 -*-
import streamlit as st
import os, io, zipfile, tempfile, re, datetime, shutil
from pathlib import Path
from docx import Document
from openpyxl import load_workbook

def rub_to_words(n):
    if n == 0: return "ноль"
    units = ["","один","два","три","четыре","пять","шесть","семь","восемь","девять",
             "десять","одиннадцать","двенадцать","тринадцать","четырнадцать","пятнадцать",
             "шестнадцать","семнадцать","восемнадцать","девятнадцать"]
    tens = ["","","двадцать","тридцать","сорок","пятьдесят","шестьдесят","семьдесят","восемьдесят","девяносто"]
    hundreds = ["","сто","двести","триста","четыреста","пятьсот","шестьсот","семьсот","восемьсот","девятьсот"]
    uf = ["","одна","две","три","четыре","пять","шесть","семь","восемь","девять"]
    def _t(n,f=False):
        if n==0: return ""
        r=[]; h=n//100
        if h: r.append(hundreds[h])
        rem=n%100
        if rem<20:
            if rem>0: r.append(uf[rem] if (f and rem in(1,2)) else units[rem])
        else:
            r.append(tens[rem//10])
            if rem%10>0: r.append(uf[rem%10] if (f and rem%10 in(1,2)) else units[rem%10])
        return " ".join(r)
    def _pl(n,a,b,c):
        if 11<=n%100<=19: return c
        l=n%10
        if l==1: return a
        if 2<=l<=4: return b
        return c
    th=n//1000; rm=n%1000; p=[]
    if th>0:
        t=_t(th,True)
        if t: p.append(t+" "+_pl(th,"тысяча","тысячи","тысяч"))
    if rm>0 or not p:
        t=_t(rm)
        if t: p.append(t)
    return " ".join(p)

def date_to_russian(s):
    m=["","января","февраля","марта","апреля","мая","июня",
       "июля","августа","сентября","октября","ноября","декабря"]
    p=s.strip().split(".")
    if len(p)!=3: return s
    return f"{int(p[0])} {m[int(p[1])]} {p[2]} г."

TMPL_DIR = Path(__file__).parent / "templates"

def get_para_full(para):
    return "".join(r.text for r in para.runs)
def replace_para_text(para, txt):
    for i,r in enumerate(para.runs): r.text = txt if i==0 else ""
def set_cell_text(cell, txt):
    ps=cell.paragraphs
    if ps:
        rs=ps[0].runs
        if rs:
            rs[0].text=str(txt)
            for r in rs[1:]: r.text=""
        else: ps[0].text=str(txt)

def generate_files(data, out_dir):
    date_str = data["date"]
    actual_amt = data["actual_amt"]
    budget_amt = data.get("budget_amt") or actual_amt
    amount_manual = data.get("amount_manual", "").strip()
    company_full = data["company_full"]
    responsible = data["responsible"]
    position = data["position"]
    participant = data["participant"]
    venue = data.get("venue", "")
    address = data.get("address", "")
    purpose = data.get("purpose", "")
    result = data.get("result", "")
    receipt_date = data.get("receipt_date", "")
    receipt_num = data.get("receipt_num", "")
    receipt_amt = data.get("receipt_amt") or actual_amt
    company_short = data.get("company_short") or company_full
    comm1_pos = data.get("comm1_pos", "")
    comm1_name = data.get("comm1_name", "")
    comm2_pos = data.get("comm2_pos", "")
    comm2_name = data.get("comm2_name", "")
    comm3_pos = data.get("comm3_pos", "")
    comm3_name = data.get("comm3_name", "")
    compiler_pos = data.get("compiler_pos", "")
    compiler_name = data.get("compiler_name", "")

    budget_words = amount_manual if amount_manual else rub_to_words(budget_amt)
    actual_words = rub_to_words(actual_amt)
    date_russian = date_to_russian(date_str)
    day, month, year = date_str.split(".")

    rec_date_clean = receipt_date
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
        try:
            dt = datetime.datetime.strptime(receipt_date, fmt)
            rec_date_clean = dt.strftime("%d.%m.%Y")
            break
        except: pass

    out_s = os.path.join(out_dir, "Смета.docx")
    out_p = os.path.join(out_dir, "报销_приказ.docx")
    out_a = os.path.join(out_dir, "АктОтчет.xlsx")

    shutil.copy2(TMPL_DIR / "Смета_шаблон.docx", out_s)
    shutil.copy2(TMPL_DIR / "приказ_шаблон.docx", out_p)
    wb_tmpl = load_workbook(TMPL_DIR / "АктОтчет_шаблон.xlsx")
    wb_tmpl.save(out_a)

    # Смета
    doc = Document(out_s)
    set_cell_text(doc.tables[0].rows[0].cells[1], date_str)
    set_cell_text(doc.tables[1].rows[1].cells[1], str(budget_amt))
    set_cell_text(doc.tables[1].rows[2].cells[1], str(budget_amt))
    doc.save(out_s)

    # приказ
    doc = Document(out_p)
    DATE_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})")
    for para in doc.paragraphs:
        full = get_para_full(para)
        if "Красногорск" in full and DATE_RE.search(full):
            replace_para_text(para, DATE_RE.sub(date_str + " г.", full))
            continue
        if DATE_RE.match(full.strip().rstrip("г.").strip()):
            replace_para_text(para, date_str + " г.")
            continue
        if "В целях установления" in full and "делегация" in full:
            idx = full.find("делегация")
            replace_para_text(para, full[:idx] + "делегация  " + company_full)
            continue
        if "Провести" in full and "провести официальный прием" in full:
            new_text = re.sub(r"Провести\s+\d{1,2}\s+\w+\s+\d{4}\s+г\.",
                            "Провести " + date_russian, full)
            replace_para_text(para, new_text)
            continue
        if "смету представительских расходов" in full and "размере" in full:
            new_text = re.sub(r"размере\s+.+?\s+рублей\s+\d{2}\s+копеек",
                            "размере " + actual_words + " рублей 00 копеек", full)
            replace_para_text(para, new_text)
            continue
        if "Ответственному за представительское мероприятие работнику" in full:
            idx = full.find("работнику")
            replace_para_text(para, full[:idx] + "работнику " + responsible)
            continue
        if full.strip().startswith("-") and ("менеджер" in full.lower() or "Региональный" in full):
            replace_para_text(para, "- " + position + "  " + participant)
            continue
        if "/________________" in full and "(" in full:
            new_text = re.sub(r"\([^)]*\)", "(" + participant + ")", full)
            replace_para_text(para, new_text)
            continue
    doc.save(out_p)

    # АктОтчет
    wb = load_workbook(out_a)
    ws = wb.active
    ws["D11"] = int(day); ws["G11"] = month; ws["J11"] = int(year)
    ws["B18"] = company_full
    ws["M22"] = int(day); ws["P22"] = month; ws["S22"] = int(year)
    ws["K23"] = venue; ws["K24"] = address
    for r in range(28, 74):
        ws["B" + str(r)] = None; ws["Y" + str(r)] = None
    for i in range(18):
        pk = "deleg_pos_" + str(i)
        nk = "deleg_nam_" + str(i)
        if pk in data and nk in data:
            pos = data[pk].strip()
            nam = data[nk].strip()
            if pos and nam:
                row = 28 + i * 2
                ws["B" + str(row)] = pos
                ws["Y" + str(row)] = nam
                ws["B" + str(row+1)] = "(должность)"
                ws["Y" + str(row+1)] = "(ФИО)"
    ws["B77"] = position; ws["Y77"] = participant
    ws["B81"] = purpose; ws["B84"] = result
    ws["N89"] = actual_amt; ws["S89"] = actual_words; ws["AE89"] = 0
    ws["Q94"] = rec_date_clean
    ws["U94"] = int(receipt_num) if str(receipt_num).isdigit() else receipt_num
    ws["Y94"] = receipt_amt
    for r in range(99, 107):
        ws["I" + str(r)] = None; ws["AC" + str(r)] = None
    if comm1_pos: ws["I99"] = comm1_pos
    if comm1_name: ws["AC99"] = comm1_name
    if comm2_pos: ws["I101"] = comm2_pos
    if comm2_name: ws["AC101"] = comm2_name
    if comm3_pos: ws["I103"] = comm3_pos
    if comm3_name: ws["AC103"] = comm3_name
    if compiler_pos: ws["I105"] = compiler_pos
    if compiler_name: ws["AC105"] = compiler_name
    wb.save(out_a)
    return out_s, out_p, out_a

st.set_page_config(page_title="报销生成器", page_icon="📋", layout="wide")
st.title("📋 柳工报销文档生成器")
st.caption("填写表单 → 一键下载三份文件")

with st.form("form"):
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("基本信息")
        date_str = st.text_input("活动日期 *", "25.05.2026")
        actual_amt = st.number_input("实际报销金额 *", 0, value=22370)
        budget_amt = st.number_input("Смета预算金额 *", 0, value=30000, help="比实际大且取整")
        amount_manual = st.text_input("金额大写(不填=自动)", "")
        st.subheader("公司与人员")
        company_full = st.text_input("对方公司全名 *", "ЗК Урюм")
        responsible = st.text_input("负责人 *", "Иван Ли")
        position = st.text_input("参与人职位 *", "Менеджер По Продажам")
        participant = st.text_input("参与人姓名 *", "Иван Ли")
    with c2:
        st.subheader("会议地点")
        venue = st.text_input("地点名称 *", "чуаньюй")
        address = st.text_input("地址 *", "МОСКВА, ПРОСПЕКТ.МИЧУРИНСКИЙ, ДОМ 7")
        st.subheader("会议内容")
        purpose = st.text_input("目的", "Продаж наши машины")
        result = st.text_input("成果", "Судим по проектом")
        st.subheader("小票")
        receipt_date = st.text_input("小票日期", "2026-05-24")
        receipt_num = st.text_input("小票编号", "")
        receipt_amt = st.number_input("小票金额(不填=实际)", 0, value=0)

    st.subheader("对方代表团 (每人2300 RUB)")
    nd = st.number_input("人数", 1, 18, 5)
    dc1, dc2 = st.columns(2)
    dpos, dnam = [], []
    for i in range(nd):
        with dc1: dpos.append(st.text_input("职位" + str(i+1), key="dp"+str(i)))
        with dc2: dnam.append(st.text_input("姓名" + str(i+1), key="dn"+str(i)))

    st.subheader("Комиссия 成员")
    cm1, cm2 = st.columns(2)
    with cm1:
        comm1_pos = st.text_input("成员1 职位(I99)", "Генеральный директор", key="c1p")
        comm1_name = st.text_input("成员1 姓名(AC99)", "Сяо Юаньсян", key="c1n")
    with cm2:
        comm2_pos = st.text_input("成员2 职位(I101)", "", key="c2p")
        comm2_name = st.text_input("成员2 姓名(AC101)", "", key="c2n")
    cm3, cm4 = st.columns(2)
    with cm3:
        comm3_pos = st.text_input("成员3 职位(I103)", "", key="c3p")
        comm3_name = st.text_input("成员3 姓名(AC103)", "", key="c3n")
    with cm4:
        compiler_pos = st.text_input("编制人 职位(I105)", "менеджер по продажам", key="cmp_p")
        compiler_name = st.text_input("编制人 姓名(AC105)", "Иван Ли", key="cmp_n")

    company_short = st.text_input("对方公司简称(不填=全名)", "")

    go = st.form_submit_button("生成报销文件", type="primary", use_container_width=True)

if go:
    data = {
        "date": date_str, "actual_amt": actual_amt, "budget_amt": budget_amt,
        "amount_manual": amount_manual, "company_full": company_full,
        "responsible": responsible, "position": position, "participant": participant,
        "venue": venue, "address": address, "purpose": purpose, "result": result,
        "receipt_date": receipt_date, "receipt_num": receipt_num,
        "receipt_amt": receipt_amt if receipt_amt else actual_amt,
        "company_short": company_short,
        "comm1_pos": comm1_pos, "comm1_name": comm1_name,
        "comm2_pos": comm2_pos, "comm2_name": comm2_name,
        "comm3_pos": comm3_pos, "comm3_name": comm3_name,
        "compiler_pos": compiler_pos, "compiler_name": compiler_name,
    }
    for i,(p,n) in enumerate(zip(dpos,dnam)):
        data["deleg_pos_"+str(i)] = p
        data["deleg_nam_"+str(i)] = n

    required = {"活动日期": date_str, "实际报销金额": actual_amt, "对方公司全名": company_full,
                "负责人": responsible, "参与人职位": position, "参与人姓名": participant}
    missing = [k for k,v in required.items() if not v]
    if missing:
        st.error("请填写: " + ", ".join(missing))
    else:
        with st.spinner("正在生成..."):
            td = tempfile.mkdtemp()
            try:
                s,p,a = generate_files(data, td)
                st.success("生成完成! 预算:" + str(budget_amt) + " RUB | 实际:" + str(actual_amt) + " RUB")
                with open(s,"rb") as f: sb = f.read()
                with open(p,"rb") as f: pb = f.read()
                with open(a,"rb") as f: ab = f.read()
                zb = io.BytesIO()
                with zipfile.ZipFile(zb,"w") as zf:
                    zf.writestr("Смета.docx", sb)
                    zf.writestr("报销_приказ.docx", pb)
                    zf.writestr("АктОтчет.xlsx", ab)
                zb.seek(0)
                st.download_button("下载全部 (ZIP)", zb, "报销文件.zip", "application/zip", use_container_width=True)
                cld, crd = st.columns(2)
                with cld: st.download_button("Смета.docx", sb, "Смета.docx")
                with crd: st.download_button("报销_приказ.docx", pb, "报销_приказ.docx")
                st.download_button("АктОтчет.xlsx", ab, "АктОтчет.xlsx")
            finally:
                shutil.rmtree(td, ignore_errors=True)
