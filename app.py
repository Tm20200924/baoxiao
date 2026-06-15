# -*- coding: utf-8 -*-
import streamlit as st, os, io, zipfile, tempfile, re, datetime, shutil, random
from pathlib import Path
from docx import Document
from openpyxl import load_workbook

TMPL_DIR = Path(__file__).parent / "templates"

# ====== Russian number-to-words ======
def rub_to_words(n):
    if n == 0: return "ноль"
    u = ["","один","два","три","четыре","пять","шесть","семь","восемь","девять",
         "десять","одиннадцать","двенадцать","тринадцать","четырнадцать","пятнадцать",
         "шестнадцать","семнадцать","восемнадцать","девятнадцать"]
    t = ["","","двадцать","тридцать","сорок","пятьдесят","шестьдесят","семьдесят","восемьдесят","девяносто"]
    h = ["","сто","двести","триста","четыреста","пятьсот","шестьсот","семьсот","восемьсот","девятьсот"]
    uf = ["","одна","две","три","четыре","пять","шесть","семь","восемь","девять"]
    def _x(n,f=False):
        if n==0: return ""
        r=[]; hh=n//100
        if hh: r.append(h[hh])
        rem=n%100
        if rem<20:
            if rem>0: r.append(uf[rem] if (f and rem in(1,2)) else u[rem])
        else:
            r.append(t[rem//10])
            if rem%10>0: r.append(uf[rem%10] if (f and rem%10 in(1,2)) else u[rem%10])
        return " ".join(r)
    def _p(n,a,b,c):
        if 11<=n%100<=19: return c
        l=n%10
        if l==1: return a
        if 2<=l<=4: return b
        return c
    th=n//1000; rm=n%1000; p=[]
    if th>0:
        tx=_x(th,True)
        if tx: p.append(tx+" "+_p(th,"тысяча","тысячи","тысяч"))
    if rm>0 or not p:
        tx=_x(rm)
        if tx: p.append(tx)
    return " ".join(p)

def date_to_russian(s):
    months=["","января","февраля","марта","апреля","мая","июня","июля","августа","сентября","октября","ноября","декабря"]
    parts=s.strip().split(".")
    return f"{int(parts[0])} {months[int(parts[1])]} {parts[2]} г." if len(parts)==3 else s

# ====== Name banks ======
RU_SURNAMES = ["Иванов","Петров","Сидоров","Смирнов","Кузнецов","Попов","Васильев","Михайлов","Новиков","Федоров","Морозов","Волков","Алексеев","Лебедев","Семенов","Егоров","Павлов","Козлов","Степанов","Николаев","Орлов","Андреев","Макаров","Никитин","Захаров","Зайцев","Соловьев","Борисов","Яковлев","Григорьев","Романов","Воробьев","Сергеев","Кузьмин","Фролов","Александров","Дмитриев","Королев","Гусев","Киселев","Ильин","Максимов","Поляков","Сорокин","Виноградов","Ковалев","Белов","Медведев","Антонов","Тарасов","Жуков","Баранов","Филиппов","Комаров","Давыдов","Беляев","Герасимов","Богданов","Осипов","Тимофеев"]
RU_GIVEN = ["Александр","Дмитрий","Сергей","Андрей","Алексей","Михаил","Николай","Владимир","Иван","Павел","Роман","Виктор","Юрий","Денис","Евгений","Олег","Игорь","Анатолий","Вадим","Константин","Максим","Антон","Василий","Борис","Геннадий","Григорий","Даниил","Егор","Илья","Кирилл","Лев","Леонид","Матвей","Никита","Петр","Руслан","Святослав","Семен","Станислав","Степан","Тимур","Федор","Эдуард","Ярослав","Артем","Валерий","Владислав","Георгий","Арсений","Валентин","Вячеслав","Тимофей","Всеволод","Марк","Филипп"]
CN_SURNAMES_RU = ["Чжан","Ван","Ли","Чжао","Лю","Чэнь","Ян","Хуан","Чжоу","У","Сюй","Сунь","Ма","Чжу","Ху","Го","Хэ","Гао","Линь","Ло","Чжэн","Лян","Се","Сун","Тан","Сюй","Хань","Фэн","Дэн","Цао","Пэн","Цзэн","Сяо","Тянь","Дун","Пань","Юань","Юй","Цзян","Цай","Юй","Ду","Е","Чэн","Су","Вэй","Люй","Дин","Жэнь","Шэнь","Яо","Лу","Цзян","Цуй","Чжун","Тань","Лу","Ван","Фань","Ши"]
CN_GIVEN_RU = ["Вэй","Лэй","Ян","Юн","Цзюнь","Цзе","Тао","Мин","Чао","Цян","Пэн","Цзяньхуа","Годун","Чжицян","Цзяньго","Вэньбо","Хаожань","Цзыхань","Юйсюань","Юйцзэ","Чжиюань","Сыюань","Бовэнь","Цзюньцзе","Жуй","Чэнь","Ян","Фэн","Нин","Лун","Фэй","Бо","Бинь","Ган","Хуэй","Линь","Минь","Пин","Лян","Синь","И","Сюй","Хао","Сян","Чжэ","Хэн","Юэ","Жань","И","Хань","Цзэ","Жуй","Ань","Кан","Цзянь"]
RU_POSITIONS = ["Генеральный директор","Исполнительный директор","Финансовый директор","Главный бухгалтер","Руководитель отдела продаж","Руководитель транспортного отдела","Специалист по таможенному оформлению","Менеджер по маркетингу","Менеджер по закупкам","Главный инженер","Инженер-проектировщик","Специалист по логистике","Юрисконсульт","Экономист","Специалист по ВЭД","IT-специалист","Начальник склада","Водитель-экспедитор","Переводчик (китайский язык)","Офис-менеджер","Ассистент проекта","Аналитик","Технический директор","Менеджер по продукту","Специалист по безопасности"]

# ====== Helpers ======
def _pfull(para): return "".join(r.text for r in para.runs)
def _rpt(para, txt):
    for i,r in enumerate(para.runs): r.text = txt if i==0 else ""
def _sct(cell, txt):
    ps=cell.paragraphs
    if ps:
        rs=ps[0].runs
        if rs:
            rs[0].text=str(txt)
            for r in rs[1:]: r.text=""
        else: ps[0].text=str(txt)

# ====== Generate files ======
def generate(data, out_dir):
    # Unpack with longer names to avoid shadowing
    date_str = data["date"]
    actual_amt = data["actual_amt"]
    budget_amt = data.get("budget_amt") or actual_amt
    amt_manual = data.get("amount_manual","").strip()
    company_full = data["company_full"]
    responsible = data["responsible"]
    position = data["position"]
    participant = data["participant"]
    venue = data.get("venue","")
    address = data.get("address","")
    purpose = data.get("purpose","")
    result = data.get("result","")
    receipt_date = data.get("receipt_date","")
    receipt_num = data.get("receipt_num","")
    receipt_amt = data.get("receipt_amt") or actual_amt
    company_short = data.get("company_short") or company_full

    budget_words = amt_manual if amt_manual else rub_to_words(budget_amt)
    actual_words = rub_to_words(actual_amt)
    date_russian = date_to_russian(date_str)
    day_str, month_str, year_str = date_str.split(".")
    day_int, month_int, year_int = int(day_str), month_str, int(year_str)

    # Parse receipt date to dd.mm.yyyy
    rec_date_clean = receipt_date
    for fmt in ["%Y-%m-%d %H:%M:%S","%Y-%m-%d"]:
        try:
            dt = datetime.datetime.strptime(receipt_date, fmt)
            rec_date_clean = dt.strftime("%d.%m.%Y")
            break
        except: pass

    os.makedirs(out_dir, exist_ok=True)
    path_smeta = os.path.join(out_dir, "Смета.docx")
    path_prikaz = os.path.join(out_dir, "报销_приказ.docx")
    path_akt = os.path.join(out_dir, "АктОтчет.xlsx")

    shutil.copy2(TMPL_DIR/"Смета_шаблон.docx", path_smeta)
    shutil.copy2(TMPL_DIR/"приказ_шаблон.docx", path_prikaz)
    wb = load_workbook(TMPL_DIR/"АктОтчет_шаблон.xlsx")
    wb.save(path_akt)

    # --- Смета ---
    doc = Document(path_smeta)
    _sct(doc.tables[0].rows[0].cells[1], date_str)
    _sct(doc.tables[1].rows[1].cells[1], str(budget_amt))
    _sct(doc.tables[1].rows[2].cells[1], str(budget_amt))
    doc.save(path_smeta)

    # --- приказ ---
    doc = Document(path_prikaz)
    DATE_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})")
    for para in doc.paragraphs:
        full = _pfull(para)
        if "Красногорск" in full and DATE_RE.search(full):
            _rpt(para, DATE_RE.sub(date_str + " г.", full)); continue
        if DATE_RE.match(full.strip().rstrip("г.").strip()):
            _rpt(para, date_str + " г."); continue
        if "В целях установления" in full and "делегация" in full:
            ix = full.find("делегация")
            _rpt(para, full[:ix] + "делегация  " + company_full); continue
        if "Провести" in full and "провести официальный прием" in full:
            _rpt(para, re.sub(r"Провести\s+\d{1,2}\s+\w+\s+\d{4}\s+г\.","Провести " + date_russian, full)); continue
        if "смету представительских расходов" in full and "размере" in full:
            _rpt(para, re.sub(r"размере\s+.+?\s+рублей\s+\d{2}\s+копеек","размере " + actual_words + " рублей 00 копеек", full)); continue
        if "Ответственному за представительское мероприятие работнику" in full:
            ix = full.find("работнику")
            _rpt(para, full[:ix] + "работнику " + responsible); continue
        if full.strip().startswith("-") and ("менеджер" in full.lower() or "Региональный" in full):
            _rpt(para, "- " + position + "  " + participant); continue
        if "/________________" in full and "(" in full:
            _rpt(para, re.sub(r"\([^)]*\)","(" + participant + ")", full)); continue
    doc.save(path_prikaz)

    # --- АктОтчет ---
    wb = load_workbook(path_akt)
    ws = wb.active
    ws["D11"] = day_int; ws["G11"] = month_int; ws["J11"] = year_int
    ws["B18"] = company_full
    ws["M22"] = day_int; ws["P22"] = month_int; ws["S22"] = year_int
    ws["K23"] = venue; ws["K24"] = address

    for r in range(28, 74):
        ws["B" + str(r)] = None; ws["Y" + str(r)] = None
    for i in range(18):
        pk = "dp" + str(i); nk = "dn" + str(i)
        if pk in data and nk in data:
            pos = data[pk].strip(); nam = data[nk].strip()
            if pos and nam:
                row = 28 + i * 2
                ws["B" + str(row)] = pos; ws["Y" + str(row)] = nam
                ws["B" + str(row+1)] = "(должность)"; ws["Y" + str(row+1)] = "(ФИО)"

    ws["B77"] = position; ws["Y77"] = participant
    ws["B81"] = purpose; ws["B84"] = result
    ws["N89"] = actual_amt; ws["S89"] = actual_words; ws["AE89"] = 0
    ws["Q94"] = rec_date_clean
    ws["U94"] = int(receipt_num) if str(receipt_num).isdigit() else receipt_num
    ws["Y94"] = receipt_amt

    for r in range(99, 107):
        ws["I" + str(r)] = None; ws["AC" + str(r)] = None
    if data.get("c1p",""): ws["I99"] = data["c1p"]
    if data.get("c1n",""): ws["AC99"] = data["c1n"]
    if data.get("c2p",""): ws["I101"] = data["c2p"]
    if data.get("c2n",""): ws["AC101"] = data["c2n"]
    if data.get("c3p",""): ws["I103"] = data["c3p"]
    if data.get("c3n",""): ws["AC103"] = data["c3n"]
    if data.get("cmp_p",""): ws["I105"] = data["cmp_p"]
    if data.get("cmp_n",""): ws["AC105"] = data["cmp_n"]
    wb.save(path_akt)

    return path_smeta, path_prikaz, path_akt

# ====== UI ======
st.set_page_config(page_title="效率报销-值得拥有", page_icon="📋", layout="wide")
st.title("📋 效率报销-值得拥有")

c1, c2 = st.columns(2)
with c1:
    st.subheader("📌 基本信息 (三文件共用)")
    date_str = st.text_input("活动日期 *", "25.05.2026")
    actual_amt = st.number_input("实际报销金额 *", 0, value=22370)
    budget_amt = st.number_input("Смета预算金额 *", 0, value=30000, help="≥实际金额且取整")
    amount_manual = st.text_input("金额大写(不填自动)", "")
    company_type = st.selectbox("公司类型", ["russian","chinese"],
                                format_func=lambda x: "🇷🇺 俄罗斯" if x=="russian" else "🇨🇳 中国(俄语转写)")
    company_full = st.text_input("对方公司全名 *", "ЗК Урюм")
    submitter = st.text_input("提交人", "")
    responsible = st.text_input("负责人 *", "Иван Ли")
    position = st.text_input("参与人职位 *", "Менеджер По Продажам")
    participant = st.text_input("参与人姓名 *", "Иван Ли")
    venue = st.text_input("会议地点 *", "чуаньюй")
    address = st.text_input("会议地址 *", "МОСКВА, ПРОСПЕКТ.МИЧУРИНСКИЙ, ДОМ 7")

with c2:
    st.subheader("📝 会议")
    purpose = st.text_input("目的", "Продаж наши машины")
    result = st.text_input("成果", "Судим по проектом")
    st.subheader("🧾 小票")
    receipt_date = st.text_input("日期", "2026-05-24")
    receipt_num = st.text_input("编号", "")
    receipt_amt = st.number_input("金额(不填=实际)", 0, value=0)
    st.subheader("👤 Комиссия")
    c1p = st.text_input("Ген.директор", "Генеральный директор", disabled=True, key="c1p")
    c1n = st.text_input("姓名", "Сяо Юаньсян", disabled=True, key="c1n")
    c2p = st.text_input("上级领导 职位", "", key="c2p")
    c2n = st.text_input("上级领导 姓名", "", key="c2n")
    c3p = st.text_input("主会计 职位", "Главный бухгалтер", key="c3p")
    c3n = st.text_input("主会计 姓名", "", key="c3n")
    cmp_p = st.text_input("编制人 职位", "менеджер по продажам", key="cmp_p")
    cmp_n = st.text_input("编制人 姓名", participant, key="cmp_n")
    company_short = st.text_input("公司简称(不填=全名)", "")

# === Delegation ===
st.subheader("👥 对方代表团 (每人2300₽)")
delegate_count = max(1, (actual_amt + 2299) // 2300)
st.info(f"💰 {actual_amt}₽ ÷ 2300 = **{delegate_count}人** / 预算 **{delegate_count*2300}**₽")

gcol, clcol = st.columns([1, 4])
with gcol:
    if st.button("🎲 随机生成", use_container_width=True):
        if "delegation" not in st.session_state: st.session_state.delegation = []
        shuf = list(RU_POSITIONS); random.shuffle(shuf)
        st.session_state.delegation = []
        for i in range(delegate_count):
            pos = shuf[i % len(shuf)]
            if company_type == "russian":
                name = random.choice(RU_SURNAMES) + " " + random.choice(RU_GIVEN)
            else:
                name = random.choice(CN_SURNAMES_RU) + " " + random.choice(CN_GIVEN_RU)
            st.session_state.delegation.append((pos, name))
        st.rerun()

with clcol:
    if st.button("🗑️ 清空", use_container_width=True):
        if "delegation" in st.session_state: st.session_state.delegation = []
        st.rerun()

if "delegation" not in st.session_state: st.session_state.delegation = []

dpos, dnam = [], []
if st.session_state.delegation:
    for i, (pos, name) in enumerate(st.session_state.delegation):
        x1, x2 = st.columns([3, 3])
        with x1: dpos.append(st.text_input("职位", pos, key="gdp"+str(i), label_visibility="collapsed"))
        with x2: dnam.append(st.text_input("姓名", name, key="gdn"+str(i), label_visibility="collapsed"))
else:
    for i in range(delegate_count):
        x1, x2 = st.columns([3, 3])
        with x1: dpos.append(st.text_input("职位"+str(i+1), key="mdp"+str(i)))
        with x2: dnam.append(st.text_input("姓名"+str(i+1), key="mdn"+str(i)))
while len(dpos) < MAX_DELEGATES_UI: dpos.append("")
while len(dnam) < 50: dnam.append("")

# === Submit ===
if st.button("🚀 生成报销文件", type="primary", use_container_width=True):
    missing = [k for k, v in {"活动日期": date_str, "实际金额": actual_amt, "公司全名": company_full,
                               "负责人": responsible, "职位": position, "姓名": participant}.items() if not v]
    if missing:
        st.error("请填写: " + ", ".join(missing))
    else:
        with st.spinner("生成中..."):
            data = {
                "date": date_str, "actual_amt": actual_amt, "budget_amt": budget_amt,
                "amount_manual": amount_manual, "company_full": company_full,
                "responsible": responsible, "position": position, "participant": participant,
                "venue": venue, "address": address, "purpose": purpose, "result": result,
                "receipt_date": receipt_date, "receipt_num": receipt_num,
                "receipt_amt": receipt_amt if receipt_amt else actual_amt,
                "company_short": company_short,
                "c1p": "Генеральный директор", "c1n": "Сяо Юаньсян",
                "c2p": c2p, "c2n": c2n, "c3p": c3p, "c3n": c3n,
                "cmp_p": cmp_p, "cmp_n": cmp_n,
            }
            for i, (p, n) in enumerate(zip(dpos, dnam)):
                data["dp"+str(i)] = p; data["dn"+str(i)] = n

            td = tempfile.mkdtemp()
            try:
                sm, pk, ak = generate(data, td)
                st.success(f"✅ 预算:{budget_amt}₽ | 实际:{actual_amt}₽")
                with open(sm,"rb") as f: sb = f.read()
                with open(pk,"rb") as f: pb = f.read()
                with open(ak,"rb") as f: ab = f.read()
                zb = io.BytesIO()
                with zipfile.ZipFile(zb, "w") as zf:
                    zf.writestr("Смета.docx", sb)
                    zf.writestr("报销_приказ.docx", pb)
                    zf.writestr("АктОтчет.xlsx", ab)
                zb.seek(0)
                st.download_button("📥 下载全部 ZIP", zb, "报销文件.zip", "application/zip", use_container_width=True)
                dx1, dx2 = st.columns(2)
                with dx1: st.download_button("📄 Смета", sb, "Смета.docx")
                with dx2: st.download_button("📄 Приказ", pb, "报销_приказ.docx")
                st.download_button("📊 АктОтчет", ab, "АктОтчет.xlsx")
            finally:
                shutil.rmtree(td, ignore_errors=True)
