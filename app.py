import streamlit as st
import pandas as pd
import json
import shutil
from datetime import datetime
from pathlib import Path
from io import BytesIO
from docx import Document
from docx.shared import Pt

# ── 設定 ──────────────────────────────────────────────
DATA_FILE        = Path("data/submissions.json")
TEMPLATE_FILE    = Path("rirekisho_template.docx")
ADMIN_PASS       = "admin1234"
MAX_COMPANIES    = 5

st.set_page_config(page_title="経歴入力フォーム", page_icon="📄", layout="centered")

# ── データ保存 ────────────────────────────────────────
def load_data():
    DATA_FILE.parent.mkdir(exist_ok=True)
    if DATA_FILE.exists():
        return json.loads(DATA_FILE.read_text(encoding="utf-8"))
    return []

def save_data(records):
    DATA_FILE.parent.mkdir(exist_ok=True)
    DATA_FILE.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")

# ── セルへの書き込み ──────────────────────────────────
def set_cell(cell, text):
    """セルのテキストを書き込む（既存の書式をリセット）"""
    for para in cell.paragraphs:
        for run in para.runs:
            run.text = ""
    if cell.paragraphs:
        run = cell.paragraphs[0].add_run(str(text) if text else "")
        run.font.size = Pt(10.5)
    else:
        cell.text = str(text) if text else ""

# ── 履歴書テンプレートに流し込む ─────────────────────
def make_rirekisho(d):
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError("rirekisho_template.docx が見つかりません。resume_toolフォルダに入れてください。")

    # テンプレートをメモリ上にコピー
    buf = BytesIO()
    buf.write(TEMPLATE_FILE.read_bytes())
    buf.seek(0)
    doc = Document(buf)

    edu_list = d.get("education", [])
    car_list  = d.get("career",    [])
    lic_list  = d.get("licenses",  [])

    # ── テーブル0：基本情報 ───────────────────────────
    t0 = doc.tables[0]
    set_cell(t0.rows[0].cells[1], d.get("furigana_name",""))
    set_cell(t0.rows[1].cells[1], d.get("name",""))
    set_cell(t0.rows[2].cells[1],
             f"{d.get('birthday','')}（満{d.get('age','')}歳）")
    set_cell(t0.rows[3].cells[1], d.get("furigana_address",""))
    set_cell(t0.rows[4].cells[1],
             f"〒{d.get('postal','')}　{d.get('address','')}")
    set_cell(t0.rows[5].cells[1], d.get("phone",""))
    set_cell(t0.rows[5].cells[5], d.get("email",""))

    # ── テーブル1：学歴・職歴 ────────────────────────
    t1 = doc.tables[1]
    for i, edu in enumerate(edu_list[:6]):
        r = i + 2
        if r < len(t1.rows):
            set_cell(t1.rows[r].cells[0], edu.get("year",""))
            set_cell(t1.rows[r].cells[1], edu.get("month",""))
            set_cell(t1.rows[r].cells[2], edu.get("content",""))

    for i, job in enumerate(car_list[:9]):
        r = i + 9
        if r < len(t1.rows):
            set_cell(t1.rows[r].cells[0], job.get("year",""))
            set_cell(t1.rows[r].cells[1], job.get("month",""))
            set_cell(t1.rows[r].cells[2], job.get("content",""))

    # ── テーブル2：免許・資格 ────────────────────────
    t2 = doc.tables[2]
    for i, lic in enumerate(lic_list[:3]):
        r = i + 1
        if r < len(t2.rows):
            set_cell(t2.rows[r].cells[0], lic.get("year",""))
            set_cell(t2.rows[r].cells[1], lic.get("month",""))
            set_cell(t2.rows[r].cells[2], lic.get("content",""))

    # ── テーブル3：通勤・家族情報 ────────────────────
    t3 = doc.tables[3]
    set_cell(t3.rows[1].cells[1], d.get("nearest_station",""))
    set_cell(t3.rows[1].cells[2],
             f"（配偶者を除く）{d.get('dependents','0')}人")
    set_cell(t3.rows[1].cells[3], d.get("spouse","無"))
    set_cell(t3.rows[1].cells[4], d.get("spouse_support","無"))

    # ── テーブル4：自己PR ────────────────────────────
    t4 = doc.tables[4]
    set_cell(t4.rows[1].cells[0], d.get("pr",""))

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# ── 職務経歴書（テンプレート流し込み）──────────────
SHOKUMU_TEMPLATE = Path("shokumu_template.docx")

def make_shokumu(d):
    if not SHOKUMU_TEMPLATE.exists():
        raise FileNotFoundError("shokumu_template.docx が見つかりません。resume_toolフォルダに入れてください。")

    buf_in = BytesIO(SHOKUMU_TEMPLATE.read_bytes())
    doc = Document(buf_in)
    paras = doc.paragraphs

    def set_run(para, text):
        for run in para.runs:
            run.text = ""
        if para.runs:
            para.runs[0].text = str(text) if text else ""
        else:
            para.add_run(str(text) if text else "")

    def add_bold(text):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10.5)

    def add_normal(text):
        p = doc.add_paragraph()
        p.add_run(str(text) if text else "")

    # 日付・氏名
    if paras[2].runs:
        paras[2].runs[0].text = f"{datetime.now().strftime('%Y年%m月%d日')}現在"
    set_run(paras[3], f"氏名　{d.get('name','')}")

    # 職務要約
    set_run(paras[5], d.get("summary",""))

    # スキル
    set_run(paras[7], d.get("skills",""))

    # 自己PR
    if len(paras) > 18:
        set_run(paras[18], d.get("pr",""))

    # 1社目
    companies = d.get("companies", [])
    if companies:
        c = companies[0]
        set_run(paras[9],  f"（１）{c.get('company_name','')}")
        set_run(paras[10], f"・事業内容：{c.get('business','')}")
        set_run(paras[11], f"・資本金：{c.get('capital','')}万円")
        set_run(paras[12], f"・従業員数：{c.get('employees','')}名")

        t = doc.tables[0]
        cell_period = t.rows[1].cells[0]
        for para in cell_period.paragraphs:
            for run in para.runs: run.text = ""
        if cell_period.paragraphs:
            cell_period.paragraphs[0].add_run(
                f"{c.get('period_start','')} ～ {c.get('period_end','')}")

        cell_emp = t.rows[1].cells[1]
        for para in cell_emp.paragraphs:
            for run in para.runs: run.text = ""
        if cell_emp.paragraphs:
            cell_emp.paragraphs[0].add_run(
                f"雇用形態：{c.get('employment_type','正社員')}　配属先：{c.get('department','')}")

        cell_content = t.rows[2].cells[1]
        for para in cell_content.paragraphs:
            for run in para.runs: run.text = ""
        if cell_content.paragraphs:
            cell_content.paragraphs[0].add_run(
                f"《業務内容》\n{c.get('job_content','')}\n\n《実績》\n{c.get('achievement','')}")

    # 2社目以降
    for i, c in enumerate(companies[1:], start=2):
        if not c.get("company_name"):
            continue
        add_bold(f"（{i}）{c.get('company_name','')}")
        add_normal(f"・事業内容：{c.get('business','')}")
        add_normal(f"・資本金：{c.get('capital','')}万円")
        add_normal(f"・従業員数：{c.get('employees','')}名")
        add_normal(f"・雇用形態：{c.get('employment_type','正社員')}　配属先：{c.get('department','')}")
        add_normal(f"《業務内容》\n{c.get('job_content','')}")
        add_normal(f"《実績》\n{c.get('achievement','')}")

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# ══════════════════════════════════════════════════════
# サイドバー：モード選択
# ══════════════════════════════════════════════════════
mode = st.sidebar.radio("モード", ["📝 入力フォーム（ユーザー）", "🔐 管理画面"])

# ══════════════════════════════════════════════════════
# ユーザー側：入力フォーム
# ══════════════════════════════════════════════════════
if mode == "📝 入力フォーム（ユーザー）":
    st.title("📄 経歴入力フォーム")
    st.caption("担当アドバイザーから案内されたフォームです。各項目を入力して送信してください。")
    st.markdown("---")

    with st.form("resume_form"):

        st.subheader("① 氏名")
        furigana_name = st.text_input("ふりがな", placeholder="やまだ たろう")
        name = st.text_input("氏名 *", placeholder="山田 太郎")

        st.subheader("② 生年月日")
        col1, col2 = st.columns(2)
        birthday = col1.text_input("生年月日", placeholder="1990年1月1日")
        age      = col2.text_input("年齢", placeholder="35")

        st.subheader("③ 住所")
        postal   = st.text_input("郵便番号", placeholder="123-4567")
        address  = st.text_area("住所", placeholder="東京都新宿区〇〇1-2-3", height=68)
        furigana_address = st.text_input("住所ふりがな", placeholder="とうきょうとしんじゅくく〇〇")

        st.subheader("④ 電話番号")
        phone = st.text_input("電話番号", placeholder="090-1234-5678")

        st.subheader("⑤ メールアドレス")
        email = st.text_input("メールアドレス", placeholder="example@email.com")

        st.subheader("⑥ 学歴（高校から・西暦）")
        st.caption("年・月・内容を入力してください")
        education = []
        cols_h = st.columns([1,1,4])
        cols_h[0].markdown("**年**")
        cols_h[1].markdown("**月**")
        cols_h[2].markdown("**内容**")
        for i in range(6):
            c1,c2,c3 = st.columns([1,1,4])
            y = c1.text_input("年", key=f"edu_y{i}", placeholder="2006", label_visibility="collapsed")
            m = c2.text_input("月", key=f"edu_m{i}", placeholder="3",    label_visibility="collapsed")
            c = c3.text_input("内容", key=f"edu_c{i}", placeholder="〇〇高等学校 卒業", label_visibility="collapsed")
            education.append({"year":y,"month":m,"content":c})

        st.subheader("⑦ 職歴（アルバイトも含む）")
        career = []
        cols_h2 = st.columns([1,1,4])
        cols_h2[0].markdown("**年**")
        cols_h2[1].markdown("**月**")
        cols_h2[2].markdown("**内容**")
        for i in range(8):
            c1,c2,c3 = st.columns([1,1,4])
            y = c1.text_input("年", key=f"car_y{i}", placeholder="2010", label_visibility="collapsed")
            m = c2.text_input("月", key=f"car_m{i}", placeholder="4",    label_visibility="collapsed")
            c = c3.text_input("内容", key=f"car_c{i}", placeholder="〇〇株式会社 入社", label_visibility="collapsed")
            career.append({"year":y,"month":m,"content":c})

        st.subheader("⑧ 免許・資格")
        licenses = []
        cols_h3 = st.columns([1,1,4])
        cols_h3[0].markdown("**年**")
        cols_h3[1].markdown("**月**")
        cols_h3[2].markdown("**内容**")
        for i in range(5):
            c1,c2,c3 = st.columns([1,1,4])
            y = c1.text_input("年", key=f"lic_y{i}", placeholder="2012", label_visibility="collapsed")
            m = c2.text_input("月", key=f"lic_m{i}", placeholder="6",    label_visibility="collapsed")
            c = c3.text_input("内容", key=f"lic_c{i}", placeholder="普通自動車第一種運転免許 取得", label_visibility="collapsed")
            licenses.append({"year":y,"month":m,"content":c})

        st.subheader("⑨ 自己PR")
        st.caption("どんな人物か・何を意識してきたか・どんなふうに頑張れるかを書いてください")
        pr = st.text_area("自己PR *", height=200,
                          placeholder="コミュニケーションを大切にし、チームで目標達成することを意識してきました。")

        st.subheader("⑩ 最寄り駅")
        nearest_station = st.text_input("最寄り駅", placeholder="〇〇線 〇〇駅")

        st.subheader("⑪⑫ 家族情報")
        fc1,fc2,fc3 = st.columns(3)
        dependents     = fc1.text_input("扶養家族（人数）", placeholder="0")
        spouse         = fc2.selectbox("配偶者", ["無","有"])
        spouse_support = fc3.selectbox("配偶者の扶養義務", ["無","有"])

        st.markdown("---")
        st.header("📋 職務経歴書")
        summary = st.text_area("職務要約 *", height=120,
                               placeholder="〇〇業界で△年間の経験を持ちます。")
        skills  = st.text_area("活かせるスキル", height=100,
                               placeholder="・〇〇スキル\n・△△の経験")

        st.subheader("職務経歴（最大5社）")
        companies = []
        for i in range(MAX_COMPANIES):
            with st.expander(f"（{i+1}）会社情報", expanded=(i==0)):
                cname = st.text_input("会社名", key=f"cn{i}", placeholder="〇〇株式会社")
                cc1,cc2,cc3 = st.columns(3)
                ps = cc1.text_input("入社年月", key=f"ps{i}", placeholder="2015年4月")
                pe = cc2.text_input("退社年月", key=f"pe{i}", placeholder="2020年3月")
                py = cc3.text_input("在籍年数", key=f"py{i}", placeholder="5")
                biz  = st.text_input("事業内容", key=f"biz{i}", placeholder="〇〇の製造・販売")
                cc4,cc5 = st.columns(2)
                cap  = cc4.text_input("資本金（万円）", key=f"cap{i}", placeholder="5000")
                emp  = cc5.text_input("従業員数（名）", key=f"emp{i}", placeholder="200")
                etype = st.selectbox("雇用形態", ["正社員","契約社員","アルバイト","派遣","その他"], key=f"et{i}")
                dept  = st.text_input("配属先", key=f"dept{i}", placeholder="営業部 第一課")
                jc    = st.text_area("業務内容", key=f"jc{i}", height=100,
                                     placeholder="・〇〇業務を担当")
                ach   = st.text_area("実績", key=f"ach{i}", height=80,
                                     placeholder="・〇〇を達成し前年比△%向上に貢献")
                companies.append({
                    "company_name":cname,"period_start":ps,"period_end":pe,
                    "period_years":py,"business":biz,"capital":cap,
                    "employees":emp,"employment_type":etype,
                    "department":dept,"job_content":jc,"achievement":ach,
                })

        st.markdown("---")
        submitted = st.form_submit_button("✅ 送信する", type="primary", use_container_width=True)

    if submitted:
        errors = []
        if not name.strip():    errors.append("氏名を入力してください。")
        if not pr.strip():      errors.append("自己PRを入力してください。")
        if not summary.strip(): errors.append("職務要約を入力してください。")
        if errors:
            for e in errors:
                st.error(f"❌ {e}")
        else:
            record = {
                "id": datetime.now().strftime("%Y%m%d%H%M%S"),
                "submitted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "name":name,"furigana_name":furigana_name,
                "birthday":birthday,"age":age,
                "postal":postal,"address":address,"furigana_address":furigana_address,
                "phone":phone,"email":email,
                "education":education,"career":career,"licenses":licenses,
                "pr":pr,"nearest_station":nearest_station,
                "dependents":dependents,"spouse":spouse,"spouse_support":spouse_support,
                "summary":summary,"skills":skills,"companies":companies,
            }
            records = load_data()
            records.append(record)
            save_data(records)
            st.success("✅ 送信が完了しました！担当アドバイザーが確認します。")
            st.balloons()

# ══════════════════════════════════════════════════════
# 管理者側：管理画面
# ══════════════════════════════════════════════════════
else:
    st.title("🔐 管理画面")

    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        pw = st.text_input("パスワード", type="password")
        if st.button("ログイン"):
            if pw == ADMIN_PASS:
                st.session_state.admin_logged_in = True
                st.rerun()
            else:
                st.error("❌ パスワードが違います")
    else:
        st.success("✅ ログイン中")
        if st.button("ログアウト"):
            st.session_state.admin_logged_in = False
            st.rerun()

        records = load_data()
        if not records:
            st.info("まだ送信されたデータがありません。")
        else:
            st.subheader(f"📋 送信一覧（{len(records)}件）")
            for r in reversed(records):
                with st.expander(f"📄 {r.get('name','')}　送信日時：{r.get('submitted_at','')}"):
                    col1,col2 = st.columns(2)

                    # 履歴書（テンプレート流し込み）
                    try:
                        buf_r = make_rirekisho(r)
                        col1.download_button(
                            "⬇️ 履歴書.docxをダウンロード",
                            data=buf_r,
                            file_name=f"履歴書_{r.get('name','')}_{r.get('id','')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"r_{r.get('id','')}",
                            use_container_width=True,
                        )
                    except FileNotFoundError as e:
                        col1.error(str(e))

                    # 職務経歴書
                    buf_s = make_shokumu(r)
                    col2.download_button(
                        "⬇️ 職務経歴書.docxをダウンロード",
                        data=buf_s,
                        file_name=f"職務経歴書_{r.get('name','')}_{r.get('id','')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"s_{r.get('id','')}",
                        use_container_width=True,
                    )

                    st.markdown("---")
                    st.markdown(f"**氏名：** {r.get('name','')}　**生年月日：** {r.get('birthday','')}（{r.get('age','')}歳）")
                    st.markdown(f"**住所：** 〒{r.get('postal','')}　{r.get('address','')}")
                    st.markdown(f"**電話：** {r.get('phone','')}　**メール：** {r.get('email','')}")
                    if r.get('pr'):
                        st.markdown(f"**自己PR：** {r.get('pr','')[:80]}...")
