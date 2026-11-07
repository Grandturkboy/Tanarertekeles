import random
import pandas as pd
from docx import Document

SCALE_QUESTIONS = {
    1: "Tananyag érthetősége",
    2: "Felkészültség",
    3: "Értékelési szempontok világossága",
    4: "Pontosság (kezdés/befejezés)",
    5: "Kommunikáció, elérhetőség",
    6: "Segítség kérdés esetén",
    7: "Tantárgy teljesíthetősége",
    8: "Motiváció",
    9: "Összesített jegy",
}
FINAL_QS = list(range(1, 10))   # questions averaged into the final rating; remove 9 if you don't want it included
COMMENT_Q = 10

path = r"OsztályPerTanár - responses.csv" # <== Ez legyen a kiexportált .csv fájl neve
df = pd.read_csv(path, dtype=str, keep_default_na=False)
df["question_no"] = df["question_no"].astype(int)
df["timestamp"] = pd.to_datetime(df["timestamp"])

# One response per student per class; responses without an email are kept as they are
resp = df[["class", "response_id", "email", "timestamp"]].drop_duplicates("response_id").copy()
resp["email"] = resp["email"].str.strip().str.lower()

has_email = resp["email"] != ""
deduped = resp[has_email].sort_values("timestamp").drop_duplicates(["class", "email"], keep="last")
keep = pd.concat([deduped, resp[~has_email]])

df = df[df["response_id"].isin(keep["response_id"])]

scores = df[df["question_no"].isin(SCALE_QUESTIONS)].copy()
scores["answer"] = pd.to_numeric(scores["answer"])

df = df.drop_duplicates(["response_id", "teacher", "question_no"])   # in case a killed export wrote a form twice
assert df["question_no"].between(1, 10).all(), "unexpected question number"
assert scores["answer"].between(1, 5).all(), "score outside 1-5"

comments = df[(df["question_no"] == COMMENT_Q) & (df["answer"].str.strip() != "")]

# Mean per teacher and question, per class and overall
per_class = scores.pivot_table(index=["teacher", "question_no"], columns="class",
                               values="answer", aggfunc="mean")
overall = scores.groupby(["teacher", "question_no"])["answer"].mean().rename("Összesen")
table = per_class.join(overall)

# Number of students per teacher, per class and overall
n_class = scores.groupby(["teacher", "class"])["response_id"].nunique().unstack("class")
n_total = scores.groupby("teacher")["response_id"].nunique()

def fmt(x):
    return "" if pd.isna(x) else f"{x:.2f}"

# ---- Ratings document ----
doc = Document()
doc.add_heading("Tanárértékelés – összesítés", 0)

for teacher in sorted(n_total.index):
    classes = list(n_class.loc[teacher].dropna().index)
    cols = classes + ["Összesen"]
    t = table.loc[teacher].reindex(list(SCALE_QUESTIONS)).reindex(columns=cols)
    final = t.loc[FINAL_QS].mean()

    doc.add_heading(teacher, 1)
    tbl = doc.add_table(rows=1, cols=len(cols) + 1)
    tbl.style = "Light Grid Accent 1"
    tbl.rows[0].cells[0].text = "Kérdés"
    for j, c in enumerate(cols, 1):
        tbl.rows[0].cells[j].text = c

    for q, label in SCALE_QUESTIONS.items():
        cells = tbl.add_row().cells
        cells[0].text = f"{q}. {label}"
        for j, c in enumerate(cols, 1):
            cells[j].text = fmt(t.loc[q, c])

    cells = tbl.add_row().cells
    cells[0].text = "Végső értékelés"
    for j, c in enumerate(cols, 1):
        cells[j].text = fmt(final[c])
        cells[j].paragraphs[0].runs[0].bold = True
    cells[0].paragraphs[0].runs[0].bold = True

    cells = tbl.add_row().cells
    cells[0].text = "Értékelők száma"
    for j, c in enumerate(classes, 1):
        cells[j].text = str(int(n_class.loc[teacher, c]))
    cells[len(cols)].text = str(int(n_total[teacher]))

doc.save("ertekelesek.docx")

# ---- Written feedback document (no names, classes or timestamps; shuffled) ----
fb = Document()
fb.add_heading("Tanárértékelés – szöveges visszajelzések", 0)

first = True
for teacher, g in comments.groupby("teacher"):
    items = g["answer"].str.strip().tolist()
    random.shuffle(items)

    heading = fb.add_heading(teacher, 1)
    if not first:
        heading.paragraph_format.page_break_before = True
    first = False

    for c in items:
        fb.add_paragraph(c, style="List Bullet")

fb.save("szoveges_visszajelzesek.docx")

# ---- Plain CSV summary, handy for sorting or checking ----
summary = overall.unstack("question_no").reindex(columns=list(SCALE_QUESTIONS))
summary["final"] = summary[FINAL_QS].mean(axis=1)
summary["n_students"] = n_total
summary.round(2).to_csv("ratings.csv")
