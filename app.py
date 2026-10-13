
import os
import re
import random
import secrets
from pathlib import Path
from datetime import datetime
from functools import wraps

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file, abort
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from filelock import FileLock

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"

app.secret_key = os.environ.get(
    "SECRET_KEY", secrets.token_hex(32)
)

ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")

EXAM_DURATION = 10 * 60
QUESTION_COUNT = 10
PASS_SCORE = 8

EXAM_TYPES = {
    "ААД": BASE_DIR / "ААД" / "questions.txt",
    "Нарядын систем": BASE_DIR / "Нарядын систем" / "questions.txt",
}

# Render дээр өөрийн тохируулсан хадгалах замыг ашиглаж болно.
DATA_DIR = Path(
    os.environ.get("DATA_DIR", str(BASE_DIR))
)
DATA_DIR.mkdir(parents=True, exist_ok=True)

RESULTS_FILE = DATA_DIR / "results.xlsx"
RESULTS_LOCK = DATA_DIR / "results.xlsx.lock"

# Асуултын TXT файлын формат:
# 1. Асуулт
# A. Сонголт
# B. Сонголт
# C. Сонголт
# D. Сонголт (шаардлагатай бол)
# Зөв хариулт: A


def parse_questions(file_path):
    """TXT файлаас асуултуудыг уншиж, бүтэц болгон хувиргана."""

    if not file_path.exists():
        raise FileNotFoundError(
            f"Асуултын файл олдсонгүй: {file_path}"
        )

    text = file_path.read_text(encoding="utf-8-sig")
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    questions = []
    current = None
    answer_key = None

    def save_question():
        nonlocal current, answer_key

        if not current:
            return

        if current["text"].strip() and len(current["options"]) in (3, 4):
            correct = (answer_key or "").strip().upper()

            # A/B/C/D болон зарим TXT-д хэрэглэдэг
            # "Зөв хариулт: A" форматыг дэмжинэ.
            if correct in ("A", "B", "C", "D"):
                if any(opt["key"] == correct for opt in current["options"]):
                    current["correct"] = correct
                    current["id"] = len(questions) + 1
                    questions.append(current)

        current = None
        answer_key = None

    for raw_line in lines:
        line = raw_line.strip()

        if not line:
            continue

        correct_match = re.match(
            r"^(?:зөв\s*хариулт|зөв\s*хариу|хариулт)\s*[:：\-]?\s*([A-D])\b",
            line,
            re.IGNORECASE
        )

        if correct_match:
            answer_key = correct_match.group(1).upper()
            save_question()
            continue

        # 1. Асуултын текст, 205. Асуултын текст гэх мэт
        question_match = re.match(r"^\d+\s*[.)]\s*(.+)$", line)

        # A. / A) / A - сонголтын формат
        option_match = re.match(
            r"^([A-D])\s*[.)\-:]\s*(.+)$",
            line,
            re.IGNORECASE
        )

        if option_match:
            if current is not None:
                key = option_match.group(1).upper()
                current["options"].append({
                    "key": key,
                    "text": option_match.group(2).strip()
                })
            continue

        if question_match:
            save_question()
            current = {
                "text": question_match.group(1).strip(),
                "options": [],
                "correct": None,
            }
            continue

        # Дугааргүй асуулт байж болох тохиолдол
        if current is None:
            current = {
                "text": line,
                "options": [],
                "correct": None,
            }
        elif not current["options"]:
            current["text"] += "\n" + line

    save_question()

    if len(questions) < QUESTION_COUNT:
        raise ValueError(
            f"{file_path.parent.name} шалгалтад "
            f"шаардлагатай {QUESTION_COUNT} асуулт алга. "
            f"Зөв бүтэцтэй асуулт: {len(questions)}."
        )

    return questions


def init_results_file():
    """Үр дүнгийн Excel файлыг анх удаа үүсгэнэ."""

    if RESULTS_FILE.exists():
        return

    with FileLock(str(RESULTS_LOCK), timeout=15):
        if RESULTS_FILE.exists():
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Шалгалтын дүн"

        headers = [
            "Огноо",
            "Шалгалтын төрөл",
            "Овог нэр",
            "Албан тушаал",
            "Авсан оноо",
            "Нийт оноо",
            "Хувь",
            "Үр дүн",
        ]
        ws.append(headers)

        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                "solid", fgColor="244062"
            )
            cell.alignment = Alignment(
                horizontal="center", vertical="center"
            )

        widths = [22, 24, 28, 28, 14, 14, 14, 18]
        for idx, width in enumerate(widths, start=1):
            ws.column_dimensions[
                chr(64 + idx)
            ].width = width

        wb.save(RESULTS_FILE)
        wb.close()


def save_result(result):
    """Олон хүсэлт зэрэг ирсэн үед Excel бичилтийг түгжинэ."""

    with FileLock(str(RESULTS_LOCK), timeout=20):
        init_results_file()

        wb = load_workbook(RESULTS_FILE)
        ws = wb.active
        ws.append([
            result["date"],
            result["exam_type"],
            result["name"],
            result["position"],
            result["score"],
            result["total"],
            result["percent"],
            result["status"],
        ])

        row = ws.max_row
        for cell in ws[row]:
            cell.alignment = Alignment(vertical="center")

        ws.cell(row, 7).number_format = "0.0%"
        ws.cell(row, 7).value = result["score"] / result["total"]

        wb.save(RESULTS_FILE)
        wb.close()


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))
        return view(*args, **kwargs)
    return wrapped


@app.context_processor
def inject_common():
    return {"exam_open": get_exam_open()}


def get_exam_open():
    # Хаалттай байх нь анхдагч төлөв.
    return app.config.get("EXAM_OPEN", False)


def set_exam_open(value):
    app.config["EXAM_OPEN"] = bool(value)


def create_exam_session(exam_type, name, position):
    all_questions = parse_questions(EXAM_TYPES[exam_type])
    selected = random.SystemRandom().sample(
        all_questions, QUESTION_COUNT
    )

    # Зөв хариултыг зөвхөн серверийн session-д хадгална.
    session["exam"] = {
        "exam_type": exam_type,
        "name": name,
        "position": position,
        "questions": selected,
        "started_at": datetime.now().timestamp(),
        "submitted": False,
    }


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if not get_exam_open():
            flash("Шалгалт одоогоор хаалттай байна.", "warning")
            return redirect(url_for("index"))

        exam_type = request.form.get("exam_type", "").strip()
        name = request.form.get("name", "").strip()
        position = request.form.get("position", "").strip()

        if exam_type not in EXAM_TYPES:
            flash("Шалгалтын төрлөө сонгоно уу.", "danger")
            return redirect(url_for("index"))

        if not name or not position:
            flash("Овог нэр болон албан тушаалаа оруулна уу.", "danger")
            return redirect(url_for("index"))

        try:
            create_exam_session(exam_type, name, position)
        except (OSError, ValueError) as exc:
            flash(f"Асуулт уншихад алдаа гарлаа: {exc}", "danger")
            return redirect(url_for("index"))

        return redirect(url_for("exam"))

    return render_template("index.html", exam_types=list(EXAM_TYPES.keys()))


@app.route("/exam", methods=["GET", "POST"])
def exam():
    exam_data = session.get("exam")

    if not exam_data:
        return redirect(url_for("index"))

    if exam_data.get("submitted"):
        return redirect(url_for("result"))

    elapsed = datetime.now().timestamp() - exam_data["started_at"]
    remaining = max(0, EXAM_DURATION - int(elapsed))

    if request.method == "POST" or remaining <= 0:
        answers = request.form if request.method == "POST" else {}
        score = 0

        for q in exam_data["questions"]:
            selected = answers.get(f"q_{q['id']}", "")
            if selected == q["correct"]:
                score += 1

        result_data = {
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "exam_type": exam_data["exam_type"],
            "name": exam_data["name"],
            "position": exam_data["position"],
            "score": score,
            "total": QUESTION_COUNT,
            "percent": round(score / QUESTION_COUNT * 100, 1),
            "status": "ТЭНЦСЭН" if score >= PASS_SCORE else "ТЭНЦЭЭГҮЙ",
        }

        try:
            save_result(result_data)
        except Exception:
            app.logger.exception("Excel-д үр дүн хадгалахад алдаа гарлаа")
            return (
                "Шалгалтын дүн хадгалагдсангүй. "
                "Админтай холбогдоно уу.",
                500,
            )

        session["result"] = result_data
        session.pop("exam", None)
        return redirect(url_for("result"))

    return render_template(
        "exam.html",
        exam_data=exam_data,
        remaining=remaining,
        total=QUESTION_COUNT,
    )


@app.route("/result")
def result():
    result_data = session.get("result")
    if not result_data:
        return redirect(url_for("index"))

    return render_template("result.html", result=result_data)


@app.route("/admin", methods=["GET", "POST"])
def admin_login():
    if session.get("admin_logged_in"):
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")

        if (
            ADMIN_PASSWORD
            and secrets.compare_digest(username, ADMIN_USERNAME)
            and secrets.compare_digest(password, ADMIN_PASSWORD)
        ):
            session.clear()
            session["admin_logged_in"] = True
            return redirect(url_for("admin_dashboard"))

        flash("Нэвтрэх нэр эсвэл нууц үг буруу байна.", "danger")

    return render_template("admin.html", login_mode=True)


@app.route("/admin/dashboard")
@admin_required
def admin_dashboard():
    return render_template(
        "admin.html",
        login_mode=False,
        exam_open=get_exam_open(),
        results_exist=RESULTS_FILE.exists(),
    )


@app.route("/admin/toggle", methods=["POST"])
@admin_required
def admin_toggle():
    set_exam_open(not get_exam_open())
    flash(
        "Шалгалтыг нээлээ." if get_exam_open()
        else "Шалгалтыг хаалаа.",
        "success"
    )
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/download")
@admin_required
def admin_download():
    if not RESULTS_FILE.exists():
        init_results_file()

    return send_file(
        RESULTS_FILE,
        as_attachment=True,
        download_name="results.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


@app.route("/admin/logout", methods=["POST"])
@admin_required
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))


@app.errorhandler(404)
def not_found(error):
    return "Хуудас олдсонгүй.", 404


# Анхдагчаар шалгалт хаалттай.
set_exam_open(False)

if ADMIN_PASSWORD is None:
    app.logger.warning(
        "ADMIN_PASSWORD тохируулаагүй тул админ нэвтрэх боломжгүй."
    )

if __name__ == "__main__":
    init_results_file()
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
        debug=False,
    )