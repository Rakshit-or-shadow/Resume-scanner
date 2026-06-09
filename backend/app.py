from flask import Flask, request, jsonify, send_from_directory
import pdfplumber
import docx
import os
import re

BASE_DIR = os.path.dirname(__file__)
FRONTEND_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "frontend"))

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

SKILL_CATEGORIES = {
    "Programming Languages": [
        (r"python",          "Python"),
        (r"c\+\+",           "C++"),
        (r"c#",              "C#"),
        (r"java(?!script)",  "Java"),
        (r"javascript",      "JavaScript"),
        (r"typescript",      "TypeScript"),
        (r"\bgo\b",          "Go"),
        (r"rust",            "Rust"),
        (r"kotlin",          "Kotlin"),
        (r"swift",           "Swift"),
        (r"ruby",            "Ruby"),
        (r"\bphp\b",         "PHP"),
        (r"scala",           "Scala"),
        (r"matlab",          "MATLAB"),
        (r"\bbash\b",        "Bash"),
        (r"\bsql\b",         "SQL"),
        (r"\bhtml\b",        "HTML"),
        (r"\bcss\b",         "CSS"),
        (r"assembly",        "Assembly"),
        (r"\blua\b",        "Lua"),
        (r"\bc,",            "C"),
        (r"proficient in.*?\bc\b", "C"),
    ],
    "AI / ML / Data Science": [
        (r"machine learning",           "Machine Learning"),
        (r"deep learning",              "Deep Learning"),
        (r"neural network",             "Neural Networks"),
        (r"natural language processing","NLP"),
        (r"\bnlp\b",                    "NLP"),
        (r"computer vision",            "Computer Vision"),
        (r"reinforcement learning",     "Reinforcement Learning"),
        (r"tensorflow",                 "TensorFlow"),
        (r"pytorch",                    "PyTorch"),
        (r"keras",                      "Keras"),
        (r"scikit.learn",               "scikit-learn"),
        (r"pandas",                     "Pandas"),
        (r"numpy",                      "NumPy"),
        (r"matplotlib",                 "Matplotlib"),
        (r"transformers",               "Transformers"),
        (r"sentiment analysis",         "Sentiment Analysis"),
        (r"\bgpt\b",                    "GPT"),
        (r"opencv",                     "OpenCV"),
        (r"image processing",           "Image Processing"),
        (r"\bocr\b",                    "OCR"),
        (r"hough transform",            "Hough Transform"),
        (r"contour detection",          "Contour Detection"),
        (r"image preprocessing",        "Image Preprocessing"),
        (r"text deskewing",             "Text Deskewing"),
        (r"orientation detection",      "Orientation Detection"),
    ],
    "Web & Backend": [
        (r"\bflask\b",   "Flask"),
        (r"django",      "Django"),
        (r"fastapi",     "FastAPI"),
        (r"express",     "Express"),
        (r"node\.js",    "Node.js"),
        (r"\breact\b",   "React"),
        (r"\bvue\b",     "Vue"),
        (r"angular",     "Angular"),
        (r"\brest\b",    "REST"),
        (r"\bapi\b",     "API"),
        (r"graphql",     "GraphQL"),
        (r"websocket",   "WebSocket"),
    ],
    "Databases": [
        (r"mysql",        "MySQL"),
        (r"postgresql",   "PostgreSQL"),
        (r"sqlite",       "SQLite"),
        (r"mongodb",      "MongoDB"),
        (r"\bredis\b",    "Redis"),
        (r"firebase",     "Firebase"),
    ],
    "Cloud & DevOps": [
        (r"\baws\b",      "AWS"),
        (r"\bazure\b",    "Azure"),
        (r"\bgcp\b",      "GCP"),
        (r"docker",       "Docker"),
        (r"kubernetes",   "Kubernetes"),
        (r"\bgit\b",      "Git"),
        (r"github",       "GitHub"),
        (r"\blinux\b",    "Linux"),
    ],
    "Game Development": [
        (r"\bunity\b",         "Unity"),
        (r"unreal engine",     "Unreal Engine"),
        (r"game development",  "Game Development"),
        (r"c# programming",    "C# Programming"),
        (r"game mechanics",    "Game Mechanics"),
        (r"ray tracing",       "Ray Tracing"),
        (r"3d rendering",      "3D Rendering"),
        (r"rendering engine",  "Rendering Engine"),
        (r"vector math",       "Vector Math"),
        (r"camera projection", "Camera Projection"),
        (r"shading",           "Shading"),
        (r"anti.aliasing",     "Anti-Aliasing"),
        (r"aim labs",          "Aim Labs"),
        (r"\bgodot\b",         "Godot"),
    ],
    "Systems & Low-Level": [
        (r"memory allocation",  "Memory Allocation"),
        (r"memory management",  "Memory Management"),
        (r"file i/o",           "File I/O"),
        (r"multithreading",     "Multithreading"),
        (r"\bthreading\b",      "Threading"),
        (r"parallel processing","Parallel Processing"),
        (r"makefile",           "Makefile"),
        (r"\bqsort\b",          "qsort"),
        (r"data structures",    "Data Structures"),
        (r"\balgorithm",        "Algorithms"),
        (r"\bdsa\b",            "DSA"),
        (r"\bbvh\b",          "BVH Acceleration"),
        (r"ray.sphere",         "Ray-Sphere Intersection"),
        (r"exception handling", "Exception Handling"),
    ],
    "Blockchain & Finance": [
        (r"coinbase",           "Coinbase API"),
        (r"cryptocurrency",     "Cryptocurrency"),
        (r"trading bot",        "Trading Bot"),
        (r"blockchain",         "Blockchain"),
        (r"\bsma\b",            "SMA Signals"),
        (r"real.time trading",  "Real-Time Trading"),
        (r"json trade",         "JSON Trade Logging"),
    ],
    "Hardware & Embedded": [
        (r"\brfid\b",      "RFID"),
        (r"bluetooth",     "Bluetooth"),
        (r"arduino",       "Arduino"),
        (r"raspberry pi",  "Raspberry Pi"),
        (r"\biot\b",      "IoT"),
        (r"sensors",       "Sensors"),
    ],
    "Soft Skills": [
        (r"teamwork",        "Teamwork"),
        (r"collaboration",   "Collaboration"),
        (r"leadership",      "Leadership"),
        (r"communication",   "Communication"),
        (r"problem.solving", "Problem-Solving"),
        (r"project management","Project Management"),
        (r"structured workflow","Structured Workflow"),
    ]
}


def extract_text_from_pdf(filepath):
    text = ""
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


def extract_text_from_docx(filepath):
    doc = docx.Document(filepath)
    return "\n".join([para.text for para in doc.paragraphs])


def extract_text(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(filepath)
    elif ext in (".docx", ".doc"):
        return extract_text_from_docx(filepath)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def find_skills(text):
    found = {}
    for category, entries in SKILL_CATEGORIES.items():
        seen_labels = set()
        matched = []
        for (pattern, label) in entries:
            if re.search(pattern, text, re.IGNORECASE) and label not in seen_labels:
                matched.append(label)
                seen_labels.add(label)
        if matched:
            found[category] = matched
    return found


def extract_candidate_name(text):
    for line in text.strip().splitlines():
        line = line.strip()
        if line and len(line.split()) <= 5 and not any(c in line for c in ["@", "http", "+"]):
            return line
    return "Candidate"


@app.route("/")
def index():
    return send_from_directory(FRONTEND_DIR, "index.html")


@app.route("/<path:filename>")
def frontend_static(filename):
    return send_from_directory(FRONTEND_DIR, filename)


@app.route("/upload", methods=["POST"])
def upload():
    if "resume" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["resume"]
    if file.filename == "":
        return jsonify({"error": "Empty filename"}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in (".pdf", ".docx", ".doc"):
        return jsonify({"error": "Only PDF and DOCX files are supported"}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    try:
        text = extract_text(filepath)
        skills = find_skills(text)
        name = extract_candidate_name(text)
        total = sum(len(v) for v in skills.values())

        return jsonify({
            "name": name,
            "skills": skills,
            "total": total,
            "raw_preview": text[:500]
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(filepath):
            os.remove(filepath)


if __name__ == "__main__":
    print("🚀  Resume Skill Extractor running at http://localhost:5000")
    app.run(debug=True, port=5000)
