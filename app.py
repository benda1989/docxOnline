# pip install pydocx python-docx flask
from docx2html import docx2html
from flask import Flask, request, jsonify, Response, g, send_file
import sqlite3
import hashlib
import time
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['HTML_FOLDER'] = 'htmls/'
app.config['DOCX_FOLDER'] = 'results/'
app.config['ALLOWED_EXTENSIONS'] = {'docx'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['HTML_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOCX_FOLDER'], exist_ok=True)


def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect("db")
    return db


def init_db():
    with app.app_context():
        db = get_db()
        db.cursor().executescript('''CREATE TABLE IF NOT EXISTS datas(
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            name TEXT NOT NULL,
                            md5 TEXT NOT NULL,
                            datas TEXT NOT NULL,
                            docx  bool default false,
                            createdAt timestamp default (datetime('now','localtime'))
                            );
                            CREATE INDEX md5_index ON datas(md5);''')
        db.commit()


def save_db(sql, args=()):
    with app.app_context():
        db = get_db()
        db.cursor().execute(sql, args)
        db.commit()


def query_db(query, args=(), one=False):
    cur = get_db().execute(query, args)
    rv = cur.fetchall()
    cur.close()
    return (rv[0] if rv else None) if one else rv


def getOne(md5):
    return query_db('select * from datas where md5 = ?', (md5,), one=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()


@app.route('/')
def index():
    return Response(open("index.html").read(), mimetype='text/html')


@app.route('/html/<id>', methods=['GET'])
def html(id):
    return Response(open(app.config['HTML_FOLDER'] + id).read(), mimetype='text/html')


@app.route('/docx/<id>', methods=["GET", 'POST'])
def create_docx(id):

    one = getOne(id)
    if one:
        fp = app.config['DOCX_FOLDER'] + id + ".docx"
        if request.method == "GET" and not one[4]:
            return jsonify({'msg': 'No file data'})
        if request.method == "POST":
            data = request.get_json()
            print(data)
            if not data or "inputs" not in data:
                return jsonify({'msg': 'should has inputs'})
            docx2html(app.config['UPLOAD_FOLDER'] + one[1]).save(fp, data.get("inputs"), data.get("table", []))
            save_db("update  datas set docx=true where md5 =?", (id,))
        return send_file(fp, as_attachment=True)
    else:
        return jsonify({'msg': 'No such file'})


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'msg': 'No file part'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'msg': 'No selected file'})
    if file and allowed_file(file.filename):
        fn = str(time.time())+file.filename
        fp = app.config['UPLOAD_FOLDER']+fn
        md5, datas = "", []
        with open(fp, "wb") as fs:
            file.save(fs)
        with open(fp, "rb") as fs:
            md5 = hashlib.md5(fs.read()).hexdigest()
        one = getOne(md5)
        if not one:
            docs = docx2html(fp)
            with open(app.config['HTML_FOLDER']+md5, "w", encoding="utf-8") as f:
                f.write(docs.export())
            datas = docs.inputDatas
            save_db("insert into datas(name,md5,datas)  values (?, ?, ?)",  (fn, md5, str(datas)))
        else:
            os.remove(fp)
            datas = one[3]
        return jsonify({'id': md5,
                        "datas": datas})
    else:
        return jsonify({'msg': 'File type not allowed'})


if __name__ == "__main__":
    # init_db()
    app.run()
