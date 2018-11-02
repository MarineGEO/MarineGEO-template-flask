import os
from flask import Flask, request, render_template, send_file, after_this_request
from werkzeug import secure_filename
import MarinegeoTemplateBuilder

UPLOAD_FOLDER = '/tmp/'
ALLOWED_EXTENSIONS = set(['csv'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        fieldcsv = request.files['fieldfile']
        vocabcsv = request.files['vocabfile']
        if fieldcsv and allowed_file(fieldcsv.filename) and vocabcsv and allowed_file(vocabcsv.filename):
            fieldsfilename = secure_filename(fieldcsv.filename)
            fieldcsv.save(os.path.join(app.config['UPLOAD_FOLDER'], fieldsfilename))

            vocabfilename = secure_filename(vocabcsv.filename)
            vocabcsv.save(os.path.join(app.config['UPLOAD_FOLDER'], vocabfilename))

            # build the excel template
            outputFile = "Template.xlsx"
            MarinegeoTemplateBuilder.main(os.path.join(app.config['UPLOAD_FOLDER'], outputFile),
                                          fields=os.path.join(app.config['UPLOAD_FOLDER'], fieldsfilename),
                                          vocab=os.path.join(app.config['UPLOAD_FOLDER'], vocabfilename))

            @after_this_request
            def remove_file(response):
                try:
                    os.remove(os.path.join(app.config['UPLOAD_FOLDER'], fieldsfilename))
                    os.remove(os.path.join(app.config['UPLOAD_FOLDER'], vocabfilename))
                    os.remove(os.path.join(app.config['UPLOAD_FOLDER'], outputFile))
                except Exception as error:
                    app.logger.error("Error removing or closing downloaded file handle", error)
                return response

            return send_file(os.path.join(app.config['UPLOAD_FOLDER'], outputFile), as_attachment=True)

    return render_template("uploadTemplate.html")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5001, debug=True)