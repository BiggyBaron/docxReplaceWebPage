from docx import Document
from flask import Flask, render_template, request, send_file, Markup
from werkzeug.utils import secure_filename
app = Flask(__name__)

# Write here your port, forex is 5000th port with localhost, url gonna be:
# http://localhost:5000/
port = 5000


# This function takes old docx file, finds oldtext and replace it with newtext, then saves newdocx
def replace_string(filename, oldtext, newtext, newfilename):
    doc = Document(filename)
    for p in doc.paragraphs:
        if oldtext in p.text:
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if oldtext in inline[i].text:
                    text = inline[i].text.replace(oldtext, newtext)
                    inline[i].text = text

    doc.save(newfilename)
    return 1


# This function renders html file in "/templates/" folder
@app.route('/')
def renderall():
    # In the index.html file you can find two places like {{port}} or {{name}}
    # This is the way to send variables from backend (python+flask) to the frontend (html)

    # Let`s put name variable and port is defined
    name = "Stranger"

    # But when you send a code, not a text to the html, you have markup it:
    # Forex here, port variable goes to form`s code, so it has to be markuped
    port_markedup = Markup(port)

    return render_template('index.html', port = port_markedup, name = name)


# This function takes data from html form in the index.html file
# then generates new docx file and sends to a client
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():

    # If method is POST, which is POST if smb sends file from form
    if request.method == 'POST':

        # This reads file sended
        f = request.files['file']
        # This saves file as old.docx
        f.save(secure_filename("old.docx"))
        # This function changes text in the file from data from form
        replace_string("old.docx", str(request.form["oldtext"]), str(request.form["newtext"]), "changed.docx")

        # This sends file to the user
        return send_file("changed.docx", as_attachment=True)

		
if __name__ == '__main__':
    # Run application on the 5000th port
    app.run(debug = True, port="5000")