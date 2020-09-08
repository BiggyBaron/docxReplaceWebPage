from docx import Document
from flask import Flask, render_template, request, send_file, Markup
from werkzeug.utils import secure_filename
import excel2json
import json
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


def generate_link(val):
    link_template = "<a href='/uploader/LINK'>Генерировать</a>"
    return link_template.replace("LINK", val)


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

    # Backend reads values from data xlsx file from 1st page and converts it to json.
    # WARNING: name of the sheet has to be "1"
    # Then creates html table and insert it into html page.
    # WARNING: backend create "Generate" button if cell in coulumn "Генерить" is equal to +.
    # No other cell has to be equal to + or - and id coulumn is required
    excel2json.convert_from_file('data.xlsx')
    # We open json file and convert it to variable
    with open('1.json') as f:
        data = json.load(f)

    # This is structure of table in HTML by parts.
    template_head = '<table style="width:100%">'
    template_row = '<tr>N</tr>'
    template_headers = '<th>N</th>'
    template_normal = '<td>N</td>'
    template_footer = '</table>'

    # empty output for HTML.
    output = template_head

    # Create empty string to store keys
    k = ""

    # Create headers
    for key in data[0].keys():
        k = k + template_headers.replace("N", str(key))

    # Create row to the table
    output = output + template_row.replace("N", k)

    # Create empty string to store rows
    rows = ""
    # Create other table
    for d in range(len(data)):
        # Create empty string to store each row
        row_raw = ""
        for d1 in data[d].values():
            # If value is equal to + it creates button
            if d1 == "+":
                # It creates hyperlink to the uploader page
                link = generate_link(str(int(data[d]["id"])))
                val = template_normal.replace("N", link)
            else:
                # Esle: puts value to the table
                val = template_normal.replace("N", str(d1))
            # Collects value in table
            row_raw = row_raw + val
        # Create full row of table
        rows = rows + template_row.replace("N", row_raw)
    
    # Collect all rows
    output = Markup(output + rows + template_footer)

    return render_template('index.html', port = port_markedup, name = name, output=output)


# This function takes data from html form in the index.html file
# then generates new docx file and sends to a client
@app.route('/uploader/<id>', methods = ['GET', 'POST'])
def upload_file(id):
    # We open json to find data about cat
    with open('1.json') as f:
        data = json.load(f)
    
    # We find data about cat in json
    for d in data:
        if str(int(d["id"])) == id:
            name = d["Имя"]
            status = d["Данные"]
            print(status)

    # We put data to the template docx
    replace_string("old.docx", "КОТ", str(name), "changed.docx")
    # We change data several times
    replace_string("changed.docx", "СТАТУС", str(status), "changed.docx")

    # This sends file to the user
    return send_file("changed.docx", as_attachment=True)


if __name__ == '__main__':
    # Run application on the 5000th port
    app.run(debug = True, port="5000")