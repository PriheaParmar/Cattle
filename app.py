from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
import tempfile
from datetime import datetime
import re

app = Flask(__name__)

def format_date(date_string):
    if date_string:
        try:
            date_obj = datetime.strptime(date_string, '%Y-%m-%d')
            return date_obj.strftime('%d/%m/%Y')
        except:
            return date_string
    return ''

def format_date_range(date_range_string):
    if date_range_string and ' - ' in date_range_string:
        try:
            start_date, end_date = date_range_string.split(' - ')
            formatted_start = format_date(start_date.strip())
            formatted_end = format_date(end_date.strip())
            return f"from {formatted_start} to {formatted_end}"
        except:
            return date_range_string
    return date_range_string

def to_sentence_case(text):
    if text:
        if text.lower() == 'hf cow':
            return 'HF Cow'
        return text.strip().title()
    return text

def sanitize_filename(text):
    """Remove or replace characters that are invalid in filenames"""
    if not text:
        return 'UNKNOWN'
    # Replace forward slashes, backslashes, and other invalid filename characters
    # Keep only alphanumeric, hyphens, and underscores
    sanitized = re.sub(r'[^\w\s-]', '', text)
    # Replace spaces with hyphens
    sanitized = sanitized.replace(' ', '-')
    # Remove multiple consecutive hyphens
    sanitized = re.sub(r'-+', '-', sanitized)
    # Convert to uppercase and strip
    return sanitized.upper().strip('-') or 'UNKNOWN'

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/generate', methods=['POST'])
def generate(): 
    
    claimdate = request.form.get('claimdate', '').strip().upper()

    try:
        single_date = format_date(request.form['date'])
       
        context = {
            'claimdate': claimdate,
            'dayssick': request.form.get('dayssick'),
            'date': single_date,
            'intimdate': single_date,
            'invdate': single_date,
            'lossdate': format_date(request.form['lossdate']),
            'tagnumber': request.form['tagnumber'],
            'cattletype': to_sentence_case(request.form['cattletype']),
            'ownername': to_sentence_case(request.form['ownername']),
            'ownercontact': request.form['ownercontact'],
            'location': to_sentence_case(request.form['location']),
            'taluka': to_sentence_case(request.form['taluka']),
            'losstime': request.form['losstime'],
            'policynumber': request.form['policynumber'],
            'policyperiod': format_date_range(request.form['policyperiod']),
            'insuredname': to_sentence_case(request.form['insuredname']),
            'loan': to_sentence_case(request.form['loan']),
            'disease': to_sentence_case(request.form['disease']),
            'gvc': request.form['gvc'],
            'taglocation': to_sentence_case(request.form['taglocation']),
            'cattlecolor': to_sentence_case(request.form['cattlecolor']),
            'lactation': request.form['lactation'],
            'pregnant': to_sentence_case(request.form['pregnant']),
            'cattletail': to_sentence_case(request.form['cattletail']),
            'milkey': to_sentence_case(request.form['milkey']),
            'specialmarks': to_sentence_case(request.form['specialmarks']),
            'treatment': to_sentence_case(request.form['treatment']),
            'deathtype': to_sentence_case(request.form['deathtype']),
            'visit': to_sentence_case(request.form['visit']),
            
            # BANK DETAILS
            'accountnumber': request.form['accountnumber'],
            'ifsccode': request.form['ifsccode'].upper(),
            'bankname': to_sentence_case(request.form['bankname'])
        }

        # Load and render template
        template_path = os.path.join(os.getcwd(), 'template.docx')
        template = DocxTemplate(template_path)
        template.render(context)
        
        # Sanitize filename components to remove invalid characters
        location = sanitize_filename(context['location'])
        taluka = sanitize_filename(context['taluka'])
        ownername = sanitize_filename(context['ownername'])
        tag_number = context['tagnumber'][-6:].replace('/', '-')
        
        # Create safe filename
        output_filename = f"{location}_{taluka}_{ownername}_{tag_number}.docx"
        
        # Use tempfile to create a safe temporary file
        temp_dir = tempfile.gettempdir()
        output_path = os.path.join(temp_dir, output_filename)
        
        # Ensure the temp directory exists
        os.makedirs(temp_dir, exist_ok=True)
        
        # Save the document
        template.save(output_path)
        
        return send_file(
            output_path, 
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    
    except Exception as e:
        app.logger.error(f"Error generating document: {str(e)}")
        return f"Error generating document: {str(e)}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))