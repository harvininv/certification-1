import os
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for, flash
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from PIL import Image
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import ssl
import smtplib

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

UPLOAD_FOLDER = 'uploads'
TEMPLATES_FOLDER = 'templates'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TEMPLATES_FOLDER, exist_ok=True)

# Email configuration
sender_email = 'info@harvinntechnologies.co.in'
sender_password = 'qysc gvcu ptmx nrwm'

# Required columns - simplified to just Name and Email
required_columns = ['Name', 'Email']

# Color for name text (you can customize this)
name_color = (29/255, 54/255, 84/255)  # RGB color


def create_certificate(name, template_path):
    """
    Create a certificate with just the name printed on it.
    
    Args:
        name: Student name to print on certificate
        template_path: Path to the certificate template image
    
    Returns:
        Path to the created PDF file or error message
    """
    try:
        with Image.open(template_path) as img:
            img_width, img_height = img.size

        # Create the PDF with the same dimensions as the image
        file_name = os.path.join(UPLOAD_FOLDER, f"certificate_{name}.pdf")
        c = canvas.Canvas(file_name, pagesize=(img_width, img_height))

        # Draw the template background
        c.drawImage(template_path, 0, 0, img_width, img_height)

        # Convert name to uppercase
        name = name.upper()

        # Set font and color for the name
        font_size = 64
        c.setFont("Helvetica-Bold", font_size)
        c.setFillColorRGB(*name_color)

        # Measure the name width
        name_width = c.stringWidth(name, "Helvetica-Bold", font_size)

        # Center the name on the x-axis
        name_x = (img_width - name_width) / 2
        name_y = 660  # Y-position for name (adjust this based on your template)
        
        c.drawString(name_x, name_y, name)

        c.save()
        return file_name
    except Exception as e:
        return f"Error creating certificate for {name}: {e}"


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    """Handle file upload and certificate generation."""
    if request.method == 'POST':
        # Check if template exists before processing
        template_path = os.path.join(TEMPLATES_FOLDER, "Training.jpg")
        if not os.path.exists(template_path):
            flash('Certificate template (Training.jpg) not found in templates folder!')
            return redirect(request.url)
        
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if file and file.filename.endswith('.xlsx'):
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            
            try:
                df = pd.read_excel(file_path)

                # Check if the uploaded file contains all required columns
                if set(required_columns).issubset(df.columns):
                    creation_errors = []

                    for index, row in df.iterrows():
                        try:
                            # Create certificate with just the name
                            error = create_certificate(
                                name=row['Name'],
                                template_path=os.path.join(TEMPLATES_FOLDER, "Training.jpg")
                            )
                            
                            if "Error" in error:
                                creation_errors.append(error)

                        except Exception as e:
                            creation_errors.append(f"Error processing row {index}: {e}")

                    if creation_errors:
                        for error in creation_errors:
                            flash(error)
                    else:
                        flash("All certificates created successfully!")
                else:
                    missing_columns = set(required_columns) - set(df.columns)
                    flash(f"Missing columns in the uploaded file: {', '.join(missing_columns)}")
                    
            except Exception as e:
                flash(f"Error processing file: {e}")

            return redirect(url_for('upload_file'))
    
    return render_template('index.html')


@app.route('/send-emails', methods=['POST'])
def send_emails():
    """Send certificates via email to all students in the Excel file."""
    try:
        # Look for the uploaded Excel file
        excel_files = [f for f in os.listdir(UPLOAD_FOLDER) if f.endswith('.xlsx')]
        
        if not excel_files:
            flash("No Excel file found. Please upload a file first.")
            return redirect(url_for('upload_file'))
        
        df = pd.read_excel(os.path.join(UPLOAD_FOLDER, excel_files[0]))

        context = ssl.create_default_context()
        email_errors = []
        success_count = 0

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(sender_email, sender_password)

            for index, row in df.iterrows():
                try:
                    name = row['Name']
                    email = row['Email']

                    # Create a multipart message and set headers
                    message = MIMEMultipart()
                    message['From'] = sender_email
                    message['To'] = email
                    message['Subject'] = 'Certificate of Participation - Harvinn Technologies Workshop'

                    # Add body to email
                    body = f"""Dear {name},

Congratulations! Please find attached your certificate of participation for the "Latest Technology Insights and Career Pathway" workshop.

We appreciate your participation and wish you all the best in your future endeavors.

Best regards,
Harvinn Technologies Team
"""
                    message.attach(MIMEText(body, 'plain'))

                    # Attach certificate for the student
                    pdf_file = f"certificate_{name}.pdf"
                    pdf_path = os.path.join(UPLOAD_FOLDER, pdf_file)
                    
                    if not os.path.exists(pdf_path):
                        email_errors.append(f"Certificate not found for {name}")
                        continue
                    
                    with open(pdf_path, 'rb') as attachment:
                        part = MIMEApplication(attachment.read(), Name=pdf_file)
                    part['Content-Disposition'] = f'attachment; filename="{pdf_file}"'
                    message.attach(part)

                    # Send the email
                    smtp.send_message(message)
                    success_count += 1
                    
                except Exception as e:
                    email_errors.append(f"Error sending email to {email}: {e}")

        if email_errors:
            for error in email_errors:
                flash(error)
        
        flash(f"Successfully sent {success_count} certificates!")
        
    except Exception as e:
        flash(f"Error sending emails: {e}")

    return redirect(url_for('upload_file'))


if __name__ == '__main__':
    app.run(debug=True)