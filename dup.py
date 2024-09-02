def recursive(main, i, main_list):
    if i == len(main):  # Adjusted to use the length of the input string
        return ''.join(main_list)  
    else:
        main_list.append(chr(int(main[i][2:], 16))) 
        return recursive(main, i + 1, main_list)  

def generate_certificate(student_name):
    import os
    import qrcode
    from PyPDF2 import PdfWriter, PdfReader
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    import reportlab.rl_config
    import io
    import pandas as pd

    input_folder = os.path.abspath("Input")
    certificates_folder = os.path.abspath("Certificates")

    participants_file = os.path.join(input_folder, 'participants.xlsx')
    participants = pd.read_excel(participants_file)

    print(participants)  # Print the excel file as a dataframe

    for i in range(len(participants)):
        print(i)
        student = participants.loc[i, 'Student'].upper()
        date = participants.loc[i, 'Date']
        event = participants.loc[i, 'Event']
        issued_date = participants.loc[i, 'Issued_date']
        organization = participants.loc[i, 'Organization']

        packet = io.BytesIO()
        width, height = letter
        c = canvas.Canvas(packet, pagesize=(3 * width, 2 * height))  # Enlarge the canvas to fit bigger names

        reportlab.rl_config.warnOnMissingFontGlyphs = 0

        # For the student name
        c.setFillColorRGB(139 / 255, 119 / 255, 40 / 255)
        c.setFont('Times-Italic', 150)
        c.drawCentredString(800, 820, student)

        # For the event name
        c.setFillColorRGB(0, 0, 0)
        c.setFont('Times-Roman', 50)
        c.drawCentredString(830, 700, event)

        # For the event date
        c.setFillColorRGB(0, 0, 0)
        c.setFont('Times-Roman', 45)
        c.drawCentredString(875, 560, date)

        # To generate the QR code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        # Change this to change QR content
        qr.add_data(f"{student}, Certificate No: RSET/IEDC/CC/2024/{i + 1}, {event}, {date}, Issued Date: {issued_date}, Issued by: {organization} - Rajagiri School of Engineering & Technology (Autonomous), Kochi")

        qr.make(fit=True)
        qr_img = qr.make_image(fill='black', back_color='white')

        temp_qr_file = os.path.join(certificates_folder, "temp_qr_code.png")
        qr_img.save(temp_qr_file)

        c.drawImage(temp_qr_file, 250, 100, width=300, height=300)  # Adjust size and position as needed
        c.save()

        template_file = os.path.join(input_folder, "certificate_template.pdf")
        existing_pdf = PdfReader(open(template_file, "rb"))
        page = existing_pdf.pages[0]

        packet.seek(0)
        new_pdf = PdfReader(packet)
        page.merge_page(new_pdf.pages[0])

        file_name = student.replace(" ", "_")
        certificado = os.path.join(certificates_folder, file_name + "_certificate.pdf")
        outputStream = open(certificado, "wb")
        output = PdfWriter()
        output.add_page(page)
        output.write(outputStream)
        outputStream.close()

        if os.path.exists(temp_qr_file):
            os.remove(temp_qr_file)
        
        print(f"Success: Certificate generated for {student}")

if __name__ == '__main__':
    student_name = "YOUR_NAME_HERE"  # Directly pass the name you want
    generate_certificate(student_name)
