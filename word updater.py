from docx import Document

def update_word_template(template_filename, output_filename):
    # Open the template document
    document = Document(template_filename)

    # Get user inputs
    name = input('Enter name: ')
    date = input('Enter date(DD/MM/YY): ')
    serial_number = input('Enter serial number: ')
    company_name = input ('Enter company name: ')
    id = input ('Enter customer ID: ')
    Tin = input ('Enter time in (24Hr): ')
    Tout = input ('Enter time out (24Hr): ')
    Temi = input ('Enter temi version: ' )
    DOW = input ('Enter Description of work: ' )


    # Update placeholders in the document with user inputs
    for paragraph in document.paragraphs:
        print(paragraph.text)
        
        if "<name>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<name>", name)
        if "<date>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<date>", date)
        if "<serial>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<serial>", serial_number)
        if "<company>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<company>", company_name)
        if "<id>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<id>", id)   
        if "<tin>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<tin>", Tin)
        if "<tout>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<tout>", Tout)
        if "<Temi>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<Temi>", Temi)

        # Update placeholders in tables
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if "<id>" in run.text:
                            run.text = run.text.replace("<id>", id)
                            print("run id")
                        if "<date>" in run.text:
                            run.text = run.text.replace("<date>", date)
                            print("run date")
                        if "<DOW>" in run.text:
                            run.text = run.text.replace("<DOW>", DOW)
                            print("dow")

    # Save the updated document as a new Word file
    document.save(output_filename)
    print(f"Updated Word file '{output_filename}' generated successfully.")

# Example usage
template_filename = "template.docx"
output_filename = "output.docx"

update_word_template(template_filename, output_filename)
