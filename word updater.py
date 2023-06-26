from docx import Document

def update_word_template(template_filename, output_filename):
    # Open the template document
    document = Document(template_filename)

    # Get user inputs
    name = input("Enter name: ")
    date = input("Enter date: ")
    serial_number = input("Enter serial number: ")
    company_name = input ("Enter company name: ")
    id = input ('Enter customer ID: ')
    Tin = input ("Enter time in : ")
    Tout = input ('Enter time out : ')


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
        if "<ID>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<id>", id)   
        if "<tin>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<tin>", Tin)
        if "<tout>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<tout>", Tout)

    # Save the updated document as a new Word file
    document.save(output_filename)
    print(f"Updated Word file '{output_filename}' generated successfully.")

# Example usage
template_filename = "template.docx"
output_filename = "output.docx"

update_word_template(template_filename, output_filename)
