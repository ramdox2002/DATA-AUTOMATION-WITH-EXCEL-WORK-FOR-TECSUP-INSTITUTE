from docxtpl import DocxTemplate
import pandas as pd
from datetime import datetime

global doc, df, context


doc = DocxTemplate("plantilla.docx")

class ConvertDocxExcel:

    def generate_data(self):
        try:
            from faker import Faker

            # Crear instancia de Faker
            fake = Faker()

            # Generar datos falsos
            data = []
            for _ in range(10):  # Generar 10 filas de datos
                name = fake.name()
                date = fake.date()
                address = fake.address()
                phone_number = fake.phone_number()
                email = fake.email()
                title = fake.job()
                reference = fake.random_number(digits=6)
                contact_information = fake.phone_number()
                notes_or_comments = fake.text(max_nb_chars=50)
                signature = fake.name()
                dni = fake.random_number(digits=8)
                data.append([name, date, address, phone_number, email, title, reference, contact_information, notes_or_comments, signature, dni])

            # Crear un DataFrame de Pandas
            df = pd.DataFrame(data, columns=['Name', 'Date', 'Address', 'Phone Number', 'Email', 'Title', 'Reference', 'Contact Information', 'Notes or Comments', 'Signature', 'DNI'])

            # Escribir el DataFrame en un archivo Excel
            df.to_excel('false_data.xlsx', index=False)

        except (ValueError, NameError, AttributeError) as e:
            print("Bug type:", e)

    def extraction_data(self):
        df = pd.read_excel("false_data.xlsx")

        for index, row in df.iterrows():
            context = {
                "name": row["Name"],
                "date": row["Date"],
                "address": row["Address"],
                "phone_number": row["Phone Number"],
                "email": row["Email"],
                "title": row["Title"],
                "reference": row["Reference"],
                "contact_information": row["Contact Information"],
                "notes_or_comments": row["Notes or Comments"],
                "signature": row["Signature"],
                "dni": row["DNI"],
            }
            self.generate_document(context=context, index=index)
        return context

    def generate_document(self, context, index):
        doc.render(context=context)
        doc.save(f"generated_doc_{index}.docx")

    def run(self):
        self.extraction_data()

if __name__ == "__main__":
    task = ConvertDocxExcel()
    """
    Si no existe un archivo exel o no hay data con la que interactuar
    se ejecuta el siguente codigo antes del run()

    task.generate_data()
    """
    task.run()
