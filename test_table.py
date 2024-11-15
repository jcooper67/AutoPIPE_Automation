from docx import Document

doc = Document("C:\\Users\\COOPERJ\\Documents\\Projects\\AutoPIPE_Automation\\Sample Files\\Table_Template1.docx")



def duplicate_row(row_number):
    table = doc.tables[1]
    row = table.rows[row_number]


    new_row = table.add_row()


    for i,cell in enumerate(row.cells):
        new_row.cells[i].text = cell.text




support_table = doc.tables[3]
row = support_table.rows[4]
cell = row.cells[3]
cell.text = 'TEST'
doc.save("C:\\Users\\COOPERJ\\Documents\\Projects\\AutoPIPE_Automation\\Sample Files\\Table_Template1.docx")