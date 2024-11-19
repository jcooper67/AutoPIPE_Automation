from docx import Document
import math
class Editor:
    def __init__(self,output_path):
        self.doc = Document(output_path)
        self.output_path = output_path

    def update_support_table(self, support_dict):
        table = self.doc.tables[3]
        for i, key in enumerate(support_dict):
            support_list = support_dict[key]
            #print (support_list)
            first_row = table.rows[i*3+1]
            second_row = table.rows[i*3+2]
            third_row = table.rows[i*3+3]
            first_row.cells[0].text = key
            first_row.cells[1].text = str(support_list[0])
            second_row.cells[3].text = str(support_list[1])
            second_row.cells[4].text = str(support_list[2])

    def update_nozzle_tables(self, nozzle_dict):
        for i, key in enumerate(nozzle_dict):
            nozzle_list = nozzle_dict[key]
            table = self.doc.tables[i]
            
            second_row = table.rows[1]
            second_row.cells[9].text = key
            
            twelfth_row = table.rows[11]
            #Deadweight Forces
            twelfth_row.cells[2].text = str(nozzle_list[0][0])
            twelfth_row.cells[3].text = str(nozzle_list[0][1])
            twelfth_row.cells[4].text = str(nozzle_list[0][2])

            #Calc Deadweight FSR
            deadweight_fsr = math.sqrt(((nozzle_list[0][1])**2)+((nozzle_list[0][2])**2))
            deadweight_fsr = round(deadweight_fsr,0)
            twelfth_row.cells[5].text = str(deadweight_fsr)


            #Deadweight Moments
            twelfth_row.cells[6].text = str(nozzle_list[0][3])
            twelfth_row.cells[7].text = str(nozzle_list[0][4])
            twelfth_row.cells[8].text = str(nozzle_list[0][5])

            deadweight_msr = math.sqrt(((nozzle_list[0][4])**2)+((nozzle_list[0][5])**2))
            deadweight_msr = round(deadweight_msr,0)

            twelfth_row.cells[9].text = str(deadweight_msr)

            #Th-1 Forces
            thirteenth_row = table.rows[12]
            thirteenth_row.cells[2].text = str(nozzle_list[1][0])
            thirteenth_row.cells[3].text = str(nozzle_list[1][1])
            thirteenth_row.cells[4].text = str(nozzle_list[1][2])

            th1_fsr = math.sqrt((nozzle_list[1][1])**2+(nozzle_list[1][2])**2)
            th1_fsr = round(th1_fsr,0)

            thirteenth_row.cells[5].text = str(th1_fsr)



            #Th-1 Moments
            thirteenth_row.cells[6].text = str(nozzle_list[1][3])
            thirteenth_row.cells[7].text = str(nozzle_list[1][4])
            thirteenth_row.cells[8].text = str(nozzle_list[1][5])

            th1_msr = math.sqrt((nozzle_list[1][4])**2+(nozzle_list[1][5])**2)
            th1_msr = round(th1_msr,0)

            thirteenth_row.cells[9].text = str(th1_msr)




            #Th-2 Forces
            fourteenth_row = table.rows[13]
            fourteenth_row.cells[2].text = str(nozzle_list[2][0])
            fourteenth_row.cells[3].text = str(nozzle_list[2][1])
            fourteenth_row.cells[4].text = str(nozzle_list[2][2])

            th2_fsr = math.sqrt((nozzle_list[2][1])**2+(nozzle_list[2][2])**2)
            th2_fsr = round(th2_fsr,0)

            fourteenth_row.cells[5].text = str(th2_fsr)

            
            #Th-2 Moments
            fourteenth_row.cells[6].text = str(nozzle_list[2][3])
            fourteenth_row.cells[7].text = str(nozzle_list[2][4])
            fourteenth_row.cells[8].text = str(nozzle_list[2][5])

            th2_msr = math.sqrt((nozzle_list[2][4])**2+(nozzle_list[2][5])**2)
            th2_msr = round(th2_msr,0)

            fourteenth_row.cells[9].text = str(th2_msr)

            #Calculate Hot(Weight + Envelope of Expansion Cases)

            fifthteenth_row = table.rows[14]

            deadweight_list = nozzle_list[0]

            thermal1_list = nozzle_list[1]
            thermal2_list = nozzle_list[2]

            envelope_thermal = [thermal1_list[i] if abs(thermal1_list[i]) >= abs(thermal2_list[i]) else thermal2_list[i] for i in range(len(thermal1_list))]
            
            hot_list = [deadweight_list[i]+envelope_thermal[i] for i in range(len(deadweight_list))]

            hot_fsr = deadweight_fsr + max(abs(th1_fsr),abs(th2_fsr))

            hot_msr = deadweight_msr + max(abs(th1_msr),abs(th2_msr))


            fifthteenth_row.cells[2].text = str(hot_list[0])
            fifthteenth_row.cells[3].text = str(hot_list[1])
            fifthteenth_row.cells[4].text = str(hot_list[2])
            fifthteenth_row.cells[5].text = str(hot_fsr)
            
            #Th-2 Moments
            fifthteenth_row.cells[6].text = str(hot_list[3])
            fifthteenth_row.cells[7].text = str(hot_list[4])
            fifthteenth_row.cells[8].text = str(hot_list[5])
            fifthteenth_row.cells[9].text = str(hot_msr)


        #Calc Maximum of DW,Hot

            max_list = [max(abs(hot_list), abs(deadweight_list)) for hot_list, deadweight_list in zip(hot_list, deadweight_list)]

            seventeenth_row = table.rows[16]

            seventeenth_row.cells[2].text = str(max_list[0])

            max_fsr = max(abs(deadweight_fsr),abs(hot_fsr))

            seventeenth_row.cells[5].text = str(max_fsr)

            seventeenth_row.cells[6].text = str(max_list[3])

            max_msr = max(abs(deadweight_msr),abs(hot_msr))

            seventeenth_row.cells[9].text = str(max_msr)

    def save_document(self):
        self.doc.save(self.output_path)