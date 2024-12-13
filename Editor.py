from docx import Document
import math
from tkinter import messagebox
import tkinter as tk
import sys
class Editor:
    def __init__(self,output_path):
        self.doc = Document(output_path)
        self.output_path = output_path
        #Assumes the support table is the 16th table
        self.support_table_location = 15

    def update_support_table(self, support_dict):
        try:
            table = self.doc.tables[self.support_table_location]
            for i, key in enumerate(support_dict):
                support_list = support_dict[key]
                first_row = table.rows[i*3+1]
                first_row.cells[0].text = key
                first_row.cells[1].text = str(support_list[0])

                for j in range(1,4):
                    if support_list[j][0] == 0:
                        continue
                    else:
                        row = table.rows[i*3+j]
                        row.cells[3].text = str(support_list[j][0])
                        row.cells[4].text = str(support_list[j][3])

        except Exception as e:
            print(f"Error: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None
            


    def update_nozzle_tables(self, nozzle_dict):
        try:
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
                deadweight_fsr = int(round(deadweight_fsr,0))
                twelfth_row.cells[5].text = str(deadweight_fsr)


                #Deadweight Moments
                twelfth_row.cells[6].text = str(nozzle_list[0][3])
                twelfth_row.cells[7].text = str(nozzle_list[0][4])
                twelfth_row.cells[8].text = str(nozzle_list[0][5])

                deadweight_msr = math.sqrt(((nozzle_list[0][4])**2)+((nozzle_list[0][5])**2))
                deadweight_msr = int(round(deadweight_msr,0))

                twelfth_row.cells[9].text = str(deadweight_msr)

                #Th-1 Forces
                thirteenth_row = table.rows[12]
                thirteenth_row.cells[2].text = str(nozzle_list[1][0])
                thirteenth_row.cells[3].text = str(nozzle_list[1][1])
                thirteenth_row.cells[4].text = str(nozzle_list[1][2])

                th1_fsr = math.sqrt((nozzle_list[1][1])**2+(nozzle_list[1][2])**2)
                th1_fsr = int(round(th1_fsr,0))

                thirteenth_row.cells[5].text = str(th1_fsr)



                #Th-1 Moments
                thirteenth_row.cells[6].text = str(nozzle_list[1][3])
                thirteenth_row.cells[7].text = str(nozzle_list[1][4])
                thirteenth_row.cells[8].text = str(nozzle_list[1][5])

                th1_msr = math.sqrt((nozzle_list[1][4])**2+(nozzle_list[1][5])**2)
                th1_msr = int(round(th1_msr,0))

                thirteenth_row.cells[9].text = str(th1_msr)




                #Th-2 Forces
                fourteenth_row = table.rows[13]
                fourteenth_row.cells[2].text = str(nozzle_list[2][0])
                fourteenth_row.cells[3].text = str(nozzle_list[2][1])
                fourteenth_row.cells[4].text = str(nozzle_list[2][2])

                th2_fsr = math.sqrt((nozzle_list[2][1])**2+(nozzle_list[2][2])**2)
                th2_fsr = int(round(th2_fsr,0))

                fourteenth_row.cells[5].text = str(th2_fsr)

                
                #Th-2 Moments
                fourteenth_row.cells[6].text = str(nozzle_list[2][3])
                fourteenth_row.cells[7].text = str(nozzle_list[2][4])
                fourteenth_row.cells[8].text = str(nozzle_list[2][5])

                th2_msr = math.sqrt((nozzle_list[2][4])**2+(nozzle_list[2][5])**2)
                th2_msr = int(round(th2_msr,0))

                fourteenth_row.cells[9].text = str(th2_msr)


                #Th-3 Forces
                fifteenth_row = table.rows[14]
                fifteenth_row.cells[2].text = str(nozzle_list[2][0])
                fifteenth_row.cells[3].text = str(nozzle_list[2][1])
                fifteenth_row.cells[4].text = str(nozzle_list[2][2])

                th2_fsr = math.sqrt((nozzle_list[2][1])**2+(nozzle_list[2][2])**2)
                th2_fsr = int(round(th2_fsr,0))

                fifteenth_row.cells[5].text = str(th2_fsr)

                
                #Th-3 Moments
                fifteenth_row.cells[6].text = str(nozzle_list[2][3])
                fifteenth_row.cells[7].text = str(nozzle_list[2][4])
                fifteenth_row.cells[8].text = str(nozzle_list[2][5])

                th2_msr = math.sqrt((nozzle_list[2][4])**2+(nozzle_list[2][5])**2)
                th2_msr = int(round(th2_msr,0))

                fifteenth_row.cells[9].text = str(th2_msr)


                #Th-4 Forces
                sixteenth_row = table.rows[15]
                sixteenth_row.cells[2].text = str(nozzle_list[2][0])
                sixteenth_row.cells[3].text = str(nozzle_list[2][1])
                sixteenth_row.cells[4].text = str(nozzle_list[2][2])

                th2_fsr = math.sqrt((nozzle_list[2][1])**2+(nozzle_list[2][2])**2)
                th2_fsr = int(round(th2_fsr,0))

                sixteenth_row.cells[5].text = str(th2_fsr)

                
                #Th-4 Moments
                sixteenth_row.cells[6].text = str(nozzle_list[2][3])
                sixteenth_row.cells[7].text = str(nozzle_list[2][4])
                sixteenth_row.cells[8].text = str(nozzle_list[2][5])

                th2_msr = math.sqrt((nozzle_list[2][4])**2+(nozzle_list[2][5])**2)
                th2_msr = int(round(th2_msr,0))

                sixteenth_row.cells[9].text = str(th2_msr)



                #Calculate Hot(Weight + Envelope of Expansion Cases)

                seventeenth_row = table.rows[16]

                deadweight_list = nozzle_list[0]

                thermal1_list = nozzle_list[1]
                thermal2_list = nozzle_list[2]

                hot_list = [a + b for a, b in zip(deadweight_list, thermal2_list)]


                # hot_list = []

                # for i in range(len(thermal1_list)):
                #     dead_th1 = deadweight_list[i]+thermal1_list[i]
                #     dead_th2 = deadweight_list[i]+thermal2_list[i]

                #     if abs(dead_th1)>abs(dead_th2):
                #         hot_list.append(dead_th1)
                #     else:
                #         hot_list.append(dead_th2)
                #print(f"hot_list: {hot_list}")

                ## FSR and MSR to be calculated as resultant of hot x and hot z
                #hot_fsr = deadweight_fsr + max(abs(th1_fsr),abs(th2_fsr))
                hot_fsr = math.sqrt((hot_list[1])**2+(hot_list[2])**2)
                hot_fsr = int(round(hot_fsr,0))
                hot_msr = math.sqrt((hot_list[4])**2+(hot_list[5])**2)
                hot_msr = int(round(hot_msr,0))
                #hot_msr = deadweight_msr + max(abs(th1_msr),abs(th2_msr))


                seventeenth_row.cells[2].text = str(hot_list[0])
                seventeenth_row.cells[3].text = str(hot_list[1])
                seventeenth_row.cells[4].text = str(hot_list[2])
                seventeenth_row.cells[5].text = str(hot_fsr)
                
                #Th-2 Moments
                seventeenth_row.cells[6].text = str(hot_list[3])
                seventeenth_row.cells[7].text = str(hot_list[4])
                seventeenth_row.cells[8].text = str(hot_list[5])
                seventeenth_row.cells[9].text = str(hot_msr)


            #Calc Maximum of DW,Hot

                max_list = [max(abs(hot_list), abs(deadweight_list)) for hot_list, deadweight_list in zip(hot_list, deadweight_list)]

                nineteenth_row = table.rows[18]

                nineteenth_row.cells[2].text = str(max_list[0])

                max_fsr = max(abs(deadweight_fsr),abs(hot_fsr))

                nineteenth_row.cells[5].text = str(max_fsr)

                nineteenth_row.cells[6].text = str(max_list[3])

                max_msr = max(abs(deadweight_msr),abs(hot_msr))

                nineteenth_row.cells[9].text = str(max_msr)
        except Exception as e:
            print(f"Error: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None
            

    def save_document(self):
        self.doc.save(self.output_path)