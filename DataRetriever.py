import pandas as pd
import tkinter as tk
from tkinter import messagebox
import Popup
import sys
class DataRetriever:

    def __init__(self,file_path):
        self.file_path = file_path
        self.support_dict = {}
        self.nozzle_dict = {}
        self.popup = Popup.Popup()
        
    def retrieve_support_data(self,support_nodes):
        #Open the document and Set Up the Data Frame
        try:
            df = pd.read_excel(self.file_path,sheet_name = "Support_Forces")
            df.columns = df.iloc[0]
            df = df [1:]
            df = df.reset_index(drop = True)

            # Extract the numeric part of the 'GRT{i}' strings
            df['numeric_part'] = df['Load_Combination'].str.extract(r'GRT(\d+)')

            # Count the number of unique 'GRT{i}' entries
            unique_grts = df['numeric_part'].nunique()
            print(unique_grts)


            df['numeric_part'] = df['Load_Combination'].str.extract(r'Thermal (\d+){1}')

            # Count the number of unique 'i' values
            self.number_thermal_cases = df['numeric_part'].nunique()

            print(f"Number of unique Thermal{{i}} entries: {self.number_thermal_cases}")

            

            for support_node in support_nodes:

                self.support_dict[support_node] = []

                
                support_number = self.retreive_support_number(df=df, support_node=support_node)

                if support_number is None:
                    continue

                self.support_dict[support_node].append(support_number)
                self.support_dict[support_node].append([])
                self.support_dict[support_node].append([])
                self.support_dict[support_node].append([])
                #print(self.support_dict[support_node])


                # forces = self.retrieve_forces_by_load_case(df=df, support_node= support_node,load_case = 'GRT{1}')
                # self.support_dict[support_node][1].append(forces[0])
                # self.support_dict[support_node][2].append(forces[1])
                # self.support_dict[support_node][3].append(forces[2])
        

                # forces = self.retrieve_forces_by_load_case(df=df, support_node= support_node,load_case = 'GRT{2}')
                # self.support_dict[support_node][1].append(forces[0])
                # self.support_dict[support_node][2].append(forces[1])
                # self.support_dict[support_node][3].append(forces[2])


                # forces = self.retrieve_forces_by_load_case(df=df, support_node= support_node,load_case = 'GRT{3}')
                # self.support_dict[support_node][1].append(forces[0])
                # self.support_dict[support_node][2].append(forces[1])
                # self.support_dict[support_node][3].append(forces[2])

                # forces = self.retrieve_forces_by_load_case(df=df, support_node= support_node,load_case = 'GRT{4}')
                # self.support_dict[support_node][1].append(forces[0])
                # self.support_dict[support_node][2].append(forces[1])
                # self.support_dict[support_node][3].append(forces[2])

                for i in range(unique_grts):
                    print(i)
                    forces = self.retrieve_forces_by_load_case(df=df, support_node=support_node, load_case=f'GRT{{{i+1}}}')
                    self.support_dict[support_node][1].append(forces[0])
                    self.support_dict[support_node][2].append(forces[1])
                    self.support_dict[support_node][3].append(forces[2])


                sorted_x = sorted(self.support_dict[support_node][1], key = abs,reverse=True)
                sorted_y = sorted(self.support_dict[support_node][2], key = abs, reverse = True)
                sorted_z = sorted(self.support_dict[support_node][3], key = abs, reverse = True)

                self.support_dict[support_node][1] = sorted_x
                self.support_dict[support_node][2] = sorted_y
                self.support_dict[support_node][3] = sorted_z

                
            #print(self.support_dict)
            return self.support_dict
        
        except ValueError as e:
        # Specific handling for missing sheet or any other value error
            error_message = f"Error7: {str(e)} - The sheet 'Support_Forces' was not found in the Excel file."
            print(error_message)

            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error6", error_message)
            
            # Exit the program after showing the error message
            sys.exit()

        except Exception as e:
            print(f"Error8: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error8", str(e))
            sys.exit()  # Exit the program
            return None

    def retreive_support_number(self,df,support_node):
        try:
            try:
                filtered_df = df[(df.iloc[:,0] == int(support_node))]
            except:
                filtered_df = df[(df.iloc[:,0] == (support_node))]

            if filtered_df.empty:
                raise ValueError(f"Support node {support_node} not found in the data.")

            support_number = filtered_df.iloc[0]["Tag_No"]
            return support_number
        except ValueError as e:
            print(f"Error9: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error9", str(e))
            sys.exit()  # Exit the program
            return None
        except Exception as e:
            print(f"An Unexpected Error10 Occured: {e}")
            return None


    def retrieve_forces_by_load_case(self,df,support_node,load_case):
        print(f"Load Case: {load_case}")
        try:
            filtered_df = df[(df.iloc[:,0] == int(support_node)) & (df["Load_Combination"]== load_case)]
        except:
            filtered_df = df[(df.iloc[:,0] == (support_node)) & (df["Load_Combination"]== load_case)]

        forces = []
        forces.append(int(round(filtered_df.iloc[0]['Global_Force'])))
        forces.append(int(round(filtered_df.iloc[1]['Global_Force'])))
        forces.append(int(round(filtered_df.iloc[2]['Global_Force'])))

        return forces

    def retrieve_nozzle_data(self, nozzle_nodes):   
        for nozzle_node in nozzle_nodes:
            try:
                # Transform the nozzle node data
                self.popup.tranformation_popup(nozzle_node)
                self.transform_string = self.popup.get_transformstring()
                self.local_axes = self.popup.get_local_axes()
                self.process_transformation()
                self.nozzle_dict[nozzle_node] = []


                # Read data from the Excel sheet
                df = pd.read_excel(self.file_path, sheet_name='Forces_Moments')
                df.columns = df.iloc[0]
                df = df[1:]
                df = df.reset_index(drop=True)

                # Attempt to filter the dataframe based on the nozzle node
                try:
                    filtered_df = df[(df.iloc[:, 0] == int(nozzle_node))]
                except:
                    filtered_df = df[(df.iloc[:, 0] == (nozzle_node))]

                # If still empty, throw an error becuase the Nozzle number is not valid
                if filtered_df.empty:
                    raise ValueError(f"Nozzle Node: {nozzle_node} Not Found In Restraint_Loads or Forces_Moments")

                # Proceed to retrieve forces and moments for different load combinations
                forces, moments = self.retrieve_forces_moments_by_load_combinations(df, nozzle_node, "Gravity{1}")
                combined_values = forces + moments
                self.nozzle_dict[nozzle_node].append(combined_values)

                for i in range(self.number_thermal_cases):
                    forces, moments = self.retrieve_forces_moments_by_load_combinations(df, nozzle_node, f"Thermal {i+1}{{1}}")
                    combined_values = forces + moments
                    self.nozzle_dict[nozzle_node].append(combined_values)

                # forces, moments = self.retrieve_forces_moments_by_load_combinations(df, nozzle_node, "Thermal 1{1}")
                # combined_values = forces + moments
                # self.nozzle_dict[nozzle_node].append(combined_values)

                # forces, moments = self.retrieve_forces_moments_by_load_combinations(df, nozzle_node, "Thermal 2{1}")
                # combined_values = forces + moments
                # self.nozzle_dict[nozzle_node].append(combined_values)

                # forces, moments = self.retrieve_forces_moments_by_load_combinations(df, nozzle_node, "Thermal 3{1}")
                # combined_values = forces + moments
                # self.nozzle_dict[nozzle_node].append(combined_values)

                # forces, moments = self.retrieve_forces_moments_by_load_combinations(df, nozzle_node, "Thermal 4{1}")
                # combined_values = forces + moments
                # self.nozzle_dict[nozzle_node].append(combined_values)

            except ValueError as e:
                # Capture the ValueError and log the issue with the nozzle node
                print(f"Error processing nozzle node {nozzle_node}: {str(e)}")
                # Show the error message in a Tkinter popup
                root = tk.Tk()
                root.withdraw()  # Hide the root window
                messagebox.showerror("Error", str(e))
                sys.exit()  # Exit the program
                return None

        return self.nozzle_dict




    def retrieve_forces_moments_by_load_combinations(self,df,nozzle_node,load_combo):
        print(f"Load Combo: {load_combo}")
        try:
            try:
                filtered_df = df[(df.iloc[:,0] == int(nozzle_node)) & (df["Load_Combination"]== load_combo)]
            except:
                filtered_df = df[(df.iloc[:,0] == (nozzle_node)) & (df["Load_Combination"]== load_combo)]
            if filtered_df.empty:
                raise ValueError(f"No Data Found For Nozzle at point: {nozzle_node}")
            forces = filtered_df.iloc[0][["Forces_X","Forces_Y","Forces_Z"]].tolist()
            moments = filtered_df.iloc[0][["Moments_X","Moments_Y","Moments_Z"]].tolist()
            forces = [int(round(x,0)) for x in forces]
            print(f"Forces: {forces}")
            moments = [int(round(x,0)) for x in moments]
            print(f"Moments: {moments}")
            if len(self.transform_string)!=0:
                

                forces, moments = self.transform_nozzle_loads(forces,moments)

            forces = [int(round(x,0)) for x in forces]
            moments = [int(round(x,0)) for x in moments]

            return forces,moments
        except ValueError as e:
            print(f"Error1: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None
        
        except Exception as e:
            print(f"Error2: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None


    def retrieve_local_forces_moments(self,df,nozzle_node,load_combo):
        try:
            try:
                filtered_df = df[(df.iloc[:,0] == int(nozzle_node)) & (df["Combinati…"]== load_combo)]
            except:
                filtered_df = df[(df.iloc[:,0] == (nozzle_node)) & (df["Combinati…"]== load_combo)]
            if filtered_df.empty:
                raise ValueError(f"No Data Found For Nozzle at point: {nozzle_node}")
            forces = filtered_df.iloc[0][["LocalFX lbf","LocalFY lbf","LocalFZ lbf"]].tolist()
            moments = filtered_df.iloc[0][["LocalMX ft-lb","LocalMY ft-lb","LocalMZ ft-lb"]].tolist()
            #print(f"Retrieving {nozzle_node}")
            forces = [int(round(x,0)) for x in forces]
            moments = [int(round(x,0)) for x in moments]
            if len(self.transform_string)!=0:
                

                forces, moments = self.transform_nozzle_loads(forces,moments)

            forces = [int(round(x,0)) for x in forces]
            moments = [int(round(x,0)) for x in moments]

            return forces,moments
        except ValueError as e:
            print(f"Error3: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None
        
        except Exception as e:
            print(f"Error4: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None
    

    def process_transformation(self):
        if self.transform_string:
            swap_order,swap_scalar = [],[]
            self.transform_string = self.transform_string.upper()

            axis_key = {
                    "X":0,
                    "Y":1,
                    "Z":2
                }

            i = 0
            while i<(len(self.transform_string)):
                if self.transform_string[i] == "-":
                    swap_scalar.append(-1)
                    swap_order.append(axis_key[self.transform_string[i+1]])
                    i+=2
                else:
                    swap_scalar.append(1)
                    swap_order.append(axis_key[self.transform_string[i]])
                    i+=1
                self.swap_order = swap_order
                self.swap_scalar = swap_scalar

                
    def transform_nozzle_loads(self,forces,moments):
        new_forces = []
        new_moments = []
        for i in range(len(forces)):
            new_force = (forces[self.swap_order[i]])*self.swap_scalar[i]
            new_forces.append(new_force)
            new_moment = (moments[self.swap_order[i]])*self.swap_scalar[i]
            new_moments.append(new_moment)
        return new_forces, new_moments
    
    def get_number_thermal_cases(self):
        return self.number_thermal_cases