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

            for support_node in support_nodes:
                self.support_dict[support_node] = []

                
                support_number = self.retreive_support_number(df=df, support_node=support_node)

                if support_number is None:
                    continue

                self.support_dict[support_node].append(support_number)


                global_force_value = self.retrieve_max_force_by_load_case(df=df, support_node= support_node,load_case = 'GRT{1}')
                self.support_dict[support_node].append(global_force_value)

                global_force_value = self.retrieve_max_force_by_load_case(df=df, support_node= support_node,load_case = 'GRT{2}')
                self.support_dict[support_node].append(global_force_value)

            
                if abs(self.support_dict[support_node][1])<abs(self.support_dict[support_node][2]):
                    temp_value = self.support_dict[support_node][2]
                    self.support_dict[support_node][2] = self.support_dict[support_node][1]
                    self.support_dict[support_node][1] = temp_value

            #print(self.support_dict)
            return self.support_dict
        except Exception as e:
            print(f"Error: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
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
            print(f"Error: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None
        except Exception as e:
            print(f"An Unexpected Error Occured: {e}")
            return None


    def retrieve_max_force_by_load_case(self,df,support_node,load_case):
        try:
            filtered_df = df[(df.iloc[:,0] == int(support_node)) & (df["Load_Combination"]== load_case)]
        except:
            filtered_df = df[(df.iloc[:,0] == (support_node)) & (df["Load_Combination"]== load_case)]
        # Find the row with the max absolute value in the 12th column (Global_Force)
        max_index = filtered_df['Global_Force'].abs().idxmax()
        # Now get the value from the original DataFrame using the max index
        global_force_value = df.loc[max_index,'Global_Force']
        global_force_value = int(round(global_force_value,0))
        return global_force_value

    def retrieve_nozzle_data(self, nozzle_nodes):   
        for nozzle_node in nozzle_nodes:
            try:
                # Transform the nozzle node data
                self.popup.tranformation_popup(nozzle_node)
                self.transform_string = self.popup.get_transformstring()
                self.local_axes = self.popup.get_local_axes()
                self.process_transformation()
                self.nozzle_dict[nozzle_node] = []

                if self.local_axes == 1:
                    df = pd.read_excel(self.file_path,sheet_name ='Local_Axes')
                    df = df.iloc[:,1:]
                    df=df.reset_index(drop=True)
                    try:
                        filtered_df = df[(df.iloc[:, 0] == int(nozzle_node))]
                    except:
                        filtered_df = df[(df.iloc[:, 0] == (nozzle_node))]

                    if filtered_df.empty:
                        raise ValueError(f"Nozzle Node: {nozzle_node} Not Found In Local Axes")
                    
                    forces, moments = self.retrieve_local_forces_moments(df, nozzle_node, "Gravity{1}")
                    combined_values = forces + moments
                    self.nozzle_dict[nozzle_node].append(combined_values)

                    forces, moments = self.retrieve_local_forces_moments(df, nozzle_node, "Thermal  1…")
                    combined_values = forces + moments
                    self.nozzle_dict[nozzle_node].append(combined_values)

                    forces, moments = self.retrieve_local_forces_moments(df, nozzle_node, "Thermal  2…")
                    combined_values = forces + moments
                    self.nozzle_dict[nozzle_node].append(combined_values)

                else:
                    # Read data from the Excel sheet
                    df = pd.read_excel(self.file_path, sheet_name='Restraint_Loads')
                    df.columns = df.iloc[0]
                    df = df[1:]
                    df = df.reset_index(drop=True)

                    # Attempt to filter the dataframe based on the nozzle node
                    try:
                        filtered_df = df[(df.iloc[:, 0] == int(nozzle_node))]
                    except:
                        filtered_df = df[(df.iloc[:, 0] == (nozzle_node))]

                    # If the filtered dataframe for Restraint_Loads is empty, try Forces_Moments
                    if filtered_df.empty:
                        df = pd.read_excel(self.file_path, sheet_name='Forces_Moments')
                        df.columns = df.iloc[0]
                        df = df[1:]
                        df = df.reset_index(drop=True)

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

                    forces, moments = self.retrieve_forces_moments_by_load_combinations(df, nozzle_node, "Thermal 1{1}")
                    combined_values = forces + moments
                    self.nozzle_dict[nozzle_node].append(combined_values)

                    forces, moments = self.retrieve_forces_moments_by_load_combinations(df, nozzle_node, "Thermal 2{1}")
                    combined_values = forces + moments
                    self.nozzle_dict[nozzle_node].append(combined_values)

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
            moments = [int(round(x,0)) for x in moments]
            if len(self.transform_string)!=0:
                

                forces, moments = self.transform_nozzle_loads(forces,moments)

            forces = [int(round(x,0)) for x in forces]
            moments = [int(round(x,0)) for x in moments]

            return forces,moments
        except ValueError as e:
            print(f"Error: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None
        
        except Exception as e:
            print(f"Error: {e}")
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
            print(f"Error: {e}")
            # Show the error message in a Tkinter popup
            root = tk.Tk()
            root.withdraw()  # Hide the root window
            messagebox.showerror("Error", str(e))
            sys.exit()  # Exit the program
            return None
        
        except Exception as e:
            print(f"Error: {e}")
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