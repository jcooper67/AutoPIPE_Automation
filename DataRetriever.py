import pandas as pd

class DataRetriever:

    def __init__(self,file_path,user_transform):
        self.file_path = file_path
        self.support_dict = {}
        self.nozzle_dict = {}
        self.transform_string = user_transform
        self.process_transformation()
    def retrieve_support_data(self,support_nodes):
        #Open the document and Set Up the Data Frame
        print("Executing Retrieval")
        df = pd.read_excel(self.file_path,sheet_name = "Support_Forces")
        df.columns = df.iloc[0]
        df = df [1:]
        df = df.reset_index(drop = True)

        for support_node in support_nodes:
            self.support_dict[support_node] = []

            
            support_number = self.retreive_support_number(df=df, support_node=support_node)
            self.support_dict[support_node].append(support_number)


            global_force_value = self.retrieve_max_force_by_load_case(df=df, support_node= support_node,load_case = 'GRT{1}')
            self.support_dict[support_node].append(global_force_value)

            global_force_value = self.retrieve_max_force_by_load_case(df=df, support_node= support_node,load_case = 'GRT{2}')
            self.support_dict[support_node].append(global_force_value)

        
            if abs(self.support_dict[support_node][1])<abs(self.support_dict[support_node][2]):
                temp_value = self.support_dict[support_node][2]
                self.support_dict[support_node][2] = self.support_dict[support_node][1]
                self.support_dict[support_node][1] = temp_value

        print(self.support_dict)
        return self.support_dict

    def retreive_support_number(self,df,support_node):
        filtered_df = df[(df.iloc[:,0] == int(support_node))]
        support_number = filtered_df.iloc[0]["Tag_No"]
        return support_number


    def retrieve_max_force_by_load_case(self,df,support_node,load_case):
        filtered_df = df[(df.iloc[:,0] == int(support_node)) & (df["Load_Combination"]== load_case)]

        # Find the row with the max absolute value in the 12th column (Global_Force)
        max_index = filtered_df['Global_Force'].abs().idxmax()
        # Now get the value from the original DataFrame using the max index
        global_force_value = df.loc[max_index,'Global_Force']
        global_force_value = round(global_force_value,0)
        return global_force_value

    def retrieve_nozzle_data(self,nozzle_nodes):
        df = pd.read_excel(self.file_path, sheet_name='Restraint_Loads')
        df.columns = df.iloc[0]
        df = df[1:]
        df = df.reset_index(drop = True)

        for nozzle_node in nozzle_nodes:
            self.nozzle_dict[nozzle_node] = []
            #Gravity Load Combo
            forces,moments = self.retrieve_forces_moments_by_load_combinations(df,nozzle_node,"Gravity{1}")
            combined_values = forces+moments
            self.nozzle_dict[nozzle_node].append(combined_values)
            #Thermal 1 Load Combo
            forces,moments = self.retrieve_forces_moments_by_load_combinations(df,nozzle_node,"Thermal 1{1}")
            combined_values = forces+moments
            self.nozzle_dict[nozzle_node].append(combined_values)
            #Thermal 2 Load Combo
            forces,moments = self.retrieve_forces_moments_by_load_combinations(df,nozzle_node,"Thermal 2{1}")
            combined_values = forces+moments
            self.nozzle_dict[nozzle_node].append(combined_values)
        
        print(self.nozzle_dict)
        return self.nozzle_dict


    def retrieve_forces_moments_by_load_combinations(self,df,nozzle_node,load_combo):
        filtered_df = df[(df.iloc[:,0] == int(nozzle_node)) & (df["Load_Combination"]== load_combo)]
        forces = filtered_df.iloc[0][["Forces_X","Forces_Y","Forces_Z"]].tolist()
        moments = filtered_df.iloc[0][["Moments_X","Moments_Y","Moments_Z"]].tolist()

        forces = [round(x,0) for x in forces]
        moments = [round(x,0) for x in moments]
        print(f"Transform String : {self.transform_string}")
        if len(self.transform_string)!=0:

            forces, moments = self.transform_nozzle_loads(forces,moments)

        forces = [round(x,0) for x in forces]
        moments = [round(x,0) for x in moments]

        return forces,moments
    

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