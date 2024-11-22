from docx import Document
import Menu
import DataRetriever
import Editor
class Controller:
    def __init__(self):
        self.menu = Menu.Menu(self)
        #self.data_retriever = DataRetriever.DataRetriever()
        #self.data_parser = DataParser.DataParser()
        #self.table_editor = Editor.Editor()

    def run(self):
        self.menu.run()

    def handle_user_input(self, file_path, nozzle_nodes, support_nodes,output_path):
        self.data_retriever = DataRetriever.DataRetriever(file_path)
        self.table_editor = Editor.Editor(output_path)
        print("Done")
        support_data = self.data_retriever.retrieve_support_data(support_nodes=support_nodes)
        nozzle_data = self.data_retriever.retrieve_nozzle_data(nozzle_nodes=nozzle_nodes)
        self.table_editor.update_support_table(support_dict=support_data)
        self.table_editor.update_nozzle_tables(nozzle_dict=nozzle_data)
        self.table_editor.save_document()

