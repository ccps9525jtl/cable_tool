# %%
import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from Model.cable import *
from View.cable_tool_ui import *

basedir = os.getcwd()
# print(basedir)

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'mycompany.myproduct.subproduct.version'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

# %%
class Cable_Tool_Main(QMainWindow):
    def __init__(self, parent = None):
        super(QMainWindow, self).__init__(parent)
        # super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle('Cable Tool')
        self.setWindowIcon(QtGui.QIcon(os.path.join(basedir, 'icons/cable_tool.ico')))
        self.show() 

        ## Define the flag that indicates the current page
        self.is_page_2_port = False
        self.is_page_3_port = False
        self.is_page_4_port = False

        self.cable_file_path = ''

        ## Define netlist path
        self.board_a_netlist = ''
        self.board_b_netlist = ''
        self.board_c_netlist = ''
        self.board_d_netlist = ''
        # self.board_a = Board(self.board_a_netlist)
        # self.board_b = Board(self.board_b_netlist)

        ## Define board name 
        self.board_a_name = ''
        self.board_b_name = ''
        self.board_c_name = ''
        self.board_d_name = ''

        ## Define board components_list
        self.board_a_components_list = []
        self.board_b_components_list = []
        self.board_c_components_list = []
        self.board_d_components_list = []
        self.board_components_list = None

        self.board_list = []
        self.port_list = []

        self.cable_dict = {}

        self.ui.board_a_confirm_button.clicked.connect(self.board_a_browsefile)
        self.ui.board_b_confirm_button.clicked.connect(self.board_b_browsefile)
        self.ui.board_c_confirm_button.clicked.connect(self.board_c_browsefile)
        self.ui.board_d_confirm_button.clicked.connect(self.board_d_browsefile)
        self.ui.cable_browsefile_button.clicked.connect(self.cable_browsefile)
        self.ui.next_button.clicked.connect(self.next_button_clicked)

        self.ui.previous_button_2_port.clicked.connect(self.previous_button_clicked)
        self.ui.previous_button_3_port.clicked.connect(self.previous_button_clicked)
        self.ui.previous_button_4_port.clicked.connect(self.previous_button_clicked)

        self.ui.confirm_button_3_port.clicked.connect(self.confirm_button_clicked)
        self.ui.confirm_button_2_port.clicked.connect(self.confirm_button_clicked)
        self.ui.confirm_button_4_port.clicked.connect(self.confirm_button_clicked)

        ## Page 1 (Main page) tool tips
        self.ui.cable_browsefile_button.setToolTip(f'Browse to select the .xlsx file for cable definitions.')
        self.ui.board_a_confirm_button.setToolTip(f'Browse to select the .zip file for netlist.')
        self.ui.board_b_confirm_button.setToolTip(f'Browse to select the .zip file for netlist.')
        self.ui.board_c_confirm_button.setToolTip(f'Browse to select the .zip file for netlist.')
        self.ui.board_d_confirm_button.setToolTip(f'Browse to select the .zip file for netlist.')
        self.ui.cable_filename.setToolTip(f'Cable file path')
        self.ui.board_a_filename.setToolTip(f'Netlist path')
        self.ui.board_b_filename.setToolTip(f'Netlist path')
        self.ui.board_c_filename.setToolTip(f'Netlist path')
        self.ui.board_d_filename.setToolTip(f'Netlist path')

        ## Page 2 (2 port) tool tips
        self.ui.p1_lineEdit_2_port.setToolTip(f'Enter a location.')
        self.ui.p2_lineEdit_2_port.setToolTip(f'Enter a location.')
        self.ui.p1_board_comboBox_2_port.setToolTip("Select the board that connects to Port 1.")
        self.ui.p2_board_comboBox_2_port.setToolTip("Select the board that connects to Port 2.")

        ## Page 3 (3 port) tool tips
        self.ui.p1_lineEdit_3_port.setToolTip(f'Enter a location.')
        self.ui.p2_lineEdit_3_port.setToolTip(f'Enter a location.')
        self.ui.p3_lineEdit_3_port.setToolTip(f'Enter a location.')
        self.ui.p1_board_comboBox_3_port.setToolTip("Select the board that connects to Port 1.")
        self.ui.p2_board_comboBox_3_port.setToolTip("Select the board that connects to Port 2.")
        self.ui.p3_board_comboBox_3_port.setToolTip("Select the board that connects to Port 3.")

        ## Page 4 (4 port) tool tips
        self.ui.p1_lineEdit_4_port.setToolTip(f'Enter a location.')
        self.ui.p2_lineEdit_4_port.setToolTip(f'Enter a location.')
        self.ui.p3_lineEdit_4_port.setToolTip(f'Enter a location.')
        self.ui.p4_lineEdit_4_port.setToolTip(f'Enter a location.')
        self.ui.p1_board_comboBox_4_port.setToolTip("Select the board that connects to Port 1.")
        self.ui.p2_board_comboBox_4_port.setToolTip("Select the board that connects to Port 2.")
        self.ui.p3_board_comboBox_4_port.setToolTip("Select the board that connects to Port 3.")
        self.ui.p4_board_comboBox_4_port.setToolTip("Select the board that connects to Port 4.")


    def board_a_browsefile(self):
        fname=QFileDialog.getOpenFileName(self, 'Open file', '/Users/tony/Downloads/power_net_tool_gui/netlist.zip', 'Zip file (*.zip)')
        self.ui.board_a_filename.setText(fname[0])  

    def board_b_browsefile(self):
        fname=QFileDialog.getOpenFileName(self, 'Open file', '/Users/tony/Downloads/power_net_tool_gui/netlist.zip', 'Zip file (*.zip)')
        self.ui.board_b_filename.setText(fname[0])  

    def board_c_browsefile(self):
        fname=QFileDialog.getOpenFileName(self, 'Open file', '/Users/tony/Downloads/power_net_tool_gui/netlist.zip', 'Zip file (*.zip)')
        self.ui.board_c_filename.setText(fname[0])  

    def board_d_browsefile(self):
        fname=QFileDialog.getOpenFileName(self, 'Open file', '/Users/tony/Downloads/power_net_tool_gui/netlist.zip', 'Zip file (*.zip)')
        self.ui.board_d_filename.setText(fname[0])  

    def cable_browsefile(self):
        fname=QFileDialog.getOpenFileName(self, 'Open file', '/Users/tony/Downloads/power_net_tool_gui/netlist.zip', '*.xlsx')
        self.ui.cable_filename.setText(fname[0])   

    def next_button_clicked(self):
        if (verify_zipped_netlist(self.ui.board_a_filename.text()) and 
            verify_zipped_netlist(self.ui.board_b_filename.text()) and 
            verify_zipped_netlist(self.ui.board_c_filename.text()) and
            verify_zipped_netlist(self.ui.board_d_filename.text()) and
            self.ui.cable_filename.text() != ''):

            # print(self.board_a_netlist)
            self.board_a_netlist = self.ui.board_a_filename.text()
            self.board_b_netlist = self.ui.board_b_filename.text()
            self.board_c_netlist = self.ui.board_c_filename.text()
            self.board_d_netlist = self.ui.board_d_filename.text()
            self.cable_file_path =  self.ui.cable_filename.text()

            self.board_a = Board(self.board_a_netlist)
            self.board_b = Board(self.board_b_netlist)
            self.board_c = Board(self.board_c_netlist)
            self.board_d = Board(self.board_d_netlist)

            self.board_a_name = self.board_a.board_name
            self.board_b_name = self.board_b.board_name
            self.board_c_name = self.board_c.board_name
            self.board_d_name = self.board_d.board_name

            self.board_a_components_list = self.board_a.get_components_list()
            self.board_b_components_list = self.board_b.get_components_list()
            self.board_c_components_list = self.board_c.get_components_list()
            self.board_d_components_list = self.board_d.get_components_list()

            self.board_components_list = [self.board_a_components_list, self.board_b_components_list, self.board_c_components_list, self.board_d_components_list]

            self.board_list = [self.board_a_name, self.board_b_name, self.board_c_name, self.board_d_name]
            print(f'\033[92m    Boards {self.board_a_name}, {self.board_b_name}, {self.board_c_name} and {self.board_d_name} have been imported.\033[0m')

            cable = Cable(self.cable_file_path)
            self.port_list = cable.port_list
            print(f'\033[92m    From {cable.cable_file_name} find {cable.number_of_port} port {self.port_list}\n\033[0m')
            print(cable.df_cable)

            ## 4-Port case
            if len(cable.port_list) == 4:
                self.ui.p1_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p3_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p4_board_comboBox_4_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_4_port)
                self.is_page_4_port = True

            ## 3-Port case
            elif len(cable.port_list) == 3:
                self.ui.p1_board_comboBox_3_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_3_port.addItems(self.board_list)
                self.ui.p3_board_comboBox_3_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_3_port)
                self.is_page_3_port = True

            ## 2-Port case
            else:
                self.ui.p1_board_comboBox_2_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_2_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_2_port)
                self.is_page_2_port = True
        
        elif (verify_zipped_netlist(self.ui.board_a_filename.text()) and
              verify_zipped_netlist(self.ui.board_b_filename.text()) and 
              verify_zipped_netlist(self.ui.board_c_filename.text()) and 
              self.ui.board_d_filename.text() == "" and 
              self.ui.cable_filename.text() != ''):
            
            self.board_a_netlist = self.ui.board_a_filename.text()
            self.board_b_netlist = self.ui.board_b_filename.text()
            self.board_c_netlist = self.ui.board_c_filename.text()

            self.cable_file_path =  self.ui.cable_filename.text()

            self.board_a = Board(self.board_a_netlist)
            self.board_b = Board(self.board_b_netlist)
            self.board_c = Board(self.board_c_netlist)

            self.board_a_name = self.board_a.board_name
            self.board_b_name = self.board_b.board_name
            self.board_c_name = self.board_c.board_name

            self.board_a_components_list = self.board_a.get_components_list()
            self.board_b_components_list = self.board_b.get_components_list()
            self.board_c_components_list = self.board_c.get_components_list()

            self.board_components_list = [self.board_a_components_list, self.board_b_components_list, self.board_c_components_list]

            self.board_list = [self.board_a_name, self.board_b_name, self.board_c_name]
            print(f'\033[92m    Boards {self.board_a_name}, {self.board_b_name} and {self.board_c_name} have been imported.\033[0m')

            cable = Cable(self.cable_file_path)
            self.port_list = cable.port_list
            
            print(f'\033[92m    From {cable.cable_file_name} find {cable.number_of_port} port {self.port_list}\n\033[0m')
            print(cable.df_cable)

            ## 4-Port case
            if len(cable.port_list) == 4:
                self.ui.p1_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p3_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p4_board_comboBox_4_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_4_port)
                self.is_page_4_port = True

            ## 3-Port case
            elif len(cable.port_list) == 3:
                self.ui.p1_board_comboBox_3_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_3_port.addItems(self.board_list)
                self.ui.p3_board_comboBox_3_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_3_port)
                self.is_page_3_port = True

            ## 2-Port case
            else:
                self.ui.p1_board_comboBox_2_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_2_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_2_port)
                self.is_page_2_port = True

        elif (verify_zipped_netlist(self.ui.board_a_filename.text()) and
              verify_zipped_netlist(self.ui.board_b_filename.text()) and  
              self.ui.board_c_filename.text() == "" and 
              self.ui.board_d_filename.text() == "" and 
              self.ui.cable_filename.text() != ''):
            
            self.board_a_netlist = self.ui.board_a_filename.text()
            self.board_b_netlist = self.ui.board_b_filename.text()

            self.cable_file_path =  self.ui.cable_filename.text()

            self.board_a = Board(self.board_a_netlist)
            self.board_b = Board(self.board_b_netlist)

            self.board_a_name = self.board_a.board_name
            self.board_b_name = self.board_b.board_name

            self.board_a_components_list = self.board_a.get_components_list()
            self.board_b_components_list = self.board_b.get_components_list()

            self.board_components_list = [self.board_a_components_list, self.board_b_components_list]

            self.board_list = [self.board_a_name, self.board_b_name]
            print(f'\033[92m    Boards {self.board_a_name} and {self.board_b_name} have been imported.\033[0m')

            cable = Cable(self.cable_file_path)
            self.port_list = cable.port_list
            print(f'\033[92m    From {cable.cable_file_name} find {cable.number_of_port} port {self.port_list}\n\033[0m')
            print(cable.df_cable)

            ## 4-Port case
            if len(cable.port_list) == 4:
                self.ui.p1_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p3_board_comboBox_4_port.addItems(self.board_list)
                self.ui.p4_board_comboBox_4_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_4_port)
                self.is_page_4_port = True

            ## 3-Port case
            elif len(cable.port_list) == 3:
                self.ui.p1_board_comboBox_3_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_3_port.addItems(self.board_list)
                self.ui.p3_board_comboBox_3_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_3_port)
                self.is_page_3_port = True

            ## 2-Port case
            else:
                self.ui.p1_board_comboBox_2_port.addItems(self.board_list)
                self.ui.p2_board_comboBox_2_port.addItems(self.board_list)
                self.ui.stackedWidget.setCurrentWidget(self.ui.page_2_port)
                self.is_page_2_port = True

        else:
            print("\033[31mERROR: The configurations failed to operate correctly.\033[0m")


            
    
    def previous_button_clicked(self):
        self.board_list.clear
        self.port_list.clear
        self.ui.p1_board_comboBox_4_port.clear()
        self.ui.p2_board_comboBox_4_port.clear()
        self.ui.p3_board_comboBox_4_port.clear()
        self.ui.p4_board_comboBox_4_port.clear()

        self.ui.p1_board_comboBox_3_port.clear()
        self.ui.p2_board_comboBox_3_port.clear()
        self.ui.p3_board_comboBox_3_port.clear()

        self.ui.p1_board_comboBox_2_port.clear()
        self.ui.p2_board_comboBox_2_port.clear()

        self.ui.p1_lineEdit_2_port.clear()
        self.ui.p2_lineEdit_2_port.clear()

        self.ui.p1_lineEdit_3_port.clear()
        self.ui.p2_lineEdit_3_port.clear()
        self.ui.p3_lineEdit_3_port.clear()

        self.ui.p1_lineEdit_4_port.clear()
        self.ui.p2_lineEdit_4_port.clear()
        self.ui.p3_lineEdit_4_port.clear()
        self.ui.p4_lineEdit_4_port.clear()

        self.is_page_2_port = False
        self.is_page_3_port = False
        self.is_page_4_port = False

        self.cable_dict = {}

        self.ui.stackedWidget.setCurrentWidget(self.ui.page_1)

    def check_port_wiring(self, port, combo_box, line_edit):
        ## Board comboBox  = board_a
        if combo_box.currentText() == self.board_a_name:
            # print(f"It is {combo_box.currentText()}, location is {line_edit.text()}")
            # print(self.board_a_components_list)
            if line_edit.text() in self.board_a_components_list:
                self.cable_dict.update({port: (self.board_a_netlist, line_edit.text())})
                print(f"\033[92m{port} with a valid connection.\033[0m")
                return True
            else:
                print(f"\033[91mERROR: {port} - The location \"{line_edit.text()}\" does not exist in the {combo_box.currentText()}.\033[0m")
                return False

        ## Board comboBox  = board_b
        elif combo_box.currentText() == self.board_b_name:
            # print(f"It is {combo_box.currentText()}, location is {line_edit.text()}")
            # print(self.board_a_components_list)
            if line_edit.text() in self.board_b_components_list:
                self.cable_dict.update({port: (self.board_b_netlist, line_edit.text())})
                print(f"\033[92m{port} with a valid connection.\033[0m")
                return True
            else:
                print(f"\033[91mERROR: {port} - The location \"{line_edit.text()}\" does not exist in the {combo_box.currentText()}.\033[0m")
                return False
            
        ## Board comboBox  = board_c
        elif combo_box.currentText() == self.board_c_name:
            # print(f"It is {combo_box.currentText()}, location is {line_edit.text()}")
            # print(self.board_a_components_list)
            if line_edit.text() in self.board_c_components_list:
                self.cable_dict.update({port: (self.board_c_netlist, line_edit.text())})
                print(f"\033[92m{port} with a valid connection.\033[0m")
                return True
            else:
                print(f"\033[91mERROR: {port} - The location \"{line_edit.text()}\" does not exist in the {combo_box.currentText()}.\033[0m")
                return False
            
        ## Board comboBox  = board_d
        else:
            # print(f"It is {combo_box_3_port.currentText()}, location is {line_edit.text()}")
            if line_edit.text() in self.board_d_components_list:
                self.cable_dict.update({port: (self.board_d_netlist, line_edit.text())})
                print(f"\033[92m{port} with a valid connection.\033[0m")
                return True
            else:
                print(f"\033[91mERROR: {port} - The location \"{line_edit.text()}\" does not exist in the {combo_box.currentText()}.\033[0m")
                return False
            
    def confirm_button_clicked(self):
        if self.is_page_4_port == True:
            p1_wiring_stat = self.check_port_wiring('P1', self.ui.p1_board_comboBox_4_port, self.ui.p1_lineEdit_4_port)
            p2_wiring_stat = self.check_port_wiring('P2', self.ui.p2_board_comboBox_4_port, self.ui.p2_lineEdit_4_port)
            p3_wiring_stat = self.check_port_wiring('P3', self.ui.p3_board_comboBox_4_port, self.ui.p3_lineEdit_4_port)
            p4_wiring_stat = self.check_port_wiring('P4', self.ui.p4_board_comboBox_4_port, self.ui.p4_lineEdit_4_port)
            print("\n")

            if (p1_wiring_stat and p2_wiring_stat and p3_wiring_stat and p4_wiring_stat):
                # print(self.cable_dict)
                cable = Cable(self.cable_file_path)
                # cable.generate_board_connection(**self.cable_dict)
                excel_file_name = Excel_former().generate_cable_routing_report(cable.generate_board_connection(**self.cable_dict), 'cable')
                Excel_former().friendly_cable_report(excel_file_name)
                Excel_former().add_excel_pass_fail_condition(excel_file_name, cable.excel_pass_fail_condition)
                print("\033[42m\033[30mReport generated.\033[0m\n\n")

        elif self.is_page_3_port == True:
            p1_wiring_stat = self.check_port_wiring('P1', self.ui.p1_board_comboBox_3_port, self.ui.p1_lineEdit_3_port)
            p2_wiring_stat = self.check_port_wiring('P2', self.ui.p2_board_comboBox_3_port, self.ui.p2_lineEdit_3_port)
            p3_wiring_stat = self.check_port_wiring('P3', self.ui.p3_board_comboBox_3_port, self.ui.p3_lineEdit_3_port)
            print("\n")

            if (p1_wiring_stat and p2_wiring_stat and p3_wiring_stat):
                # print(self.cable_dict)
                cable = Cable(self.cable_file_path)
                # cable.generate_board_connection(**self.cable_dict)
                excel_file_name = Excel_former().generate_cable_routing_report(cable.generate_board_connection(**self.cable_dict), 'cable')
                Excel_former().friendly_cable_report(excel_file_name)
                Excel_former().add_excel_pass_fail_condition(excel_file_name, cable.excel_pass_fail_condition)
                print("\033[42m\033[30mReport generated.\033[0m\n\n")
                
        elif self.is_page_2_port == True:
            p1_wiring_stat = self.check_port_wiring('P1', self.ui.p1_board_comboBox_2_port, self.ui.p1_lineEdit_2_port)
            p2_wiring_stat = self.check_port_wiring('P2', self.ui.p2_board_comboBox_2_port, self.ui.p2_lineEdit_2_port)

            if (p1_wiring_stat and p2_wiring_stat):
                # print(self.cable_dict)
                cable = Cable(self.cable_file_path)
                # cable.generate_board_connection(**self.cable_dict)
                excel_file_name = Excel_former().generate_cable_routing_report(cable.generate_board_connection(**self.cable_dict), 'cable')
                Excel_former().friendly_cable_report(excel_file_name)
                Excel_former().add_excel_pass_fail_condition(excel_file_name, cable.excel_pass_fail_condition)
                print("\033[42m\033[30mReport generated.\033[0m\n\n")


# %%
def main():  
    print("\nCABLE TOOL LOG") 
    app = QtWidgets.QApplication(sys.argv)
    # app.setWindowIcon(QtGui.QIcon(os.path.join(basedir, 'icons/cable_tool.ico')))
    main = Cable_Tool_Main()
    app.exec_() 

if __name__ == "__main__":
    main()


