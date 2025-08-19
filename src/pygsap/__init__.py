import win32com.client
import time
import subprocess
import json
from numpy import array, char
import os

class SAPConnection:

    def __init__(self, system_name, user, password, language="EN", n_sessions=1, max_sessions=6):
        # Store connection credentials and state
        self.system_name = system_name
        self.user = user
        self.password = password
        self.language = language

        self.session_type = None
        self.sap_app = None
        self.sap_connection = None
        self.sap_session = None
        self.is_connected = False

        self.sap_session = None
        self.main_page_label = None

        # Close SAP in case it's open
        self.close_sap()

        # Wait for SAP to close
        time.sleep(1)

        # Define number of sessions
        self.n_sessions = min(max(n_sessions, 1), max_sessions)

        # Dictionary to store important session characteristics
        self.session_characteristics = {
            "system_name": None,
            "client": None,
            "user": None,
            "program_name": None,
            "transaction_code": None,
        }
        
        # Open SAP
        return self._open_sap()
    
    def _open_sap(self):
        # Open SAP
        subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", shell=False)

        # Wait for SAP to open
        time.sleep(1)

        # Connect
        timer = 0
        while True:
            try:
                sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
                break
            except Exception as e:
                timer += 1
                if timer == 120:
                    raise TimeoutException("SAP Logon was taking too much time to open")
                time.sleep(1)
        conn = sap.OpenConnection(self.system_name)
        session = conn.Children(0)

        session.findbyId("wnd[0]/usr/txtRSYST-BNAME").text = self.user
        session.findbyId("wnd[0]/usr/pwdRSYST-BCODE").text = self.password

        session.findbyId("wnd[0]").sendVKey(0)
        self.main_page_label = session.findbyId("wnd[0]").text
        if self.n_sessions != 1:
            sessions = [session]
            for i in range(self.n_sessions - 1):
                sessions[0].createSession()        
            time.sleep(3)
            for i in range(1, self.n_sessions):
                sessions += [conn.Children(i)]   

            self.sap_session = sessions
            self.session_type = "List_of_sessions"
            self.is_connected = True
        else:
            self.sap_session = session
            self.session_type = "Single_session"
            self.is_connected = True
        
    def close_sap(self):
        try:
            subprocess.Popen(r"taskkill /IM saplogon.exe /F", shell=False)
        except:
            pass
    
    def _update_session_characteristics(self):
        """
        Internal method to update the session characteristics dictionary.
        """
        if self.is_connected:
            self.session_characteristics["system_name"] = self.sap_session.info.systemName
            self.session_characteristics["client"] = self.sap_session.info.client
            self.session_characteristics["user"] = self.sap_session.info.user
            self.session_characteristics["program_name"] = self.sap_session.info.program
            self.session_characteristics["transaction_code"] = self.sap_session.info.transaction

    def get_session_info(self):
        self._update_session_characteristics()
        return self.session_characteristics
    
    # def execute_tcode(self, tcode):
    #     if self.is_connected and self.sap_session:
    #         try:
    #             self.sap_session.findById("wnd[0]/tbar[0]/okcd").Text = tcode
    #             self.sap_session.findById("wnd[0]").sendVKey(0)
    #             print(f"Executing transaction code: {tcode}")
    #             time.sleep(1) # Wait for screen to load
    #             self._update_session_characteristics()
    #         except Exception as e:
    #             print(f"Failed to execute transaction code '{tcode}': {e}")
    #     else:
    #         print("Cannot execute T-code. Not connected to SAP.")

    def StartTransaction(self, tcode):
        if self.is_connected and self.sap_session:
            try:
                self.sap_session.StartTransaction(tcode)
                self._update_session_characteristics()
            except Exception as e:
                print(f"Failed to execute transaction code '{tcode}': {e}")
        else:
            print("Cannot execute T-code. Not connected to SAP.")

    def EndTransaction(self):
        if self.is_connected and self.sap_session:
            try:
                self.sap_session.EndTransaction()
                self._update_session_characteristics()
            except Exception as e:
                print(f"Failed to end transaction: {e}")
        else:
            print("Cannot execute T-code. Not connected to SAP.")

    def name_validation(self, name, window_number=0, case_sesitive=True):
        if case_sesitive:
            if self.sap_session.findById(f"wnd[{window_number}]").text != name:
                raise ValueError("Validation error")
        else:
            if self.sap_session.findById(f"wnd[{window_number}]").text.upper() != name.upper():
                raise ValueError("Validation error")
            
    def find_input_by_label(self, label_text):
        if not self.is_connected or not self.sap_session:
            print("Not connected to SAP.")
            return None

        try:
            active_window = self.sap_session.ActiveWindow
            for i in range(active_window.Children.Count):
                element = active_window.Children(i)
                # Check if the element is a GuiLabel and if its text matches
                if element.Type == "GuiLabel" and element.Text == label_text:
                    # The next element in the collection is usually the input field
                    if i + 1 < active_window.Children.Count:
                        input_field = active_window.Children(i + 1)
                        print(f"Found input field for label '{label_text}'.")
                        return input_field
            print(f"Input field for label '{label_text}' not found.")
            return None
        except Exception as e:
            print(f"An error occurred while searching for input field: {e}")
            return None
    
    def is_window_open(self, window_id):
        if not self.is_connected or not self.sap_session:
            return False

        try:
            # Try to find the window. If it doesn't exist, an exception will be raised.
            self.sap_session.findById(window_id)
            return True
        except Exception:
            return False
        
    # def go_to_main_page(self):
    #     # Check if an extra window is opened
    #     if self.is_window_open("wnd[1]"):
    #         self.close_element("wnd[1]")
    #         if self.is_window_open("wnd[1]") and self._is_exit_box("wnd[1]"):
    #             try:
    #                 self.click_on("wnd[1]/usr/btnBUTTON_YES")
    #             except:
    #                 pass

    #     # Check if we already are in the main page
    #     if self.sap_session.findById("wnd[0]").text == self.main_page_label:
    #         return 
    #     else:
    #         # Go back to main page
    #         self.sap_session.findById("wnd[0]/tbar[0]/btn[3]").press()
    #         self.go_to_main_page()

    def click_on(self, element):
        try:
            self.sap_session.findById(element).press()
        except:
            self.sap_session.findById(element).select()

    def close_element(self, element):
        self.sap_session.findById(element).close()

    def get_text(self, element):
        if type(element) == list:
            return [self.get_text(x) for x in element]
        return self.sap_session.findById(element).text
    
    def find_all_elemts(self, root):
        data = json.loads(self.sap_session.getObjectTree(root))
        return self._find_all_elemts(data, id_list=None)

    def _find_all_elemts(self, data, id_list=None):
        if id_list is None:
            id_list = []

        if isinstance(data, dict):
            for key, value in data.items():
                if key == "Id":
                    id_list.append(value)
                self._find_all_elemts(value, id_list)
        elif isinstance(data, list):
            for item in data:
                self._find_all_elemts(item, id_list)
        
        return id_list
    
    def _is_exit_box(self, element):
        return any((x.upper().find("EXIT") != -1) and (x.find("?") != -1) for x in map(self.get_text, self.find_all_elemts(element)))

    def find_element_by_text(self, text, root='wnd[0]', casesensitive=False):
        # Creating arrays
        id_list = array(self.find_all_elemts(root))
        description_list = array(self.get_text(self.find_all_elemts(root)))

        # Case sensitive option
        if not casesensitive:
            description_list = char.lower(description_list)
            text = text.lower()

        # Create a boolean mask by checking where the result is not equal to -1.
        mask = char.find(description_list, text) != -1

        # Return result
        if sum(mask) == 0:
            return "Not found"
        elif sum(mask) == 1:
            return id_list[mask].tolist()[0]
        else:
            return id_list[mask].tolist()
    
    def execute(self):
        self.sap_session.findById("wnd[0]").sendVKey(8)

    def extract_excel_report(self, file_name, file_path):
        spreadsheet_b = self.find_element_by_text("spreadsheet")
        if (spreadsheet_b != "Not found") and (type(spreadsheet_b) != list):
            # Export spreadsheet
            self.click_on(spreadsheet_b)

            # Validate wnd
            if (self.find_element_by_text("Directory", "wnd[1]") == "Not found") or (self.find_element_by_text("File Name", "wnd[1]") == "Not found"):
                # Select XLSX format
                self.sap_session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "31"

                # Run download action
                self.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()

            # Set report name and path
            self.sap_session.findById("wnd[1]/usr/ctxtDY_PATH").text = file_path
            self.sap_session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name

            # If file already exist we click on Replace, otherwise we click on Generate
            if os.path.exists(file_path + "\\" + file_name):
                self.sap_session.findById("wnd[1]/tbar[0]/btn[11]").press()  # btn[11] = "Replace"
            else:
                self.sap_session.findById("wnd[1]/tbar[0]/btn[0]").press()   # btn[0] = "Generate"

            # Close excel
            try:
                subprocess.Popen(r"taskkill /IM excel.exe /F", shell=False)
            except:
                pass
    
    def set_text(self, element, text):
        self.sap_session.findById(element).text = text

    def set_multiple_text(self, text_element_dict):
        for i in text_element_dict.keys():
            self.set_text(i, text_element_dict[i])

class TimeoutException(Exception):
    pass