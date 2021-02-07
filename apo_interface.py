import sys
import win32com.client
import locale   #https://www.delftstack.com/de/howto/python/how-to-convert-string-to-float-or-int/


class Logistical_Material_Data:
    def __init__(self,fpc):
        locale.setlocale(locale.LC_ALL, 'nl_NL')
        print("Reading material data from KHP...")
        self.read_APO_data(fpc)
        print(f"{fpc} {self.product_description} \ndone")
        print("\n")



    def read_APO_data(self,material):
        try:

            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not type(SapGuiAuto) == win32com.client.CDispatch:
                return

            application = SapGuiAuto.GetScriptingEngine
            if not type(application) == win32com.client.CDispatch:
                SapGuiAuto = None
                return

            #find and connect to APO session
            for i in range(0, len(application.Children)):
                connection = application.Children(i)
                if not type(connection) == win32com.client.CDispatch:
                    application = None
                    SapGuiAuto = None
                    return
                session = connection.Children(0)
                if not type(session) == win32com.client.CDispatch:
                    connection = None
                    application = None
                    SapGuiAuto = None
                    return
                if session.info.systemname == 'KHP':
                    break
            #print(session.info.systemname) # could print the system name ('F6P', 'KHP', etc)


            session.findById("wnd[0]").resizeWorkingPane(173, 36, 0)
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n/SAPAPO/MAT1"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/ctxt/SAPAPO/MATIO-MATNR").text = material
            session.findById("wnd[0]/usr/radDV_GLOBAL_DATA").select()
            session.findById("wnd[0]/usr/radDV_GLOBAL_DATA").setFocus()
            session.findById("wnd[0]/usr/btnSHOW").press()
            session.findById("wnd[0]/usr/subRIDER:/SAPAPO/SAPLMAT_MASTER:0150/tabsTABS/tabpUNIT").select()
            self.product_description = session.findById("wnd[0]/usr/subHEAD:/SAPAPO/SAPLMAT_MASTER:1201/txt/SAPAPO/MATIO-MAKTX").text
            general_id = "wnd[0]/usr/subRIDER:/SAPAPO/SAPLMAT_MASTER:0150/tabsTABS/tabpUNIT/ssubTABS_AREA_MAT:/SAPAPO/SAPLMAT_MASTER:2600/tbl/SAPAPO/SAPLMAT_MASTERTC_ME_2600/"
            
            # reads weight and height for unit of measure "LE" (layer of a pallet)
            row_index_LE = 0
            unit_of_measure = None
            while not(unit_of_measure == 'LE'):
                unit_of_measure = session.findById(f"{general_id}ctxt/SAPAPO/MATIO-MEINH[1,{row_index_LE}]").text
                row_index_LE += 1
            row_index_LE -= 1
            self.LE_weight = locale.atof(session.findById(f"{general_id}txt/SAPAPO/MATIO-BRGEW_TC[8,{row_index_LE}]").text)
            if (session.findById(f"{general_id}ctxt/SAPAPO/MATIO-GEWEI_TC[10,{row_index_LE}]").text) == 'G': #checks if unit of weight is in gramm
                self.LE_weight = self.LE_weight / 1000 #converts to KG
            self.LE_items = int(locale.atof(session.findById(f"{general_id}txt/SAPAPO/MATIO-UMREZ[3,{row_index_LE}]").text))
            self.LE_height = (locale.atof(session.findById(f"{general_id}txt/SAPAPO/MATIO-HOEHE[16,{row_index_LE}]").text)) / 10
            
            # reads weight and height for unit of measure "P2" (standard pallet)
            row_index_P2 = 0
            unit_of_measure = None
            while not(unit_of_measure == 'P2'):
                unit_of_measure = session.findById(f"{general_id}ctxt/SAPAPO/MATIO-MEINH[1,{row_index_P2}]").text
                row_index_P2 += 1
            row_index_P2 -= 1
            self.P2_weight = locale.atof(session.findById(f"{general_id}txt/SAPAPO/MATIO-BRGEW_TC[8,{row_index_P2}]").text)
            if (session.findById(f"{general_id}ctxt/SAPAPO/MATIO-GEWEI_TC[10,{row_index_P2}]").text) == 'G': #checks if unit of weight is in gramm
                self.P2_weight = self.P2_weight / 1000 #converts to KG
            self.P2_height = (locale.atof(session.findById(f"{general_id}txt/SAPAPO/MATIO-HOEHE[16,{row_index_P2}]").text)) / 10

      

        except:
            print(sys.exc_info()[0])

        finally:
            session = None
            connection = None
            application = None
            SapGuiAuto = None