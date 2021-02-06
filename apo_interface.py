import sys
import win32com.client
import locale   #https://www.delftstack.com/de/howto/python/how-to-convert-string-to-float-or-int/


class Logistical_Material_Data:
    def __init__(self,fpc):
        locale.setlocale(locale.LC_ALL, 'nl_NL')
        self.read_APO_data(fpc)
        print(locale.atof(self.LE_height))
        print(locale.atof(self.LE_weight))



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
            general_id = "wnd[0]/usr/subRIDER:/SAPAPO/SAPLMAT_MASTER:0150/tabsTABS/tabpUNIT/ssubTABS_AREA_MAT:/SAPAPO/SAPLMAT_MASTER:2600/tbl/SAPAPO/SAPLMAT_MASTERTC_ME_2600/"
            
            row_index_LE = 0
            unit_of_measure = None
            while not(unit_of_measure == 'LE'):
                unit_of_measure = session.findById(f"{general_id}ctxt/SAPAPO/MATIO-MEINH[1,{row_index_LE}]").text
                row_index_LE += 1
            row_index_LE -= 1
            self.LE_weight = session.findById(f"{general_id}txt/SAPAPO/MATIO-BRGEW_TC[8,{row_index_LE}]").text
            self.LE_height = session.findById(f"{general_id}txt/SAPAPO/MATIO-HOEHE[16,{row_index_LE}]").text
            
            row_index_P2 = 0
            B = session.findById(f"{general_id}txt/SAPAPO/MATIO-BRGEW_TC[8,0]").text
            C = session.findById(f"{general_id}txt/SAPAPO/MATIO-EAN11[5,0]").text
            

        except:
            print(sys.exc_info()[0])

        finally:
            session = None
            connection = None
            application = None
            SapGuiAuto = None


new_material_data_1 = Logistical_Material_Data(81739020)
new_material_data_2 = Logistical_Material_Data(81731105)
#print(read_APO_data(81742965))