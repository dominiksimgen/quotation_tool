import sys
import win32com.client


class Logistical_Material_Data:
    def __init__(self,fpc):
        print("hello")



def read_APO_data(material):
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
        unit_of_measure = session.findById(f"{general_id}ctxt/SAPAPO/MATIO-MEINH[1,5]").text
        B = session.findById(f"{general_id}txt/SAPAPO/MATIO-BRGEW_TC[8,0]").text
        C = session.findById(f"{general_id}txt/SAPAPO/MATIO-EAN11[5,0]").text
        return unit_of_measure, B, C

    except:
        print(sys.exc_info()[0])

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


print(read_APO_data(81670114))