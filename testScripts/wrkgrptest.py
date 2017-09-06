"""
Automates Interaction Administrator
"""
#import time
import pywinauto
from pywinauto.application import Application


app = Application(backend='uia')
p = pywinauto.findwindows.find_element(title="Interaction Administrator - [HiltonACD]")
app.connect(handle=p.handle)
dlg = app.window(title="Interaction Administrator - [HiltonACD]")


#CMS
orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
lvn_outbnd_cms = ["LAS_OUT_SUP", "MKT-Outbound-Callback", "MKT-Outbound-Main2"]
spg_ct_cms = ["LOC-SPG-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]

orl_ct_cms = [
    "CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC",
    "MKT-InbCT-Callback", "MKT-InbCT-HRCC",
    ]
lvn_ct_cms = [
    "CT Priority 1", "CT Priority 2", "LOC-LV-MKT-HRCC",
    "MKT-InbCT-Callback", "MKT-InbCT-HRCC",
    ]
orl_act_cms = [
    "LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main",
    "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService",
    "MKT-CC-CustomerService-Priority2",
    ]
orl_cc_cms = [
    "LOC-ORL-MKT-CC", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1",
    "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2",
    ]

wrkqueues = {
    "orl_outbnd_cms":orl_outbnd_cms, "orl_ct_cms":orl_ct_cms, "orl_act_cms":orl_act_cms,
    "orl_cc_cms":orl_cc_cms, "spg_ct_cms":spg_ct_cms, "lvn_outbnd_cms":lvn_outbnd_cms,
    "lvn_ct_cms":lvn_ct_cms,
    }
"""
def agent_cms_workgroups(wrkgrps):
    Assign CMS WorkGroups
    num = 0
    app.dlg.Workgroups.click_input()
    app.dlg.OK.click_input()

    for x in wrkgrps:
        try:
            app.dlg[x].click_input()
            app.dlg.Add.click_input()
        except:
            listbox_pos(-2)
"""
def agent_cms_workgroups(wrkgrps):
    """Assign CMS WorkGroups"""
    num = 0
    
    while app.dlg["ListBox2"].item_count() != len(orl_outbnd_cms):
        try:
            app.dlg[orl_outbnd_cms[num]].click_input()
            app.dlg.Add.click_input()
            num+=1
        except:
            listbox_pos(-3)
        print(num)

def listbox_pos(scrollpos):
    """Gets Scrollbox Position"""
    app.dlg['ListBox'].click_input()
    app.dlg['ListBox'].wheel_mouse_input(wheel_dist=scrollpos)
