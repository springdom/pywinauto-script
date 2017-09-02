<<<<<<< HEAD
orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
orl_act = ["LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
orl_cc = ["LOC-ORL-MKT-CC", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
spg_ct = ["LOC-SPG-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
lvn_outbnd_cms = ["LAS_OUT_SUP","MKT-Outbound-Callback", "MKT-Outbound-Main2"]
lv_ct = ["CT Priority 1", "CT Priority 2", "LOC-LV-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
=======

>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf

wrkqueues = {}


<<<<<<< HEAD
=======
def Queues():
    orl_outbnd_cms = ["MKT-Outbound-Callback", "MKT-Outbound-Main2", "Orl_OUT_SUP"]
    orl_ct = ["CT Priority 1", "CT Priority 2", "LOC-ORL-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
    orl_act = ["LOC-ORL-MKT-ACT", "MKT-Activations-CallBack", "MKT-ACT-Main", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
    orl_cc = ["LOC-ORL-MKT-CC", "MKT-CC-BookDates", "MKT-CC-BookDates-Priority1", "MKT-CC-CustomerService", "MKT-CC-CustomerService-Priority2"]
    spg_ct = ["LOC-SPG-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
    lvn_outbnd_cms = ["LAS_OUT_SUP","MKT-Outbound-Callback", "MKT-Outbound-Main2"]
    lv_ct = ["CT Priority 1", "CT Priority 2", "LOC-LV-MKT-HRCC", "MKT-InbCT-Callback", "MKT-InbCT-HRCC"]
    for i in ('orl_outbnd_cms','orl_ct','orl_act','orl_cc','spg_ct','lvn_outbnd_cms','lv_ct'):
        wrkqueues[i] = locals()[i]

Queues()
>>>>>>> 8e45f202e44ca3f92e7df0b9aeb2e201bf981abf
