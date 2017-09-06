def agent_cms_workgroups(wrkgrps):
    """Assign CMS WorkGroups"""
    num = 0
    app.dlg.Workgroups.click_input()
    app.dlg.OK.click_input()

    for x in wrkgrps:
        try:
            app.dlg[x].click_input()
            app.dlg.Add.click_input()
        except:
            listbox_pos(-2)


def listbox_pos(scrollpos):
    """Gets Scrollbox Position"""
    app.dlg['ListBox'].click_input()
    app.dlg['ListBox'].wheel_mouse_input(wheel_dist=scrollpos)




def agent_cms_workgroups(wrkgrps):
    """Assign CMS WorkGroups"""
    num = 0
    app.dlg.Workgroups.click_input()
    app.dlg.OK.click_input()

    for x in wrkgrps:
        try:
            app.dlg[x].click_input()
            app.dlg.Add.click_input()
        except:
            listbox_pos(-2)

    while app.dlg['listBox2'].item_count() != len(wrkgrps):
        try:
            app.dlg[x].click_input()
            app.dlg.Add.click_input()
        except:
            listbox_pos(-2)
            
