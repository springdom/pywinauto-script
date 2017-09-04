from tkinter import Tk, Label, Button, Entry, StringVar, DISABLED, NORMAL, END, W, E

class MyFirstGUI:
    def __init__(self, master, server = 0, department = ""):
        self.server = server
        self.department = department

        self.master = master
        master.title("IA Auto")
        self.master.minsize(width=300,height=200)
        
        self.serv_label = Label(master, text="Select Server")
        self.serv_label.grid(columnspan = 2, sticky = W)
        
        self.cms_button = Button(master, text="CMS", command=lambda:self.get_server(1))
        self.cms_button.grid(row = 1)

        self.sf_button = Button(master, text="Salesforce", command=lambda:self.get_server(2))
        self.sf_button.grid(row = 1, column = 1)

        self.loc_label = Label(master, text="Select Location")
        self.loc_label.grid(row = 2, columnspan = 2, sticky = W)

        self.orl_button = Button(master, text="orl", command=lambda:self.get_location("orl"), state=DISABLED)
        self.orl_button.grid(row = 3)
        
        self.lvn_button = Button(master, text="lvn", command=lambda:self.get_location("lvn"), state=DISABLED)
        self.lvn_button.grid(row = 3, column = 1)
        
        self.spg_button = Button(master, text="spg", command=lambda:self.get_location("spg"), state=DISABLED)
        self.spg_button.grid(row = 3, column = 2)

        
        """
        self.close_button = Button(master, text="Close", command=master.quit)
        self.close_button.pack(side=RIGHT)
        """
        
    def get_server(self, button_id):
        if button_id == 1:
            self.server = 1
            print("CMS")
        elif button_id == 2:
            self.server = 2
            print("Salesforce")
            
        self.cms_button.configure(state=DISABLED)
        self.sf_button.configure(state=DISABLED)
        self.orl_button.configure(state=NORMAL)
        self.lvn_button.configure(state=NORMAL)
        self.spg_button.configure(state=NORMAL)
    
    def get_location(self, location):
        self.department = location
        print(location)
        self.orl_button.configure(state=DISABLED)
        self.lvn_button.configure(state=DISABLED)
        self.spg_button.configure(state=DISABLED)
        

root = Tk()
my_gui = MyFirstGUI(root)
root.mainloop()
