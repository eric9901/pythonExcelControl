import wx
class ExcelControl(wx.Frame):  #frame for excel control panel
    def __init__(self,*args,**kw):
        # ensure the parent's __init__ is called
        super(ExcelControl,self).__init__(*args,**kw)

        # create a panel in the frame
        pnl = wx.Panel(self)
        pnl.SetSizer(wx.BoxSizer())

        # and put some text with a larger bold font on it
        st = wx.StaticText(pnl, label="Xlas File Processer", pos=((25), 25))
        font = st.GetFont()
        font.PointSize += 10
        font = font.Bold()
        st.SetFont(font)
        st2 = wx.StaticLine(pnl, pos=(25,100))
        # create a menu bar
        self.makeMenuBar()
        #self.CreateToolBar()
        #self.SetToolBar()
# and a status bar
        self.CreateStatusBar()
        self.SetStatusText("Welcome to Excel Worker!")

    def makeMenuBar(self):
        """
        A menu bar is composed of menus, which are composed of menu items.
        This method builds a set of menus and binds handlers to be called
        when the menu item is selected.
        """

        # Make a file menu with Hello and Exit items
        fileMenu = wx.Menu()
        # The "\t..." syntax defines an accelerator key that also triggers
        # the same event
        excelItem = fileMenu.Append(-1, "&Excel...\tCtrl-H",
                                    "Help string shown in status bar for this menu item")
        fileMenu.AppendSeparator()
        # When using a stock ID we don't need to specify the menu item's
        # label
        exitItem = fileMenu.Append(wx.ID_EXIT)

        # Now a help menu for the about item
        helpMenu = wx.Menu()
        aboutItem = helpMenu.Append(wx.ID_ABOUT)

        # Make the menu bar and add the two menus to it. The '&' defines
        # that the next letter is the "mnemonic" for the menu item. On the
        # platforms that support it those letters are underlined and can be
        # triggered from the keyboard.
        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, "&File")
        menuBar.Append(helpMenu, "&Help")

        # Give the menu bar to the frame
        self.SetMenuBar(menuBar)

        # Finally, associate a handler function with the EVT_MENU event for
        # each of the menu items. That means that when that menu item is
        # activated then the associated handler function will be called.
        self.Bind(wx.EVT_MENU, self.OnHello, excelItem)
        self.Bind(wx.EVT_MENU, self.OnExit, exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

    def OnExit(self, event):
        """Close the frame, terminating the application."""
        self.Close(True)


    def OnHello(self, event):
        """Say hello to the user."""
        wx.MessageBox("Hello again from wxPython")


    def OnAbout(self, event):
        """Display an About Dialog"""
        wx.MessageBox("This is a program that for process the xlsx file by python language instead of SQL code.\n "+
                      "This program now only support read and write the xlas file",
                      "About this program.",
                      wx.OK|wx.ICON_INFORMATION)


#if __name__ == '__main__':
    # When this module is run (not imported) then create the app, the
    # frame, show it, and start the event loop.
app = wx.App()
frm = ExcelControl(None,title='excelProcessor/',pos=(60,60),size=(1000,600))
frm.Show()
app.MainLoop()