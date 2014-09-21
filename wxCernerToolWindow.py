__author__ = 'nbf1707'
#amazing layout tutorial: http://zetcode.com/wxpython/layout/

import BTsr
import ExcelFunctions
import wx
import os
import ConfigParser
import base64
import time
import SDPlus

version = '1.6'

srDetails = [
    {
        'srNumber' : '0011',
        'description' : 'Cycle Location Servers',
        'url' : 'https://bthealth.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?sysparm_id=0ec722b40a0aa0070106951a0db5bacc',
        'comment' : 'Please cycle the PMUtility server in CERT and PROD.',
        'attachmentRequired' : False
    },
    {
        'srNumber' : '0089',
        'description' : 'Alias RVJs (14231)',
        'url' : 'https://bthealth.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?sysparm_id=0ec723210a0aa00701e121d003765d45',
        'comment' : 'Please alias the following generic resources in CERT and PROD: ' + chr(13) + chr(13) + 'RVJ ? to Person ID ?',
        'attachmentRequired' : False
    },
    {
        'srNumber' : '0060',
        'description' : 'Alias Appointment Types (14230)',
        'url' : 'https://bthealth.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?sysparm_id=0ec715e80a0aa00701ddea5c15b7a0ce',
        'comment' : 'Please alias the following appointment types (codeset 14230) in CERT and PROD:',
        'attachmentRequired' : False
    },
    {
        'srNumber' : '0060',
        'description' : 'Add TCI Location',
        'url' : 'https://bthealth.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?sysparm_id=0ec715e80a0aa00701ddea5c15b7a0ce',
        'comment' : 'Please add TCI locations (see attached DCW, codeset 100300) in CERT and PROD.',
        'attachmentRequired' : True
    },
    {
        'srNumber' : '0128',
        'description' : 'Letters - Change flex rules',
        'url' : 'https://bthealth.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?sysparm_id=13c5ef5f0a0aa00700c70b414a607652',
        'comment' : 'Please amend the letter flex rules in CERT and PROD (see attached DCW)',
        'attachmentRequired' : True
    },
    {
        'srNumber' : '0123',
        'description' : 'Letters - upload RTFs',
        'url' : 'https://bthealth.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?sysparm_id=13c5eef60a0aa007000467ed78fafe88',
        'comment' : 'Please upload the attached RTFs in CERT and PROD.',
        'attachmentRequired' : True
    },
    {
        'srNumber' : '0134',
        'description' : 'Letters - upload telephone numbers',
        'url' : 'https://bthealth.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?sysparm_id=13c5eb8f0a0aa00701d5f47df422562a',
        'comment' : 'Please upload the attached telephone numbers in CERT and PROD.',
        'attachmentRequired' : True
    },
    {
        'srNumber' : '0124',
        'description' : 'Letters - change template',
        'url' : 'https://bthealth.service-now.com/com.glideapp.servicecatalog_cat_item_view.do?sysparm_id=13c5ef420a0aa0070074acc902e6c9c5',
        'comment' : 'Please change the letter templates (see attached dcw) in CERT and PROD.',
        'attachmentRequired' : True
    }
]

class MainWindow(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title="Cerner Tools")

        #items
        comboChoices = []
        for item in srDetails:
            comboChoices.append(item['description'])

        self.panel = wx.Panel(self)

        self.choiceCombo = wx.ComboBox(self.panel, choices=comboChoices)
        self.commentLabel = wx.StaticText(self.panel, label="Comment:")
        self.comment = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE, size=(200, 100))
        self.localRefLabel = wx.StaticText(self.panel, label="Local Ref:")
        self.localRef = wx.TextCtrl(self.panel)
        self.srNoLabel = wx.StaticText(self.panel)
        self.attachmentLabel = wx.StaticText(self.panel)
        self.attachmentButtonAdd = wx.Button(self.panel, label="Add Attachment")
        self.attachmentButtonClear = wx.Button(self.panel, label="Clear")
        self.logSRButton = wx.Button(self.panel, label="Log SR")
        self.staticLine = wx.StaticLine(self.panel)
        self.findRITMButton = wx.Button(self.panel, label="Get RITM number")
        self.ritmNo = wx.TextCtrl(self.panel)
        #self.confirmCheck = wx.CheckBox(self.panel, label="Submit Order")
        #self.ritmNo.SetBackgroundColour(self.panel.GetBackgroundColour())   #make same colour as parent
        self.addToBTButton = wx.Button(self.panel, label="Add to BT Sheet")

        #sizer for the Frame
        self.windowSizer = wx.BoxSizer()
        self.windowSizer.SetMinSize((350, 370))
        self.windowSizer.Add(self.panel, 1, wx.ALL | wx.EXPAND)

        #sizer for the Panel
        self.sizer = wx.GridBagSizer(hgap=10, vgap=5)
        self.sizer.Add(self.choiceCombo, pos=(0, 0), span=(1, 2), flag=wx.EXPAND)  #span 2 columns
        self.sizer.Add(self.commentLabel, (1, 0))
        self.sizer.Add(self.comment, (1, 1), flag=wx.EXPAND)
        self.sizer.Add(self.localRefLabel, (2, 0))
        self.sizer.Add(self.localRef, (2, 1), flag=wx.EXPAND)
        self.sizer.Add(self.srNoLabel, (3, 1))
        self.sizer.Add(self.attachmentLabel, (4, 0), (1, 2), flag=wx.EXPAND|wx.ALL, border=5)   #expand & 5px border all sides
        self.sizer.Add(self.attachmentButtonAdd, (5, 0))
        self.sizer.Add(self.attachmentButtonClear, (5, 1))
        self.sizer.Add(self.logSRButton, (6, 0))
        #self.sizer.Add(self.confirmCheck, (8, 1), flag=wx.EXPAND)  #removed at Baljit request
        self.sizer.Add(self.staticLine, (7, 0), (1, 2), flag=wx.EXPAND)
        self.sizer.Add(self.findRITMButton, (8, 0))
        self.sizer.Add(self.ritmNo, (8, 1), flag=wx.EXPAND)
        self.sizer.Add(self.addToBTButton, (9, 0))

        self.CreateStatusBar()

        self.sizer.AddGrowableCol(1)    #allows column 1 items to grow with the window

        # Set simple sizer for a nice border
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        # Use the sizers
        self.panel.SetSizerAndFit(self.border)
        self.SetSizerAndFit(self.windowSizer)

        #set up menu
        filemenu = wx.Menu()
        sdpAdd = wx.Menu()
        sdpUpdateRitm = wx.Menu()
        helpmenu = wx.Menu()
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu,"&File")
        menuBar.Append(sdpAdd,"&SDPlus")
        menuBar.Append(helpmenu,"&Help")
        self.SetMenuBar(menuBar)  # add the MenuBar to the Frame

        #set up menu sub menus
        # wx.ID_ABOUT and wx.ID_EXIT are standard ids provided by wxWidgets.
        ID_ADD_SDPLUS = wx.NewId()
        ID_UPDATE_SDPLUS = wx.NewId()
        ID_LOGINS = wx.NewId()
        menuLogins = filemenu.Append(ID_LOGINS, "Setup &Logins", "Set up logins for windows etc.")
        menuExit = filemenu.Append(wx.ID_EXIT,"E&xit","Terminate the program")
        menuSDPAdd = sdpAdd.Append(ID_ADD_SDPLUS, "&Add SR", "Add this service request to SDPlus")
        menuSDPUpdateRitm = sdpAdd.Append(ID_UPDATE_SDPLUS, "&Update RITM", "Update the RITM in SDPlus")
        menuAbout = helpmenu.Append(wx.ID_ABOUT, "&About","Information about this program")

        # Set events
        #http://wiki.wxpython.org/ListOfEvents
        self.choiceCombo.Bind(wx.EVT_COMBOBOX, self.onCombochoiceCombo)
        self.localRef.Bind(wx.EVT_TEXT, self.updateScreen)  #when text changes somehow in localRef, call updateScreen
        self.attachmentButtonAdd.Bind(wx.EVT_BUTTON, self.addAttachment)
        self.attachmentButtonClear.Bind(wx.EVT_BUTTON, self.clearAttachment)
        self.logSRButton.Bind(wx.EVT_BUTTON, self.BTLogin)
        self.findRITMButton.Bind(wx.EVT_BUTTON, self.GetRITM)
        self.addToBTButton.Bind(wx.EVT_BUTTON, self.appendToBTSheet)
        self.ritmNo.Bind(wx.EVT_TEXT, self.updateScreen)
        self.Bind(wx.EVT_MENU, self.loginsMenu, menuLogins)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        self.Bind(wx.EVT_MENU, self.sdPlusAdd, menuSDPAdd)
        self.Bind(wx.EVT_MENU, self.sdPlusUpdateSupplierRef, menuSDPUpdateRitm)
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_CLOSE, self.OnExit)

        #setup values
        self.choiceCombo.SetStringSelection(srDetails[0]['description'])
        self.comment.SetValue(srDetails[0]['comment'])
        #self.confirmCheck.Hide()
        self.updateScreen(wx.EVT_BUTTON)
        self.readConfig()
        self.Show()

    def updateScreen(self, e):
        self.findRITMButton.Disable()
        if os.path.isfile(self.attachmentLabel.GetLabel()): #is a valid file
            self.attachmentButtonClear.Enable()
            self.updateLogSRButton()
        else:
            self.attachmentButtonClear.Disable()
            if not srDetails[self.choiceCombo.GetSelection()]['attachmentRequired']:    #no attachment required
                self.attachmentLabel.SetLabel("(No attachment needed)")
                self.attachmentLabel.Disable()
                self.updateLogSRButton()
            else:
                self.logSRButton.Disable()
                self.attachmentLabel.SetLabel("Please choose an attachment:")
                self.attachmentLabel.Enable()
        #others
        if e == wx.EVT_COMBOBOX:
            self.comment.SetValue(srDetails[self.choiceCombo.GetSelection()]['comment'])
        self.srNoLabel.SetLabel('SR Number: ' + srDetails[self.choiceCombo.GetSelection()]['srNumber'])
        self.sizer.Layout() #needed to keep srNoLabel's right alignment after SetLabel()
        self.updateAddToBTButton()

    def updateLogSRButton(self):
        if len(self.localRef.GetValue()) >= 5:
            self.logSRButton.Enable()
        else:
            self.logSRButton.Disable()

    def updateAddToBTButton(self):
        if len(self.ritmNo.GetValue()) >= 9:
            self.addToBTButton.Enable()
        else:
            self.addToBTButton.Disable()

    def onCombochoiceCombo(self, e):
        self.updateScreen(wx.EVT_COMBOBOX)

    def addAttachment(self, e):
        dlg = wx.FileDialog(self, "Choose a file", '', "", "All Files (*.*)|*.*", wx.OPEN)    #wilcard needs "A|B" format else it crashes xp!!!
        if dlg.ShowModal() == wx.ID_OK:
            if os.path.isfile(dlg.GetPath()):
                self.attachmentLabel.SetLabel(dlg.GetPath())
                self.attachmentLabel.Enable()
        dlg.Destroy()
        self.updateScreen(wx.EVT_BUTTON)

    def clearAttachment(self, e):
        self.attachmentLabel.SetLabel("")
        self.updateScreen(wx.EVT_BUTTON)

    def OnExit(self,e):
        self.Destroy()

    def OnAbout(self,e):
        # A message dialog box with an OK button. wx.OK is a standard ID in wxWidgets.
        dlg = wx.MessageDialog(self, "Cerner Tools version " + version + " by Simon Crouch", "About Cerner Tools", wx.OK)
        dlg.ShowModal() # Show it
        dlg.Destroy() # finally destroy it when finished.

    def OnTimer(self, e):  #clears statusbar text after a time
        self.SetStatusText("")

    def SetStatusTextTimer(self, text="", timeout=3):
        self.SetStatusText(text)
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.OnTimer, self.timer)
        self.timer.Start(timeout*1000, oneShot=True)    #oneShot negates the need for self.timer.Stop()

    def BTLogin(self, e):
        self.SetStatusText("Logging in.")
        global browser   #make global for other functions
        browser, successFlag = BTsr.Login(emailAddress, btPassword)
        if not successFlag:
            self.SetStatusText("Problem with logging in. Please check username/password.")
        else:
            self.SetStatusText("Log in successful. Making service request.")
            url = srDetails[self.choiceCombo.GetSelection()]['url']
            localRef = self.localRef.GetValue()
            comm = self.comment.GetValue()
            #submit = self.confirmCheck.GetValue()
            submit = False
            attachPath = self.attachmentLabel.GetLabel()
            if not os.path.isfile(attachPath):
                attachPath = ""
            browser, successFlag = BTsr.logSR(browser, url, localRef, comm, submit, attachPath) #do the work
            if not successFlag:
                self.SetStatusText("Problem with logging service request.")
                self.findRITMButton.Disable()
            else:
                self.SetStatusText("Service Request logged.")
                self.findRITMButton.Enable()

    def GetRITM(self, e):   #note: requires global var 'browser' to exist first
        ritmNumber, successFlag = BTsr.getRITMNumber(browser)
        if successFlag:
            self.ritmNo.SetValue(ritmNumber)
        else:
            self.SetStatusText("Can't work out RITM number. Please type manually.")

    def appendToBTSheet(self, e):
        self.SetStatusText('Adding to BT Sheet ...')
        btFile = 'I:\Information Systems\Back Office Sharepoint\Cerner\calls to BT v1.xlsx'
        worksheetname = 'all calls to BT'
        xlDate = time.strftime("%m/%d/%Y")    #Excel seems to take only m/d/y in COM, tests ok.
        localRef = self.localRef.GetValue()
        summary = self.choiceCombo.GetValue()
        by = emailAddress[:emailAddress.find('.')].title()
        btRef = self.ritmNo.GetValue()
        nbtPriority = 4
        status = 'With BT'
        data = [xlDate, localRef, summary, by, btRef, nbtPriority, status]
        if ExcelFunctions.append(btFile, worksheetname, data):
            self.SetStatusText("Success writing to BT file.")
        else:
            self.SetStatusText("Error writing to BT file.")

    def loginsMenu(self, e):
        dlg = LoginsDialog(frame)
        dlg.ShowModal()

    def sdPlusAdd(self, e):
        self.SetStatusText('Adding call to SDPlus...')
        username = os.environ.get("USERNAME")
        SDPlus.setupLoginParams(username, windowsPassword)
        name = emailAddress[:emailAddress.find('.')].title() + ' ' + emailAddress[emailAddress.find('.')+1:emailAddress.find('@')].title()
        status, details = SDPlus.add(
            subject=self.choiceCombo.GetValue(),
            description='Service Request to ' + self.choiceCombo.GetValue()
                        + '\n(SR Number: ' + srDetails[self.choiceCombo.GetSelection()]['srNumber'] + ')'
                        + '\n\n' + self.comment.GetValue(),
            requester=name,
            requesteremail=emailAddress,
            technician=name,
            technicianemail=emailAddress,
            status='Hold - Awaiting Third Party'
            )
        if status:
            self.SetStatusText(details['message'])
            self.localRef.SetValue(details['workorderid'])
        else:
            self.SetStatusText('Fail writing to SDPlus - is windows password right?')

    def sdPlusUpdateSupplierRef(self, e):
        self.SetStatusText('Updating SR')
        username = os.environ.get("USERNAME")
        SDPlus.setupLoginParams(username, windowsPassword)
        workorderid = self.localRef.GetValue()
        ritm = self.ritmNo.GetValue()
        status, details = SDPlus.update(workorderid, supplierRef=ritm)
        if status:
            self.SetStatusText(details['message'])
        else:
            self.SetStatusText('Fail writing to SDPlus - is windows password right?')

    def readConfig(self):   #called at startup
        global emailAddress
        global btPassword
        global windowsPassword
        emailAddress = ''
        btPassword = ''
        windowsPassword = ''
        config = ConfigParser.RawConfigParser()
        try:
            config.read('settings.cfg')
            emailAddress = config.get('logindetails', 'emailaddress')
            btPassword = base64.b64decode(config.get('logindetails', 'btpassword'))
            windowsPassword = base64.b64decode(config.get('logindetails', 'windowspassword'))
            #self.confirmCheck.SetValue(config.getboolean('logindetails', 'confirmcheck'))   #getboolean is brilliant as get needs extra code for booleans!
            self.SetStatusTextTimer("Logins OK", 2)
        except:
            self.SetStatusText("Fail reading Logins. Please File, Setup Logins.")

class LoginsDialog(wx.Dialog):    #http://www.gcat.org.uk/tech/?p=56
    def __init__(self, parent, id=-1, title="Enter password"):
        wx.Dialog.__init__(self, parent, id, title) #size=(250, 130)
        self.panel = wx.Panel(self)

        self.emailAddressLabel = wx.StaticText(self.panel, label="Email Address:")
        self.emailAddress = wx.TextCtrl(self.panel, style=wx.TE_RICH|wx.TE_PROCESS_ENTER, size=(230, 20))
        self.btPasswordLabel = wx.StaticText(self.panel, label="BT SR Password:")
        self.btPassword = wx.TextCtrl(self.panel, style=wx.TE_PASSWORD|wx.TE_PROCESS_ENTER)
        self.windowsPasswordLabel = wx.StaticText(self.panel, label="Windows password:")
        self.windowsPassword = wx.TextCtrl(self.panel, value="", style=wx.TE_PASSWORD|wx.TE_PROCESS_ENTER)
        self.okButton = wx.Button(self.panel, label="OK", id=wx.ID_OK)
        self.cancelButton = wx.Button(self.panel, label="Cancel", id=wx.ID_CANCEL)

        #sizer for the Frame
        self.windowSizer = wx.BoxSizer()
        self.windowSizer.SetMinSize((250, 130))
        self.windowSizer.Add(self.panel, 1, wx.ALL | wx.EXPAND)

        #sizer for the Panel
        self.sizer = wx.GridBagSizer(hgap=10, vgap=5)
        self.sizer.Add(self.emailAddressLabel, (0, 0))
        self.sizer.Add(self.emailAddress, (0, 1), flag=wx.EXPAND)
        self.sizer.Add(self.btPasswordLabel, (1, 0))
        self.sizer.Add(self.btPassword, (1, 1), flag=wx.EXPAND)
        self.sizer.Add(self.windowsPasswordLabel, (2, 0))
        self.sizer.Add(self.windowsPassword, (2, 1), flag=wx.EXPAND)
        self.sizer.Add(self.okButton, (3, 0))
        self.sizer.Add(self.cancelButton, (3, 1))

        # Set simple sizer for a nice border
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        # Use the sizers
        self.panel.SetSizerAndFit(self.border)
        self.SetSizerAndFit(self.windowSizer)

        #events
        self.Bind(wx.EVT_BUTTON, self.onOK, id=wx.ID_OK)
        self.Bind(wx.EVT_BUTTON, self.onCancel, id=wx.ID_CANCEL)
        self.Bind(wx.EVT_TEXT_ENTER, self.onOK)

        #setup
        frame.readConfig()
        self.emailAddress.SetValue(emailAddress)
        self.btPassword.SetValue(btPassword)
        self.windowsPassword.SetValue(windowsPassword)

    def onOK(self, event):
        self.writeConfig(self)
        self.Destroy()

    def onCancel(self, event):
        self.Destroy()

    def writeConfig(self, e):
        global emailAddress
        global btPassword
        global windowsPassword
        config = ConfigParser.RawConfigParser()
        config.add_section('logindetails')
        emailAddress = self.emailAddress.GetValue()
        btPassword = self.btPassword.GetValue()
        windowsPassword = self.windowsPassword.GetValue()
        btPasswordEnc = base64.b64encode(self.btPassword.GetValue())
        windowsPasswordEnc = base64.b64encode(self.windowsPassword.GetValue())
        config.set('logindetails', 'emailaddress', emailAddress)
        config.set('logindetails', 'btpassword', btPasswordEnc)
        config.set('logindetails', 'windowspassword', windowsPasswordEnc)
        try:
            with open('settings.cfg', 'wb') as configfile:
                config.write(configfile)
        except:
            frame.SetStatusText("Couldn't write to configuration file...")
        else:
            frame.SetStatusTextTimer("Saved.", 2)

app = wx.App(False)
frame = MainWindow(None)
app.MainLoop()
