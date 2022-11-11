import smartsheet
from geopy.geocoders import Nominatim
import sys
import tkinter
from tkinter import *
import win32gui
import win32con


if sys.stdin.isatty():
    hide = win32gui.GetForegroundWindow()
    win32gui.ShowWindow(hide, win32con.SW_HIDE)


def splash():
    splash_win = Tk()
    splash_win.title('uCPE')
    splash_win.iconbitmap('terminal.ico')
    w = 240
    h = 195
    ws = splash_win.winfo_screenwidth()
    hs = splash_win.winfo_screenheight()
    x = ws / 2 - w / 2
    y = hs / 2 - h / 2
    splash_win.geometry('%dx%d+%d+%d' % (w, h, x, y))
    img = PhotoImage(file='ucpe.png')
    splashLogo = tkinter.Label(splash_win, image=img)
    splashLogo.pack()
    greeting = tkinter.Label(text='Author: Mikey Marcotte')
    greeting.pack()
    splash_win.after(3000, lambda: splash_win.destroy())
    splash_win.mainloop()


def main():
    def printValue(event):
        global tkt
        tkt = ticket.get()
        main.destroy()

    def printValue2():
        global tkt
        tkt = ticket.get()
        main.destroy()

    main = Tk()
    main.title('uCPE')
    main.iconbitmap('terminal.ico')
    w = 300
    h = 100
    ws = main.winfo_screenwidth()
    hs = main.winfo_screenheight()
    x = ws / 2 - w / 2
    y = hs / 2 - h / 2
    main.geometry('%dx%d+%d+%d' % (w, h, x, y))
    greeting = tkinter.Label(text='Please enter the ticket number: ')
    greeting.pack(pady=10)
    ticket = Entry(main)
    ticket.pack(pady=5)
    main.bind('<Return>', printValue)
    Button(main,
           text='Submit',
           padx=0,
           pady=0,
           command=printValue2).pack()
    main.mainloop()


def tryAgain():
    def printValue(event):
        global tkt
        tkt = ticket.get()
        ta.destroy()

    def printValue2():
        global tkt
        tkt = ticket.get()
        ta.destroy()

    ta = Tk()
    ta.title('uCPE')
    ta.iconbitmap('terminal.ico')
    w = 300
    h = 100
    ws = ta.winfo_screenwidth()
    hs = ta.winfo_screenheight()
    x = ws / 2 - w / 2
    y = hs / 2 - h / 2
    ta.geometry('%dx%d+%d+%d' % (w, h, x, y))
    greeting = tkinter.Label(text='Ticket not found. Please re-enter the ticket number: ')
    greeting.pack(pady=10)
    ticket = Entry(ta)
    ticket.pack(pady=5)
    ta.bind('<Return>', printValue)
    Button(ta,
           text='Submit',
           padx=0,
           pady=0,
           command=printValue2).pack()
    ta.bind('<Return>', printValue)
    ta.mainloop()


def popUp():
    global startOver

    def restart():
        win.destroy()

    def endProg():
        sys.exit()

    win = Tk()
    win.title('uCPE')
    win.iconbitmap('terminal.ico')
    w = 700
    h = 500
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = ws / 2 - w / 2
    y = hs / 2 - h / 2
    win.geometry('%dx%d+%d+%d' % (w, h, x, y))
    v = Scrollbar(win, orient='vertical')
    v.pack(side=RIGHT, fill='y')
    text = Text(win, font=('Consolas', '12'), yscrollcommand=v.set)
    text.insert(END, write1)
    v.config(command=text.yview)
    text.pack()
    Button(win,
           text='Start Over',
           width=10,
           height=1,
           padx=0,
           pady=0,
           command=restart).place(x=260, y=470)
    Button(win,
           text='Exit',
           width=10,
           height=1,
           padx=0,
           pady=0,
           command=endProg).place(x=360, y=470)
    win.mainloop()


geolocator = Nominatim(user_agent='geoapiExercises')
SMARTSHEET_ACCESS_TOKEN = 'vrwMoy8Wy9APWHRxhWvWkoo0k0GuYs4npJ2Kw'
smart = smartsheet.Smartsheet(SMARTSHEET_ACCESS_TOKEN)
smart.errors_as_exceptions(True)
sheet_id = 5140872800036740
sheet = smart.Sheets.get_sheet(sheet_id)
column_dict = {column.title: column.id for column in sheet.columns}


def getInfo():
    global exists, account, lanDhcp, lteSP, lteIP, custName, address, svc, cdtName, pipTempName, rtrType, tempName, vpnName, wanNNI, wanEVC, wanOuter, wanHandoff, wanNeg, custHandoff, pubSub, custVLAN, wanCPETag, wanMgmt, wanIP, wanPfx, wanGateway, lanPfx, lanIp
    exists = False
    for row in sheet.rows:
        if tkt == row.get_column(column_dict['Equipment Ticket/PWO']).display_value:
            exists = True
            account = row.get_column(column_dict['Account #']).display_value
            custName = row.get_column(column_dict['Customer Name']).display_value
            address = row.get_column(column_dict["Customer Location"]).display_value
            svc = row.get_column(column_dict['Service']).display_value
            cdtName = row.get_column(column_dict['Conductor Name']).display_value
            pipTempName = row.get_column(column_dict['PIP Template Name']).display_value
            rtrType = row.get_column(column_dict['Router Type']).display_value
            tempName = row.get_column(column_dict['uCPE Template Name']).display_value
            vpnName = row.get_column(column_dict['VPN Name']).display_value
            wanNNI = row.get_column(column_dict['WAN NNI CID']).display_value
            wanEVC = row.get_column(column_dict['WAN EVC ID  (UNI CID)']).display_value
            wanOuter = row.get_column(column_dict['WAN Outer VLAN']).display_value
            wanHandoff = row.get_column(column_dict['WAN Carrier Handoff']).display_value
            wanNeg = row.get_column(column_dict['WAN Carrier Negotiation']).display_value
            custHandoff = row.get_column(column_dict['Customer Handoff']).display_value
            pubSub = row.get_column(column_dict['Req Pub Subnet']).display_value
            custVLAN = row.get_column(column_dict['Customer VLAN']).display_value
            wanCPETag = row.get_column(column_dict['WAN CPE Tag']).display_value
            wanMgmt = row.get_column(column_dict['WAN Management IP Allocation']).display_value
            wanIP = row.get_column(column_dict['WAN IP Address']).display_value
            wanPfx = row.get_column(column_dict['WAN Prefix Length']).display_value
            wanGateway = row.get_column(column_dict['WAN gateway.']).display_value
            lanPfx = row.get_column(column_dict['LAN Prefix Length']).display_value
            lanIp = row.get_column(column_dict['LAN IP Address']).display_value
            lanDhcp = row.get_column(column_dict['LAN DHCP Server']).display_value
            lteSP = row.get_column(column_dict['LTE Service Provider']).display_value
            lteIP = row.get_column(column_dict['LTE IP Allocation']).display_value


splash()
startOver = True
while startOver is True:
    main()
    getInfo()
    while not exists:
        tryAgain()
        getInfo()

    write1 = ""
    dict1 = {"Ticket": tkt,
             "PIP Template Name": str(pipTempName)[:6],
             "uCPE Template Name": str(tempName)[:6],
             "Conductor Name": cdtName,
             "Router Type": rtrType,
             "Customer Name": custName,
             "Customer Account": account,
             "Customer Location": address,
             "LTE Service Provider": lteSP,
             "LTE IP Allocation": lteIP,
             "WAN NNI CID": wanNNI,
             "WAN EVC ID (UNI CID)": wanEVC,
             "WAN CPE Tag": wanCPETag,
             "WAN Outer VLAN": wanOuter,
             "WAN Carrier Handoff": wanHandoff,
             "WAN Carrier Negotiation": wanNeg,
             "WAN Management IP Allocation": wanMgmt,
             "WAN IP Address": wanIP,
             "WAN Prefix Length": wanPfx,
             "WAN Gateway": wanGateway,
             "Customer Handoff": custHandoff,
             "Customer VLAN": custVLAN,
             "Req Pub Subnet": pubSub,
             "LAN IP Address": lanIp,
             "LAN Prefix Length": lanPfx,
             "Lan DHCP Server": lanDhcp,
             "VPN Name": vpnName}

    for key, value in dict1.items():
        if str(value) != "None":
            write1 += key + ": " + str(value) + "\n"

    popUp()
