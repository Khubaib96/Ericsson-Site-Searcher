from appJar import gui
import xlrd
from datetime import datetime
import clipboard

#####################################################################################

workbook = xlrd.open_workbook("warid.xlsx")
worksheet = workbook.sheet_by_index(0)
numbers = workbook.sheet_by_index(1)

######################################################################################
sites = []
bsc=[]
vendor = []
city = []
mbu = []
rbu = []
prime = []

#######################################################################################

mbu_number=[]
owner=[]
contact=[]
manager=[]
managercontact=[]
test_mbu = []
test_rbu = []
city_for_sms = []
total_sites = []
First_find = False
###################################################################################

for x in range(1, 5668):
    sites.append(worksheet.cell(x, 0).value)
    bsc.append(worksheet.cell(x, 1).value)
    vendor.append(worksheet.cell(x, 6).value)
    city.append(worksheet.cell(x, 7).value)
    mbu.append(worksheet.cell(x, 18).value)
    rbu.append(worksheet.cell(x, 19).value)
    prime.append(worksheet.cell(x, 23).value)

for x in range(1,98):
    mbu_number.append(numbers.cell(x, 2).value)
    owner.append(numbers.cell(x, 3).value)
    contact.append(numbers.cell(x, 5).value)
    manager.append(numbers.cell(x, 6).value)
    managercontact.append(numbers.cell(x, 7).value)

for y in range(0, 5667):
    temp = sites[y]
    sites[y] = temp[-6:]


# the title of the button will be received as a parameter
def press(btn):
    First_find = False
    app.clearTextArea("Vendor")
    app.clearTextArea("BSC")
    app.clearTextArea("City Name")
    app.clearTextArea("MBU")
    app.clearTextArea("RBU")
    app.clearTextArea("Prime site/Not prime site")
    app.clearTextArea("Mbu_owner")
    app.clearTextArea("Contact")
    app.clearTextArea("Zonal Manager Name")
    app.clearTextArea("Zonal Manager contact")
    site_id = app.getTextArea("Enter Site ID")
    site_id = site_id.upper()
    app.clearTextArea("Enter Site ID")
    if site_id.endswith('\n'):
        site_id = site_id[-7:]
        site_id = site_id[:-1]
    else:
        site_id = site_id [-6:]
    print (site_id)
    for z in range(0, 5667):
        if site_id == sites[z] and First_find == False:
            app.setTextArea("Enter Site ID", site_id)
            print (site_id)
            app.setTextArea("BSC", bsc[z])
            print (bsc[z])
            app.setTextArea("Vendor", vendor[z])
            print (vendor[z])
            app.setTextArea("City Name", city[z])
            print (city[z])
            app.setTextArea("MBU", mbu[z])
            print (mbu[z])
            app.setTextArea("RBU", rbu[z])
            print (rbu[z])
            app.setTextArea("Prime site/Not prime site", prime[z])
            print (prime[z])
            First_find = True

    test_mbu = app.getTextArea("MBU")

    for z in range(0, 97):
        if test_mbu == mbu_number[z]:
            app.setTextArea("Mbu_owner", owner[z])
            app.setTextArea("Contact", contact[z])
            app.setTextArea("Zonal Manager Name", manager[z])
            app.setTextArea("Zonal Manager contact", managercontact[z])

################################################################################

def SMS2G(btn):
    start_time = datetime.now()
    total_sites = str(app.getTextArea("Total Sites"))
    MBU = str(app.getTextArea("MBU"))
    City_NAME = str(app.getTextArea("City Name"))
    start_time = datetime.strftime(start_time, '%d/%m/%Y')
    start_time = str(start_time)
    RBU = str(app.getTextArea("RBU"))
    Vendor = str(app.getTextArea("Vendor"))


    output_2G = "["+"TT number: "+ "Start" + "]" + "\r\n" \
    + RBU + "\r\n" \
    + "Domain: RAN Ericsson"+ "\r\n" \
    + "FLM Vendor: " + Vendor+ "\r\n" \
    + "\r\n" \
    + "Event Description: Outage started on " + total_sites +" sites (2G) of MBU "+ MBU +" Serving " + City_NAME + " and Surroundings." + "\r\n" \
    + "Service Impact: Call, Data and SMS services affected for subscribers in coverage area" + "\r\n" \
    + "\r\n" \
    + "Reason: Place reason here"+ "\r\n" \
    + "\r\n" \
    + "Business Escalation: CC" + "\r\n" \
    + "\r\n" \
    + "Start Time: " + start_time + "\r\n" \
    + "Close Time: --" + "\r\n" \
    + "Action Taken: Escalated to NOSS ,FME , MBU lead and FO TRX"

    clipboard.copy(output_2G)


def SMS4G(btn):
    start_time = datetime.now()
    total_sites = str(app.getTextArea("Total Sites"))
    MBU = str(app.getTextArea("MBU"))
    City_NAME = str(app.getTextArea("City Name"))
    start_time = datetime.strftime(start_time, '%d/%m/%Y')
    start_time = str(start_time)
    RBU = str(app.getTextArea("RBU"))
    Vendor = str(app.getTextArea("Vendor"))

    output_4G = "[" + "TT number: " + "Start" + "]" + "\r\n" \
                + RBU + "\r\n" \
                + "Domain: RAN Ericsson" +  "\r\n" \
                + "FLM Vendor: " + Vendor + "\r\n" \
                + "\r\n" \
                + "Event Description: Outage started on " + total_sites + " sites (4G) of MBU " + MBU + " Serving " + City_NAME + " and Surroundings." + "\r\n" \
                + "Service Impact: Data services affected for subscribers in coverage area" + "\r\n" \
                + "\r\n" \
                + "Reason: Place reason here" + "\r\n" \
                + "\r\n" \
                + "Business Escalation: CC" + "\r\n" \
                + "\r\n" \
                + "Start Time: " + start_time + "\r\n" \
                + "Close Time: --" + "\r\n" \
                + "Action Taken: Escalated to NOSS ,FME , MBU lead and FO TRX"

    clipboard.copy(output_4G)

app = gui()

app.enableEnter(press)
app.addButton("Check", press,0,1)
app.addButton("Create 2G Multiple SMS", SMS2G,2,3)
app.addButton("Create 4G Multiple SMS", SMS4G,3,3)
app.enableEnter(press)


app.addLabel("l1", "Enter Site ID",2,0)
app.addTextArea("Enter Site ID",3,0)
app.setTextAreaHeight("Enter Site ID", 1)

app.addLabel("l2", "Vendor",4,0)
app.addTextArea("Vendor",5,0)
app.setTextAreaHeight("Vendor", 1)

app.addLabel("l12", "BSC",6,0)
app.addTextArea("BSC",7,0)
app.setTextAreaHeight("BSC", 1)

app.addLabel("l3", "City Name",8,0)
app.addTextArea("City Name",9,0)
app.setTextAreaHeight("City Name", 1)

app.addLabel("l4", "MBU",10,0)
app.addTextArea("MBU",11,0)
app.setTextAreaHeight("MBU", 1)

app.addLabel("l5", "RBU",12,0)
app.addTextArea("RBU",13,0)
app.setTextAreaHeight("RBU", 1)

app.addLabel("l6", "Prime site/Not prime site",14,0)
app.addTextArea("Prime site/Not prime site",15,0)
app.setTextAreaHeight("Prime site/Not prime site", 1)

app.addLabel("l8", "Mbu_owner",16,0)
app.addTextArea("Mbu_owner",17,0)
app.setTextAreaHeight("Mbu_owner", 1)

app.addLabel("l9", "Contact",18,0)
app.addTextArea("Contact",19,0)
app.setTextAreaHeight("Contact", 1)

app.addLabel("l10", "Zonal Manager Name",20,0)
app.addTextArea("Zonal Manager Name",21,0)
app.setTextAreaHeight("Zonal Manager Name", 1)

app.addLabel("l11", "Zonal Manager contact",22,0)
app.addTextArea("Zonal Manager contact",23,0)
app.setTextAreaHeight("Zonal Manager contact", 1)

app.addLabel("l13", "Total Sites",0,3)
app.addTextArea("Total Sites",1,3)
app.setTextAreaHeight("Total Sites", 1)

app.go()
