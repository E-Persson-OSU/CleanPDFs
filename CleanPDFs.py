#script to organize and name PDFs in PDF folder on my H drive and send me a report via email


#imports
from os               import listdir, rename, remove
from os.path          import isfile, join, getctime
from time             import localtime, strftime
from ntpath           import basename
from smtplib          import SMTP_SSL
from email.mime.text  import MIMEText
from configparser     import RawConfigParser

#config loading
config = RawConfigParser()
config.read('cleanpdf.ini')
login = config['login']
pw = config['password']
rcp = config['recipient']

#begin functions
def ispdf(file):
    return file.lower().endswith('.pdf')

def renameandfile(file,di):
    cts = getctime(file)                            # creation time in seconds
    ctl = localtime(cts)                            # creation time local as struct_time object
    ctf = strftime("%Y_%m_%d_%H_%M_%S", ctl)        # creation time formatted to YEAR_MONTH_DAY_HOUR_MINUTE_SECOND 
    ctf = ctf + '_' + str(di) + '.pdf'
    bn = basename(file)
    folder = 'Misc\\'
    if bn.lower().startswith('ps'):
        folder = 'Packing Slip\\'
    elif bn.lower().startswith('rs'):
        folder = 'Return Slips\\'
    elif bn.lower().startswith('qu'):
        folder = 'Quotes\\'
    #take action
    rename(file, pfp + folder + ctf)
    #log action
    return "Successfully renamed and moved " + file + " to " + pfp + folder + ctf + "\n"


#TODO make python file for sending emails
def mkreport(actionlist):
    fnw = 'H:\\Scripts\\PDFCleanupReports\\' + strftime('%d_%m_%Y') + '.txt'
    with open(fnw ,'w+') as fw:
        for a in actionlist:
            fw.write(a)
        fw.write("\nEnd report.")
    return fnw

def sendreport(filename):
    date = strftime("%a, %m %d" ,localtime())
    with open(filename) as re:
        msg = MIMEText(re.read())
    
    msg['Subject'] = "PDF Cleanup on " + date
    msg['From'] = login
    msg['To'] = rc

    s = SMTP_SSL('smtp.gmail.com',465,timeout=10)
    try:
        s.login(login,password)
        s.sendmail(msg['From'], msg['To'], msg.as_string())
    finally:
        s.quit()



#end functions

#main
pfp = 'H:\\pdfs\\' #pdf location
pdflist = listdir(pfp)
di = 1
al = []
for f in pdflist:
    fn = pfp + f
    if ispdf(fn):
        al.append(renameandfile(fn,di))
        di += 1
msgfile = mkreport(al)
sendreport(msgfile)