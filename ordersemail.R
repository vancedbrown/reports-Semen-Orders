library(sendmailR)
library(rmarkdown)

sendmail_options(smtpServer="shpsmtp.murphybrownllc.com")

#------------------------------East Multiplication--------------------------------------#

from <- "vbrown@smithfield.com"

to <- c("vbrown@smithfield.com",
        "narmstrong@smithfield.com",
        "jbaldwin@smithfield.com",
        "rcollins@smithfield.com",
        "rdgarcia@smithfield.com",
        "jmauney@smithfield.com",
        "lparr@smithfield.com",
        "vmurrillo@smithfield.com")

subject <- "East Multiplication Division Semen Orders"

msg <- c("See the attached semen orders by specialist, please use the submission form to submit permanent semen order changes. 
         Send all order changes to boarstuds@spgenetics.com and update any missing contact information for your farms.",
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/NEIL_ARMSTRONG.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JASON_BALDWIN.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RICKY_COLLINS.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RIGOBERTO_GARCIA.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JIMMY_MAUNEY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/LAURA_PARR.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ORDER_CHANGE_SUBMISSION_FORM.xlsx"))

sendmail(from,to,subject,msg)

#-----------------------------East Commercial----------------------------#

from <- "vbrown@smithfield.com"

to <- c("vbrown@smithfield.com",
        "aarmas@smithfield.com",
        "jautry@smithfield.com",
        "sbereza@smithfield.com",
        "mbryant@smithfield.com",
        "jtgainey@smithfield.com",
        "rgholland@smithfield.com",
        "cholt@smithfield.com",
        "rjolly@smithfield.com",
        "sjozefowicz@smithfield.com",
        "mmehlenbacher@smithfield.com",
        "dmurray@smithfield.com",
        "mneill@smithfield.com",
        "rosorto@smithfield.com",
        "jlrouse@smithfield.com",
        "rscott@smithfield.com",
        "jtyer@smithfield.com",
        "bupchurch@smithfield.com",
        "bwarren@smithfield.com",
        "ewatts@smithfield.com",
        "jworley@smithfield.com",
        "vmurrillo@smithfield.com")

subject <- "East Commercial Division Semen Orders"

msg <- c("See the attached semen orders by specialist, please use the submission form to submit permanent semen order changes. 
         Send all order changes to boarstuds@spgenetics.com and update any missing contact information for your farms.",
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ANDRES_ARMAS.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JEREMY_AUTRY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/SEBASTIAN_BEREZA.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/MARVIN_BRYANT.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JAMES_GAINEY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ROBERT_HOLLAND.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/CHRISTOPHER_HOLT.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RANDY_JOLLY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/SLAWOMIR_JOZEFOWICZ.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/MATTHEW_MEHLENBACHER.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/DEBRA_MURRAY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/MICHAEL_NEILL.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ROGER_ANTONIO_OSORTO.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JASON_ROUSE.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RAY_SCOTT.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JOHN_TYER.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/WILLIAM_UPCHURCH.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/BAILEY_WARREN.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ERIC_WATTS.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JEFF_WORLEY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ORDER_CHANGE_SUBMISSION_FORM.xlsx"))

sendmail(from,to,subject,msg)


#-------------------------------------------BOAR STUDS---------------------------------------#

from <- "vbrown@smithfield.com"

to <- c("vbrown@smithfield.com",
        "vmurrillo@smithfield.com",
        "williamstonbs70930@smithfield.com",
        "robersonvillebs7094@smithfield.com",
        "askinsbs70920@smithfield.com",
        "DeercroftBS70810@smithfield.com",
        "laurelhillbs70820@smithfield.com")

subject <- "Semen Orders File"

msg <- c("See the attached semen orders by stud and please confirm with Victor that your orders match.",

         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/7092_orders.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/7093_orders.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/7094_orders.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/7081_orders.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/7082_orders.xlsx"))

sendmail(from,to,subject,msg)