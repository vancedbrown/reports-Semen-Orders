library(sendmailR)
library(rmarkdown)

sendmail_options(smtpServer="smtp.murphybrownllc.com")

#------------------------------NORTH DIVISION--------------------------------------#

from <- "vbrown@smithfield.com"

to <- c("vbrown@smithfield.com",
        "mbryant@smithfield.com",
        "bupchurch@smithfield.com",
        "mneill@smithfield.com",
        "cholt@smithfield.com",
        "jbaldwin@smithfield.com",
        "ycastelo@smithfield.com",
        "rgholland@smithfield.com",
        "vmurrillo@smithfield.com")

subject <- "North Division Semen Orders"

msg <- c("See the attached semen orders by specialist, please use the submission form to submit permanent semen order changes. Send all order changes to boarstuds@spgenetics.com and update any missing contact information for your farms.",
         
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/MARVIN_BRYANT.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/WILLIAM_UPCHURCH.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/MICHAEL_NEILL.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/CHRISTOPHER_HOLT.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JASON_BALDWIN.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/YOLANDA_CASTELO.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ROBERT_HOLLAND.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ORDER_CHANGE_SUBMISSION_FORM.xlsx"))

sendmail(from,to,subject,msg)

#-----------------------------CENTRAL DIVISION----------------------------#

from <- "vbrown@smithfield.com"

to <- c("vbrown@smithfield.com",
        "aarmas@smithfield.com",
        "jautry@smithfield.com",
        "rdonnelly@smithfield.com",
        "jtgainey@smithfield.com",
        "rgautier@smithfield.com",
        "rjolly@smithfield.com",
        "mmehlenbacher@smithfield.com",
        "dmurray@smithfield.com",
        "jtyer@smithfield.com",
        "munderwood@smithfield.com",
        "bwarren@smithfield.com",
        "jworley@smithfield.com",
        "vmurrillo@smithfield.com")

subject <- "Central Division Semen Orders"

msg <- c("See the attached semen orders by specialist, please use the submission form to submit permanent semen order changes. Send all order changes to boarstuds@spgenetics.com and update any missing contact information for your farms.",
         
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ANDRES_ARMAS.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JEREMY_AUTRY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RICHARD_DONELLY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JAMES_GAINEY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RONALD_GAUTIER.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RANDY_JOLLY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/MATTHEW_MEHLENBACHER.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/DEBRA_MURRAY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JOHN_TYER.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/MARK_UNDERWOOD.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/BAILEY_WARREN.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JEFF_WORLEY.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ORDER_CHANGE_SUBMISSION_FORM.xlsx"))

sendmail(from,to,subject,msg)


#-------------------------------------------SOUTH DIVISION---------------------------------------#

from <- "vbrown@smithfield.com"

to <- c("vbrown@smithfield.com",
        "narmstrong@smithfield.com",
        "sbereza@smithfield.com",
        "mpcarroll@smithfield.com",
        "rcollins@smithfield.com",
        "jdraughon@smithfield.com",
        "sjozefowicz@smithfield.com",
        "jlrouse@smithfield.com",
        "rscott@smithfield.com",
        "ewatts@smithfield.com",
        "thagood@smithfield.com",
        "jpate@smithfield.com",
        "pzubrowicz@smithfield.com",
        "rdgarcia@smithfield.com",
        "vmurrillo@smithfield.com")

subject <- "South Division Semen Orders"

msg <- c("See the attached semen orders by specialist, please use the submission form to submit permanent semen order changes. Send all order changes to boarstuds@spgenetics.com and update any missing contact information for your farms.",
         
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/NEIL_ARMSTRONG.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/SEBASTIAN_BEREZA.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/MICHAEL_CARROLL.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RICKY_COLLINS.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JASON_DRAUGHON.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/SLAWOMIR_JOZEFOWICZ.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/JASON_ROUSE.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RAY_SCOTT.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/ERIC_WATTS.xlsx"),
         mime_part("C:/Users/vance/Documents/projects/Working Project Directory/reports/reports-Semen-Orders/RIGOBERTO_GARCIA.xlsx"),
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