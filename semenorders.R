library(readxl)
library(tidyverse)
library(readr)
library(writexl)
library(openxlsx)

orders <- read_excel("C:/Users/vance/Desktop/orders.xlsx", 
                     sheet = "Master")

orders1<-orders[c(3,5,7,9,10,11,12:15,18,19:33,35)]

hs<- createStyle(textDecoration = "BOLD")

aarmas<-orders1 %>% filter(Specialist=="ARMAS TORRES, ANDRES ANTONIO")
write.xlsx(aarmas,"ANDRES_ARMAS.xlsx", colWidths="auto", headerStyle = hs)

barmstrong<-orders1 %>% filter(Specialist=="ARMSTRONG, BARRY N.")
write.xlsx(barmstrong,"NEIL_ARMSTRONG.xlsx", colWidths="auto", headerStyle = hs)

jautry<-orders1 %>% filter(Specialist=="AUTRY, JEREMY W.")
write.xlsx(jautry,"JEREMY_AUTRY.xlsx", colWidths="auto", headerStyle = hs)

jbaldwin<-orders1 %>% filter(Specialist=="BALDWIN, JASON D.")
write.xlsx(jbaldwin,"JASON_BALDWIN.xlsx", colWidths="auto", headerStyle = hs)

sbereza<-orders1 %>% filter(Specialist=="BEREZA, SEBASTIAN R.")
write.xlsx(sbereza,"SEBASTIAN_BEREZA.xlsx", colWidths="auto", headerStyle = hs)

mbryant<-orders1 %>% filter(Specialist=="BRYANT, MARVIN W.")
write.xlsx(mbryant,"MARVIN_BRYANT.xlsx", colWidths="auto", headerStyle = hs)

mcarroll<-orders1 %>% filter(Specialist=="CARROLL, MICHAEL P.")
write.xlsx(mcarroll,"MICHAEL_CARROLL.xlsx", colWidths="auto", headerStyle = hs)

mcastelo<-orders1 %>% filter(Specialist=="CASTELO, MARIA Y.")
write.xlsx(mcastelo,"YOLANDA_CASTELO.xlsx", colWidths="auto", headerStyle = hs)

rcollins<-orders1 %>% filter(Specialist=="COLLINS, RICKY D.")
write.xlsx(rcollins,"RICKY_COLLINS.xlsx", colWidths="auto", headerStyle = hs)

rdonnelly<-orders1 %>% filter(Specialist=="DONNELLY JR, RICHARD")
write.xlsx(rdonnelly,"RICHARD_DONELLY.xlsx", colWidths="auto", headerStyle = hs)

jdraughon<-orders1 %>% filter(Specialist=="DRAUGHON, JASON C.")
write.xlsx(jdraughon,"JASON_DRAUGHON.xlsx", colWidths="auto", headerStyle = hs)

jgainey<-orders1 %>% filter(Specialist=="GAINEY, JAMES T.")
write.xlsx(jgainey,"JAMES_GAINEY.xlsx", colWidths="auto", headerStyle = hs)

rgautier<-orders1 %>% filter(Specialist=="GAUTIER, RONALD A.")
write.xlsx(rgautier,"RONALD_GAUTIER.xlsx", colWidths="auto", headerStyle = hs)

rholland<-orders1 %>% filter(Specialist=="HOLLAND, ROBERT G.")
write.xlsx(rholland,"ROBERT_HOLLAND.xlsx", colWidths="auto", headerStyle = hs)

cholt<-orders1 %>% filter(Specialist=="HOLT, CHRISTOPHER M.")
write.xlsx(cholt,"CHRISTOPHER_HOLT.xlsx", colWidths="auto", headerStyle = hs)

rjolly<-orders1 %>% filter(Specialist=="JOLLY, RANDY S.")
write.xlsx(rjolly,"RANDY_JOLLY.xlsx", colWidths="auto", headerStyle = hs)

sjozefowicz<-orders1 %>% filter(Specialist=="JOZEFOWICZ, SLAWOMIR")
write.xlsx(sjozefowicz,"SLAWOMIR_JOZEFOWICZ.xlsx", colWidths="auto", headerStyle = hs)

mmehlenbacher<-orders1 %>% filter(Specialist=="MEHLENBACHER, MATTHEW L.")
write.xlsx(mmehlenbacher,"MATTHEW_MEHLENBACHER.xlsx", colWidths="auto", headerStyle = hs)

dmurray<-orders1 %>% filter(Specialist=="MURRAY, DEBRA B.")
write.xlsx(dmurray,"DEBRA_MURRAY.xlsx", colWidths="auto", headerStyle = hs)

mneill<-orders1 %>% filter(Specialist=="NEILL, MICHAEL T.")
write.xlsx(mneill,"MICHAEL_NEILL.xlsx", colWidths="auto", headerStyle = hs)

jrouse<-orders1 %>% filter(Specialist=="ROUSE, JASON L.")
write.xlsx(jrouse,"JASON_ROUSE.xlsx", colWidths="auto", headerStyle = hs)

wscott<-orders1 %>% filter(Specialist=="SCOTT, WILLIAM R.")
write.xlsx(wscott,"RAY_SCOTT.xlsx", colWidths="auto", headerStyle = hs)

jtyer<-orders1 %>% filter(Specialist=="TYER, JOHN W.")
write.xlsx(jtyer,"JOHN_TYER.xlsx", colWidths="auto", headerStyle = hs)

junderwood<-orders1 %>% filter(Specialist=="UNDERWOOD, JOHN M.")
write.xlsx(junderwood,"MARK_UNDERWOOD.xlsx", colWidths="auto", headerStyle = hs)

wupchurch<-orders1 %>% filter(Specialist=="UPCHURCH, WILLIAM L.")
write.xlsx(wupchurch,"WILLIAM_UPCHURCH.xlsx", colWidths="auto", headerStyle = hs)

bwarren<-orders1 %>% filter(Specialist=="WARREN, BAILEY A.")
write.xlsx(bwarren,"BAILEY_WARREN.xlsx", colWidths="auto", headerStyle = hs)

ewatts<-orders1 %>% filter(Specialist=="WATTS, ERIC L.")
write.xlsx(ewatts,"ERIC_WATTS.xlsx", colWidths="auto", headerStyle = hs)

jworley<-orders1 %>% filter(Specialist=="WORLEY, JEFFERY D.")
write.xlsx(jworley,"JEFF_WORLEY.xlsx", colWidths="auto", headerStyle = hs)

rgarcia<-orders1 %>% filter(Specialist=="GARCIA RIVERA, RIGOBERTO DESIDERIO")
write.xlsx(rgarcia,"RIGOBERTO_GARCIA.xlsx", colWidths="auto", headerStyle = hs)

lparr<-orders1 %>% filter(Specialist=="PARR, LAURA A.")
write.xlsx(lparr,"LAURA_PARR.xlsx", colWidths="auto", headerStyle = hs)

lparr<-orders1 %>% filter(Specialist=="PARR, LAURA A.")
write.xlsx(lparr,"LAURA_PARR.xlsx", colWidths="auto", headerStyle = hs)

####################### By Stud ##########################

orders2<-orders1[c(1:5,12:24)]

aski<-orders2 %>% filter(`Current BS`==7092)
write.xlsx(aski,"7092_orders.xlsx", colWidths="auto", headerStyle = hs)

will<-orders2 %>% filter(`Current BS`==7093)
write.xlsx(will,"7093_orders.xlsx", colWidths="auto", headerStyle = hs)

robe<-orders2 %>% filter(`Current BS`==7094)
write.xlsx(robe,"7094_orders.xlsx", colWidths="auto", headerStyle = hs)

deer<-orders2 %>% filter(`Current BS`==7081)
write.xlsx(deer,"7081_orders.xlsx", colWidths="auto", headerStyle = hs)

laur<-orders2 %>% filter(`Current BS`==7082)
write.xlsx(laur,"7082_orders.xlsx", colWidths="auto", headerStyle = hs)

maco<-orders2 %>% filter(`Current BS`==9644)
write.xlsx(maco,"9644_orders.xlsx", colWidths="auto", headerStyle = hs)













 