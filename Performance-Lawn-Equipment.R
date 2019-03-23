#2i charts
#Dealer Satisfaction
DSNA = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 1,colIndex = c(1:8),rowIndex = c(5:9),header = FALSE)
DSSA = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 1,colIndex = c(1:8),rowIndex = c(12:16),header = FALSE)
DSEUR = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex =1,colIndex = c(1:8),rowIndex = c(19:23),header = FALSE)
DSPR = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 1,colIndex = c(1:8),rowIndex = c(26:30),header = FALSE)
DSCN = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 1,colIndex = c(1:8),rowIndex = c(33:35),header = FALSE)
navgscore2010 = (1*DSNA$X3[1]+2*DSNA$X4[1]+3*DSNA$X5[1]+4*DSNA$X6[1]+5*DSNA$X7[1])/DSNA$X8[1]
navgscore2011 = (1*DSNA$X3[2]+2*DSNA$X4[2]+3*DSNA$X5[2]+4*DSNA$X6[2]+5*DSNA$X7[2])/DSNA$X8[2]
navgscore2012 = (1*DSNA$X3[3]+2*DSNA$X4[3]+3*DSNA$X5[3]+4*DSNA$X6[3]+5*DSNA$X7[3])/DSNA$X8[3]
navgscore2013 = (1*DSNA$X3[4]+2*DSNA$X4[4]+3*DSNA$X5[4]+4*DSNA$X6[4]+5*DSNA$X7[4])/DSNA$X8[4]
navgscore2014 = (1*DSNA$X3[5]+2*DSNA$X4[5]+3*DSNA$X5[5]+4*DSNA$X6[5]+5*DSNA$X7[5])/DSNA$X8[5]
Northamerica = c(navgscore2010,navgscore2011,navgscore2012,navgscore2013,navgscore2014)
savgscore2010 = (1*DSSA$X3[1]+2*DSSA$X4[1]+3*DSSA$X5[1]+4*DSSA$X6[1]+5*DSSA$X7[1])/DSSA$X8[1]
savgscore2011 = (1*DSSA$X3[2]+2*DSSA$X4[2]+3*DSSA$X5[2]+4*DSSA$X6[2]+5*DSSA$X7[2])/DSSA$X8[2]
savgscore2012 = (1*DSSA$X3[3]+2*DSSA$X4[3]+3*DSSA$X5[3]+4*DSSA$X6[3]+5*DSSA$X7[3])/DSSA$X8[3]
savgscore2013 = (1*DSSA$X3[4]+2*DSSA$X4[4]+3*DSSA$X5[4]+4*DSSA$X6[4]+5*DSSA$X7[4])/DSSA$X8[4]
savgscore2014 = (1*DSSA$X3[5]+2*DSSA$X4[5]+3*DSSA$X5[5]+4*DSSA$X6[5]+5*DSSA$X7[5])/DSSA$X8[5]
Southamerica = c(savgscore2010,savgscore2011,savgscore2012,savgscore2013,savgscore2014)
eavgscore2010 = (1*DSEUR$X3[1]+2*DSEUR$X4[1]+3*DSEUR$X5[1]+4*DSEUR$X6[1]+5*DSEUR$X7[1])/DSEUR$X8[1]
eavgscore2011 = (1*DSEUR$X3[2]+2*DSEUR$X4[2]+3*DSEUR$X5[2]+4*DSEUR$X6[2]+5*DSEUR$X7[2])/DSEUR$X8[2]
eavgscore2012 = (1*DSEUR$X3[3]+2*DSEUR$X4[3]+3*DSEUR$X5[3]+4*DSEUR$X6[3]+5*DSEUR$X7[3])/DSEUR$X8[3]
eavgscore2013 = (1*DSEUR$X3[4]+2*DSEUR$X4[4]+3*DSEUR$X5[4]+4*DSEUR$X6[4]+5*DSEUR$X7[4])/DSEUR$X8[4]
eavgscore2014 = (1*DSEUR$X3[5]+2*DSEUR$X4[5]+3*DSEUR$X5[5]+4*DSEUR$X6[5]+5*DSEUR$X7[5])/DSEUR$X8[5]
Europe = c(eavgscore2010,eavgscore2011,eavgscore2012,eavgscore2013,eavgscore2014)
pavgscore2010 = (1*DSPR$X3[1]+2*DSPR$X4[1]+3*DSPR$X5[1]+4*DSPR$X6[1]+5*DSPR$X7[1])/DSPR$X8[1]
pavgscore2011 = (1*DSPR$X3[2]+2*DSPR$X4[2]+3*DSPR$X5[2]+4*DSPR$X6[2]+5*DSPR$X7[2])/DSPR$X8[2]
pavgscore2012 = (1*DSPR$X3[3]+2*DSPR$X4[3]+3*DSPR$X5[3]+4*DSPR$X6[3]+5*DSPR$X7[3])/DSPR$X8[3]
pavgscore2013 = (1*DSPR$X3[4]+2*DSPR$X4[4]+3*DSPR$X5[4]+4*DSPR$X6[4]+5*DSPR$X7[4])/DSPR$X8[4]
pavgscore2014 = (1*DSPR$X3[5]+2*DSPR$X4[5]+3*DSPR$X5[5]+4*DSPR$X6[5]+5*DSPR$X7[5])/DSPR$X8[5]
Pacific = c(pavgscore2010,pavgscore2011,pavgscore2012,pavgscore2013,pavgscore2014)
cavgscore2012 = (1*DSCN$X3[1]+2*DSCN$X3[1]+3*DSCN$X5[1]+4*DSCN$X6[1]+5*DSCN$X7[1])/DSCN$X8[1]
cavgscore2013 = (1*DSCN$X3[2]+2*DSCN$X3[2]+3*DSCN$X5[2]+4*DSCN$X6[2]+5*DSCN$X7[2])/DSCN$X8[2]
cavgscore2014 = (1*DSCN$X3[3]+2*DSCN$X3[3]+3*DSCN$X5[3]+4*DSCN$X6[3]+5*DSCN$X7[3])/DSCN$X8[3]
China = c(NA,NA,cavgscore2012,cavgscore2013,cavgscore2014)
plot(Northamerica,type = 'l', col = 'red', axes =FALSE, ylim=range(0,5,1), main = 'Dealer Satisfaction', xlab = 'Year', ylab = 'Survey results')
axis(1, at=1:5, labels = c('2010','2011','2012','2013','2014'))
axis(2, las = 1, at=1*0:range(20))
box()
lines(Southamerica,col = 'blue')
lines(Europe,col = 'green')
lines(Pacific, col = 'black')
lines(China, col = 'purple')
legend(1,2,cex=0.8,c('NA','SA','EUR','PR','CN'),col = c('red','blue','green','black','purple'),lty = 1)

#End-User Satisfaction
EUNA = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(4:8),header = FALSE)
EUSA = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(11:15),header = FALSE)
EUEUR = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(18:22),header = FALSE)
EUPR = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(25:29),header = FALSE)
EUCN = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(32:34),header = FALSE)
navgscore2010 = (1*EUNA$X3[1]+2*EUNA$X4[1]+3*EUNA$X5[1]+4*EUNA$X6[1]+5*EUNA$X7[1])/EUNA$X8[1]
navgscore2011 = (1*EUNA$X3[2]+2*EUNA$X4[2]+3*EUNA$X5[2]+4*EUNA$X6[2]+5*EUNA$X7[2])/EUNA$X8[2]
navgscore2012 = (1*EUNA$X3[3]+2*EUNA$X4[3]+3*EUNA$X5[3]+4*EUNA$X6[3]+5*EUNA$X7[3])/EUNA$X8[3]
navgscore2013 = (1*EUNA$X3[4]+2*EUNA$X4[4]+3*EUNA$X5[4]+4*EUNA$X6[4]+5*EUNA$X7[4])/EUNA$X8[4]
navgscore2014 = (1*EUNA$X3[5]+2*EUNA$X4[5]+3*EUNA$X5[5]+4*EUNA$X6[5]+5*EUNA$X7[5])/EUNA$X8[5]
Northamerica = c(navgscore2010,navgscore2011,navgscore2012,navgscore2013,navgscore2014)
savgscore2010 = (1*EUSA$X3[1]+2*EUSA$X4[1]+3*EUSA$X5[1]+4*EUSA$X6[1]+5*EUSA$X7[1])/EUSA$X8[1]
savgscore2011 = (1*EUSA$X3[2]+2*EUSA$X4[2]+3*EUSA$X5[2]+4*EUSA$X6[2]+5*EUSA$X7[2])/EUSA$X8[2]
savgscore2012 = (1*EUSA$X3[3]+2*EUSA$X4[3]+3*EUSA$X5[3]+4*EUSA$X6[3]+5*EUSA$X7[3])/EUSA$X8[3]
savgscore2013 = (1*EUSA$X3[4]+2*EUSA$X4[4]+3*EUSA$X5[4]+4*EUSA$X6[4]+5*EUSA$X7[4])/EUSA$X8[4]
savgscore2014 = (1*EUSA$X3[5]+2*EUSA$X4[5]+3*EUSA$X5[5]+4*EUSA$X6[5]+5*EUSA$X7[5])/EUSA$X8[5]
Southamerica = c(savgscore2010,savgscore2011,savgscore2012,savgscore2013,savgscore2014)
eavgscore2010 = (1*EUEUR$X3[1]+2*EUEUR$X4[1]+3*EUEUR$X5[1]+4*EUEUR$X6[1]+5*EUEUR$X7[1])/EUEUR$X8[1]
eavgscore2011 = (1*EUEUR$X3[2]+2*EUEUR$X4[2]+3*EUEUR$X5[2]+4*EUEUR$X6[2]+5*EUEUR$X7[2])/EUEUR$X8[2]
eavgscore2012 = (1*EUEUR$X3[3]+2*EUEUR$X4[3]+3*EUEUR$X5[3]+4*EUEUR$X6[3]+5*EUEUR$X7[3])/EUEUR$X8[3]
eavgscore2013 = (1*EUEUR$X3[4]+2*EUEUR$X4[4]+3*EUEUR$X5[4]+4*EUEUR$X6[4]+5*EUEUR$X7[4])/EUEUR$X8[4]
eavgscore2014 = (1*EUEUR$X3[5]+2*EUEUR$X4[5]+3*EUEUR$X5[5]+4*EUEUR$X6[5]+5*EUEUR$X7[5])/EUEUR$X8[5]
Europe = c(eavgscore2010,eavgscore2011,eavgscore2012,eavgscore2013,eavgscore2014)
pavgscore2010 = (1*EUPR$X3[1]+2*EUPR$X4[1]+3*EUPR$X5[1]+4*EUPR$X6[1]+5*EUPR$X7[1])/EUPR$X8[1]
pavgscore2011 = (1*EUPR$X3[2]+2*EUPR$X4[2]+3*EUPR$X5[2]+4*EUPR$X6[2]+5*EUPR$X7[2])/EUPR$X8[2]
pavgscore2012 = (1*EUPR$X3[3]+2*EUPR$X4[3]+3*EUPR$X5[3]+4*EUPR$X6[3]+5*EUPR$X7[3])/EUPR$X8[3]
pavgscore2013 = (1*EUPR$X3[4]+2*EUPR$X4[4]+3*EUPR$X5[4]+4*EUPR$X6[4]+5*EUPR$X7[4])/EUPR$X8[4]
pavgscore2014 = (1*EUPR$X3[5]+2*EUPR$X4[5]+3*EUPR$X5[5]+4*EUPR$X6[5]+5*EUPR$X7[5])/EUPR$X8[5]
Pacific = c(pavgscore2010,pavgscore2011,pavgscore2012,pavgscore2013,pavgscore2014)
cavgscore2012 = (1*EUCN$X3[1]+2*EUCN$X3[1]+3*EUCN$X5[1]+4*EUCN$X6[1]+5*EUCN$X7[1])/EUCN$X8[1]
cavgscore2013 = (1*EUCN$X3[2]+2*EUCN$X3[2]+3*EUCN$X5[2]+4*EUCN$X6[2]+5*EUCN$X7[2])/EUCN$X8[2]
cavgscore2014 = (1*EUCN$X3[3]+2*EUCN$X3[3]+3*EUCN$X5[3]+4*EUCN$X6[3]+5*EUCN$X7[3])/EUCN$X8[3]
China = c(NA,NA,cavgscore2012,cavgscore2013,cavgscore2014)
plot(Northamerica,type = 'l', col = 'red', axes =FALSE, ylim=range(0,5,1), main = 'End-User Satisfaction', xlab = 'Year', ylab = 'Survey results')
axis(1, at=1:5, labels = c('2010','2011','2012','2013','2014'))
axis(2, las = 1, at=1*0:range(20))
box()
lines(Southamerica,col = 'blue')
lines(Europe,col = 'green')
lines(Pacific, col = 'black')
lines(China, col = 'purple')
legend(1,2,cex=0.8,c('NA','SA','EUR','PR','CN'),col = c('red','blue','green','black','purple'),lty = 1)

#Complaints
complaints = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 4,startRow = 3, header = TRUE)
plot(complaints$World,type = 'l', col = 'red', axes =FALSE, ylim=range(0,350,50), main = 'Complaints', xlab = 'Year', ylab = 'Number of complaints')
axis(1,at =1,  labels = c('2010'))
axis(1,at =13,  labels = c('2011'))
axis(1,at =25,  labels = c('2012'))
axis(1,at =37,  labels = c('2013'))
axis(1,at =49,  labels = c('2014'))
axis(2, las = 1, at=50*0:350)
box()
lines(complaints$NA.,col = 'green')
lines(complaints$SA, col = 'black')
lines(complaints$Eur, col = 'purple')
lines(complaints$Pac, col = 'aquamarine')
lines(complaints$China, col = 'blue')
legend(1,350,cex=0.5,c('World','NA','SA','EUR','Pac','China'),col = c('red','green','black','purple','aquamarine','blue'),lty = 1)

#Mower Unit Sales
MUS = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 5,startRow = 3, header = TRUE)
plot(MUS$NA.,type = 'l', col = 'red', axes =FALSE,ylim = range(0,max(MUS[,c(2:7)]),2000), main = 'Mower Unit Sales', xlab = 'Year', ylab = 'Number of Sales')
axis(1,at =1,  labels = c('2010'))
axis(1,at =13,  labels = c('2011'))
axis(1,at =25,  labels = c('2012'))
axis(1,at =37,  labels = c('2013'))
axis(1,at =49,  labels = c('2014'))
axis(2, las = 1, at=2000*0:max(MUS[,c(2:7)]))
box()
lines(MUS$SA,col = 'green')
lines(MUS$Europe, col = 'black')
lines(MUS$Pacific, col = 'purple')
lines(MUS$China, col = 'aquamarine')
lines(MUS$World, col = 'blue')
legend(1,5000,cex=0.5,c('NA','SA','Europe','Pacific','China','World'),col = c('red','green','black','purple','aquamarine','blue'),lty = 1)

#Tractor Unit Sales
TUS = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 6,startRow = 3, header = TRUE)
plot(TUS$NA.,type = 'l', col = 'red', axes =FALSE,ylim = range(0,max(TUS[,c(2:7)]),500), main = 'Tractor Unit Sales', xlab = 'Year', ylab = 'Number of Sales')
axis(1,at =1,  labels = c('2010'))
axis(1,at =13,  labels = c('2011'))
axis(1,at =25,  labels = c('2012'))
axis(1,at =37,  labels = c('2013'))
axis(1,at =49,  labels = c('2014'))
axis(2, las = 1, at=500*0:max(TUS[,c(2:7)]))
box()
lines(TUS$SA,col = 'green')
lines(TUS$Eur, col = 'black')
lines(TUS$Pac, col = 'purple')
lines(TUS$China, col = 'aquamarine')
lines(TUS$World, col = 'blue')
legend(1,4000,cex=0.8,c('NA','SA','Eur','Pac','China','World'),col = c('red','green','black','purple','aquamarine','blue'),lty = 1)

#On-Time Delivery
OTD = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 11,startRow = 3, header = TRUE)
plot(OTD$Number.of.deliveries,type = 'l', col = 'red', axes =FALSE,ylim = range(min(OTD[,c(2:3)]),max(OTD[,c(2:3)])), main = 'On-Time Delivery', xlab = 'Year', ylab = 'Number of Delivery')
axis(1,at =1,  labels = c('2010'))
axis(1,at =13,  labels = c('2011'))
axis(1,at =25,  labels = c('2012'))
axis(1,at =37,  labels = c('2013'))
axis(1,at =49,  labels = c('2014'))
axis(2, las = 1, at=seq(min(OTD[,c(2:3)])+1,max(OTD[,c(2:3)]),50))
box()
lines(OTD$Number.On.Time,col = 'green')
legend(1,1420,cex=0.8,c('Number of deliveries','Numbers on time'),col = c('red','green'),lty = 1)

#Defects after Delivery
DAD = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 12,startRow = 4,rowIndex = c(4:16), header = TRUE)
plot(DAD$X2010,type = 'l', col = 'red', axes =FALSE,ylim = range(min(DAD[,c(2:6)]),max(DAD[,c(2:6)])), main = 'Defects after Delivery', xlab = 'Month', ylab = 'Number of Defects per million')
axis(2, las = 1, at=seq(min(DAD[,c(2:6)])+4,max(DAD[,c(2:6)]),50))
axis(1,at=1:12, labels = c('Jan','Feb','Mar','Apr','May','June','July','Aug','Sep','Oct','Nov','Dec'))
box()
lines(DAD$X2011,col = 'green')
lines(DAD$X2012, col = 'black')
lines(DAD$X2013, col = 'purple')
lines(DAD$X2014, col = 'aquamarine')
legend(1,540,cex=0.6,c('2010','2011','2012','2013','2014'),col = c('red','green','black','purple','aquamarine'),lty = 1)

#Response Time
RT = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 14,startRow = 3,rowIndex = c(3:53), header = TRUE)
values = c(mean(RT$Q1.2013),mean(RT$Q2.2013),mean(RT$Q3.2013),mean(RT$Q4.2013),mean(RT$Q1.2014),mean(RT$Q2.2014),mean(RT$Q3.2014),mean(RT$Q4.2014))
names(values) = c('Q1 2013','Q2 2013','Q3 2013','Q4 2013','Q1 2014','Q2 2014','Q3 2014','Q4 2014')
barplot(values,main = 'Response Time', xlab ='Quarters',ylab = 'Mean Response times',las = 2)


#Q2ii gross revenue market share
prices = data.frame(c(2010,2011,2012,2013,2014),c(150,175,180,185,190),c(3250,3400,3600,3700,3800))
names(prices) = c('Year','Mower price ($)','Tractor Price ($)')
#mower
MUS = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 5,startRow = 3, header = TRUE)
MUS[c(1:12),c(2:7)] = MUS[c(1:12),c(2:7)]*prices$`Mower price ($)`[1]
MUS[c(13:24),c(2:7)] = MUS[c(13:24),c(2:7)]*prices$`Mower price ($)`[2]
MUS[c(25:36),c(2:7)] = MUS[c(25:36),c(2:7)]*prices$`Mower price ($)`[3]
MUS[c(37:48),c(2:7)] = MUS[c(37:48),c(2:7)]*prices$`Mower price ($)`[4]
MUS[c(49:60),c(2:7)] = MUS[c(49:60),c(2:7)]*prices$`Mower price ($)`[5]
#tractor
trac = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 6,startRow = 3, header = TRUE)
trac[c(1:12),c(2:7)] = trac[c(1:12),c(2:7)]*prices$`Tractor Price ($)`[1]
trac[c(13:24),c(2:7)] = trac[c(13:24),c(2:7)]*prices$`Tractor Price ($)`[2]
trac[c(25:36),c(2:7)] = trac[c(25:36),c(2:7)]*prices$`Tractor Price ($)`[3]
trac[c(37:48),c(2:7)] = trac[c(37:48),c(2:7)]*prices$`Tractor Price ($)`[4]
trac[c(49:60),c(2:7)] = trac[c(49:60),c(2:7)]*prices$`Tractor Price ($)`[5]
#industry mower
cleanMUS = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 5,startRow = 3, colIndex = c(2:5), header = TRUE)
IMUS=read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 7,startRow = 3,colIndex = c(1:5), header = TRUE)
mower.mrkt.share = cbind(IMUS[1],round((cleanMUS/IMUS[-1])*100,2))
round(colSums(mower.mrkt.share[-1])/60,2)
#industry tractor
cleantrac = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 6,startRow = 3,colIndex = c(1:6), header = TRUE)
Itrac=read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 8,startRow = 3,colIndex = c(1:6), header = TRUE)
trac.mrkt.share = cbind(trac[1],round((cleantrac[-1]/Itrac[-1])*100,2))
round(colSums(trac.mrkt.share[-1])/60,2)


#Q3i mean and sd
#Dealer Satisfaction
DSNA = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 1,colIndex = c(1:8),rowIndex = c(5:9),header = FALSE)
DSSA = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 1,colIndex = c(1:8),rowIndex = c(12:16),header = FALSE)
DSEUR = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex =1,colIndex = c(1:8),rowIndex = c(19:23),header = FALSE)
DSPR = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 1,colIndex = c(1:8),rowIndex = c(26:30),header = FALSE)
DSCN = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 1,colIndex = c(1:8),rowIndex = c(33:35),header = FALSE)
#North america
NA2010 = c(rep(0,DSNA[1,2]),rep(1,DSNA[1,3]),rep(2,DSNA[1,4]),rep(3,DSNA[1,5]),rep(4,DSNA[1,6]),rep(5,DSNA[1,7]))
NA2011 = c(rep(0,DSNA[2,2]),rep(1,DSNA[2,3]),rep(2,DSNA[2,4]),rep(3,DSNA[2,5]),rep(4,DSNA[2,6]),rep(5,DSNA[2,7]))
NA2012 = c(rep(0,DSNA[3,2]),rep(1,DSNA[3,3]),rep(2,DSNA[3,4]),rep(3,DSNA[3,5]),rep(4,DSNA[3,6]),rep(5,DSNA[3,7]))
NA2013 = c(rep(0,DSNA[4,2]),rep(1,DSNA[4,3]),rep(2,DSNA[4,4]),rep(3,DSNA[4,5]),rep(4,DSNA[4,6]),rep(5,DSNA[4,7]))
NA2014 = c(rep(0,DSNA[5,2]),rep(1,DSNA[5,3]),rep(2,DSNA[5,4]),rep(3,DSNA[5,5]),rep(4,DSNA[5,6]),rep(5,DSNA[5,7]))
NAmean = c(mean(NA2010),mean(NA2011),mean(NA2012),mean(NA2013),mean(NA2014))
NASD = c(sd(NA2010),sd(NA2011),sd(NA2012),sd(NA2013),sd(NA2014))
#South america
SA2010 = c(rep(0,DSSA[1,2]),rep(1,DSSA[1,3]),rep(2,DSSA[1,4]),rep(3,DSSA[1,5]),rep(4,DSSA[1,6]),rep(5,DSSA[1,7]))
SA2011 = c(rep(0,DSSA[2,2]),rep(1,DSSA[2,3]),rep(2,DSSA[2,4]),rep(3,DSSA[2,5]),rep(4,DSSA[2,6]),rep(5,DSSA[2,7]))
SA2012 = c(rep(0,DSSA[3,2]),rep(1,DSSA[3,3]),rep(2,DSSA[3,4]),rep(3,DSSA[3,5]),rep(4,DSSA[3,6]),rep(5,DSSA[3,7]))
SA2013 = c(rep(0,DSSA[4,2]),rep(1,DSSA[4,3]),rep(2,DSSA[4,4]),rep(3,DSSA[4,5]),rep(4,DSSA[4,6]),rep(5,DSSA[4,7]))
SA2014 = c(rep(0,DSSA[5,2]),rep(1,DSSA[5,3]),rep(2,DSSA[5,4]),rep(3,DSSA[5,5]),rep(4,DSSA[5,6]),rep(5,DSSA[5,7]))
SAmean = c(mean(SA2010),mean(SA2011),mean(SA2012),mean(SA2013),mean(SA2014))
SASD = c(sd(SA2010),sd(SA2011),sd(SA2012),sd(SA2013),sd(SA2014))
#Europe
EUR2010 = c(rep(0,DSEUR[1,2]),rep(1,DSEUR[1,3]),rep(2,DSEUR[1,4]),rep(3,DSEUR[1,5]),rep(4,DSEUR[1,6]),rep(5,DSEUR[1,7]))
EUR2011 = c(rep(0,DSEUR[2,2]),rep(1,DSEUR[2,3]),rep(2,DSEUR[2,4]),rep(3,DSEUR[2,5]),rep(4,DSEUR[2,6]),rep(5,DSEUR[2,7]))
EUR2012 = c(rep(0,DSEUR[3,2]),rep(1,DSEUR[3,3]),rep(2,DSEUR[3,4]),rep(3,DSEUR[3,5]),rep(4,DSEUR[3,6]),rep(5,DSEUR[3,7]))
EUR2013 = c(rep(0,DSEUR[4,2]),rep(1,DSEUR[4,3]),rep(2,DSEUR[4,4]),rep(3,DSEUR[4,5]),rep(4,DSEUR[4,6]),rep(5,DSEUR[4,7]))
EUR2014 = c(rep(0,DSEUR[5,2]),rep(1,DSEUR[5,3]),rep(2,DSEUR[5,4]),rep(3,DSEUR[5,5]),rep(4,DSEUR[5,6]),rep(5,DSEUR[5,7]))
EURmean = c(mean(EUR2010),mean(EUR2011),mean(EUR2012),mean(EUR2013),mean(EUR2014))
EURSD = c(sd(EUR2010),sd(EUR2011),sd(EUR2012),sd(EUR2013),sd(EUR2014))
#Pacific
PR2010 = c(rep(0,DSPR[1,2]),rep(1,DSPR[1,3]),rep(2,DSPR[1,4]),rep(3,DSPR[1,5]),rep(4,DSPR[1,6]),rep(5,DSPR[1,7]))
PR2011 = c(rep(0,DSPR[2,2]),rep(1,DSPR[2,3]),rep(2,DSPR[2,4]),rep(3,DSPR[2,5]),rep(4,DSPR[2,6]),rep(5,DSPR[2,7]))
PR2012 = c(rep(0,DSPR[3,2]),rep(1,DSPR[3,3]),rep(2,DSPR[3,4]),rep(3,DSPR[3,5]),rep(4,DSPR[3,6]),rep(5,DSPR[3,7]))
PR2013 = c(rep(0,DSPR[4,2]),rep(1,DSPR[4,3]),rep(2,DSPR[4,4]),rep(3,DSPR[4,5]),rep(4,DSPR[4,6]),rep(5,DSPR[4,7]))
PR2014 = c(rep(0,DSPR[5,2]),rep(1,DSPR[5,3]),rep(2,DSPR[5,4]),rep(3,DSPR[5,5]),rep(4,DSPR[5,6]),rep(5,DSPR[5,7]))
PRmean = c(mean(PR2010),mean(PR2011),mean(PR2012),mean(PR2013),mean(PR2014))
PRSD = c(sd(PR2010),sd(PR2011),sd(PR2012),sd(PR2013),sd(PR2014))
#China
CN2012 = c(rep(0,DSCN[1,2]),rep(1,DSCN[1,3]),rep(2,DSCN[1,4]),rep(3,DSCN[1,5]),rep(4,DSCN[1,6]),rep(5,DSCN[1,7]))
CN2013 = c(rep(0,DSCN[2,2]),rep(1,DSCN[2,3]),rep(2,DSCN[2,4]),rep(3,DSCN[2,5]),rep(4,DSCN[2,6]),rep(5,DSCN[2,7]))
CN2014 = c(rep(0,DSCN[3,2]),rep(1,DSCN[3,3]),rep(2,DSCN[3,4]),rep(3,DSCN[3,5]),rep(4,DSCN[3,6]),rep(5,DSCN[3,7]))
CNmean = c(NA,NA,mean(CN2012),mean(CN2013),mean(CN2014))
CNSD = c(NA,NA,sd(CN2012),sd(CN2013),sd(CN2014))
#table for mean
meanDS = data.frame(rbind(NAmean,SAmean,EURmean,PRmean,CNmean))
names(meanDS) = c('2010','2011','2012','2013','2014')
meanDS = round(meanDS,2)
#table for standard deviation
sdDS = data.frame(rbind(NASD,SASD,EURSD,PRSD,CNSD))
names(sdDS) = c('2010','2011','2012','2013','2014')
sdDS = round(sdDS,2)

#End-User Satisfaction
EUNA = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(4:8),header = FALSE)
EUSA = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(11:15),header = FALSE)
EUEUR = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(18:22),header = FALSE)
EUPR = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(25:29),header = FALSE)
EUCN = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 2,colIndex = c(1:8),rowIndex = c(32:34),header = FALSE)
#North america
NA2010 = c(rep(0,EUNA[1,2]),rep(1,EUNA[1,3]),rep(2,EUNA[1,4]),rep(3,EUNA[1,5]),rep(4,EUNA[1,6]),rep(5,EUNA[1,7]))
NA2011 = c(rep(0,EUNA[2,2]),rep(1,EUNA[2,3]),rep(2,EUNA[2,4]),rep(3,EUNA[2,5]),rep(4,EUNA[2,6]),rep(5,EUNA[2,7]))
NA2012 = c(rep(0,EUNA[3,2]),rep(1,EUNA[3,3]),rep(2,EUNA[3,4]),rep(3,EUNA[3,5]),rep(4,EUNA[3,6]),rep(5,EUNA[3,7]))
NA2013 = c(rep(0,EUNA[4,2]),rep(1,EUNA[4,3]),rep(2,EUNA[4,4]),rep(3,EUNA[4,5]),rep(4,EUNA[4,6]),rep(5,EUNA[4,7]))
NA2014 = c(rep(0,EUNA[5,2]),rep(1,EUNA[5,3]),rep(2,EUNA[5,4]),rep(3,EUNA[5,5]),rep(4,EUNA[5,6]),rep(5,EUNA[5,7]))
NAmean = c(mean(NA2010),mean(NA2011),mean(NA2012),mean(NA2013),mean(NA2014))
NASD = c(sd(NA2010),sd(NA2011),sd(NA2012),sd(NA2013),sd(NA2014))
#South america
SA2010 = c(rep(0,EUSA[1,2]),rep(1,EUSA[1,3]),rep(2,EUSA[1,4]),rep(3,EUSA[1,5]),rep(4,EUSA[1,6]),rep(5,EUSA[1,7]))
SA2011 = c(rep(0,EUSA[2,2]),rep(1,EUSA[2,3]),rep(2,EUSA[2,4]),rep(3,EUSA[2,5]),rep(4,EUSA[2,6]),rep(5,EUSA[2,7]))
SA2012 = c(rep(0,EUSA[3,2]),rep(1,EUSA[3,3]),rep(2,EUSA[3,4]),rep(3,EUSA[3,5]),rep(4,EUSA[3,6]),rep(5,EUSA[3,7]))
SA2013 = c(rep(0,EUSA[4,2]),rep(1,EUSA[4,3]),rep(2,EUSA[4,4]),rep(3,EUSA[4,5]),rep(4,EUSA[4,6]),rep(5,EUSA[4,7]))
SA2014 = c(rep(0,EUSA[5,2]),rep(1,EUSA[5,3]),rep(2,EUSA[5,4]),rep(3,EUSA[5,5]),rep(4,EUSA[5,6]),rep(5,EUSA[5,7]))
SAmean = c(mean(SA2010),mean(SA2011),mean(SA2012),mean(SA2013),mean(SA2014))
SASD = c(sd(SA2010),sd(SA2011),sd(SA2012),sd(SA2013),sd(SA2014))
#Europe
EUR2010 = c(rep(0,EUEUR[1,2]),rep(1,EUEUR[1,3]),rep(2,EUEUR[1,4]),rep(3,EUEUR[1,5]),rep(4,EUEUR[1,6]),rep(5,EUEUR[1,7]))
EUR2011 = c(rep(0,EUEUR[2,2]),rep(1,EUEUR[2,3]),rep(2,EUEUR[2,4]),rep(3,EUEUR[2,5]),rep(4,EUEUR[2,6]),rep(5,EUEUR[2,7]))
EUR2012 = c(rep(0,EUEUR[3,2]),rep(1,EUEUR[3,3]),rep(2,EUEUR[3,4]),rep(3,EUEUR[3,5]),rep(4,EUEUR[3,6]),rep(5,EUEUR[3,7]))
EUR2013 = c(rep(0,EUEUR[4,2]),rep(1,EUEUR[4,3]),rep(2,EUEUR[4,4]),rep(3,EUEUR[4,5]),rep(4,EUEUR[4,6]),rep(5,EUEUR[4,7]))
EUR2014 = c(rep(0,EUEUR[5,2]),rep(1,EUEUR[5,3]),rep(2,EUEUR[5,4]),rep(3,EUEUR[5,5]),rep(4,EUEUR[5,6]),rep(5,EUEUR[5,7]))
EURmean = c(mean(EUR2010),mean(EUR2011),mean(EUR2012),mean(EUR2013),mean(EUR2014))
EURSD = c(sd(EUR2010),sd(EUR2011),sd(EUR2012),sd(EUR2013),sd(EUR2014))
#Pacific
PR2010 = c(rep(0,EUPR[1,2]),rep(1,EUPR[1,3]),rep(2,EUPR[1,4]),rep(3,EUPR[1,5]),rep(4,EUPR[1,6]),rep(5,EUPR[1,7]))
PR2011 = c(rep(0,EUPR[2,2]),rep(1,EUPR[2,3]),rep(2,EUPR[2,4]),rep(3,EUPR[2,5]),rep(4,EUPR[2,6]),rep(5,EUPR[2,7]))
PR2012 = c(rep(0,EUPR[3,2]),rep(1,EUPR[3,3]),rep(2,EUPR[3,4]),rep(3,EUPR[3,5]),rep(4,EUPR[3,6]),rep(5,EUPR[3,7]))
PR2013 = c(rep(0,EUPR[4,2]),rep(1,EUPR[4,3]),rep(2,EUPR[4,4]),rep(3,EUPR[4,5]),rep(4,EUPR[4,6]),rep(5,EUPR[4,7]))
PR2014 = c(rep(0,EUPR[5,2]),rep(1,EUPR[5,3]),rep(2,EUPR[5,4]),rep(3,EUPR[5,5]),rep(4,EUPR[5,6]),rep(5,EUPR[5,7]))
PRmean = c(mean(PR2010),mean(PR2011),mean(PR2012),mean(PR2013),mean(PR2014))
PRSD = c(sd(PR2010),sd(PR2011),sd(PR2012),sd(PR2013),sd(PR2014))
#China
CN2012 = c(rep(0,EUCN[1,2]),rep(1,EUCN[1,3]),rep(2,EUCN[1,4]),rep(3,EUCN[1,5]),rep(4,EUCN[1,6]),rep(5,EUCN[1,7]))
CN2013 = c(rep(0,EUCN[2,2]),rep(1,EUCN[2,3]),rep(2,EUCN[2,4]),rep(3,EUCN[2,5]),rep(4,EUCN[2,6]),rep(5,EUCN[2,7]))
CN2014 = c(rep(0,EUCN[3,2]),rep(1,EUCN[3,3]),rep(2,EUCN[3,4]),rep(3,EUCN[3,5]),rep(4,EUCN[3,6]),rep(5,EUCN[3,7]))
CNmean = c(NA,NA,mean(CN2012),mean(CN2013),mean(CN2014))
CNSD = c(NA,NA,sd(CN2012),sd(CN2013),sd(CN2014))
#table for mean
meanEU = data.frame(rbind(NAmean,SAmean,EURmean,PRmean,CNmean))
names(meanEU) = c('2010','2011','2012','2013','2014')
meanEU = round(meanEU,2)
#table for standard deviation
sdEU = data.frame(rbind(NASD,SASD,EURSD,PRSD,CNSD))
names(sdEU) = c('2010','2011','2012','2013','2014')
sdEU = round(sdEU,2)


#3ii Detailed summary
CS2014 = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 3,startRow = 3,header = TRUE)
CS2014NA = subset(CS2014,CS2014$Region == 'NA')
CS2014SA = subset(CS2014,CS2014$Region == 'SA')
CS2014Eur = subset(CS2014,CS2014$Region == 'Eur')
CS2014Pac = subset(CS2014,CS2014$Region == 'Pac')
CS2014China = subset(CS2014,CS2014$Region == 'China')
summary(CS2014NA)
summary(CS2014SA)
summary(CS2014Eur)
summary(CS2014Pac)
summary(CS2014China)
describe((CS2014NA))
describe((CS2014SA))
describe((CS2014Eur))
describe((CS2014Pac))
describe((CS2014China))
mean(c(CS2014NA$Quality,CS2014NA$Ease.of.Use,CS2014NA$Price,CS2014NA$Service))
sd(c(CS2014NA$Quality,CS2014NA$Ease.of.Use,CS2014NA$Price,CS2014NA$Service))
mean(c(CS2014SA$Quality,CS2014SA$Ease.of.Use,CS2014SA$Price,CS2014SA$Service))
sd(c(CS2014SA$Quality,CS2014SA$Ease.of.Use,CS2014SA$Price,CS2014SA$Service))
mean(c(CS2014Eur$Quality,CS2014Eur$Ease.of.Use,CS2014Eur$Price,CS2014Eur$Service))
sd(c(CS2014Eur$Quality,CS2014Eur$Ease.of.Use,CS2014Eur$Price,CS2014Eur$Service))
mean(c(CS2014Pac$Quality,CS2014Pac$Ease.of.Use,CS2014Pac$Price,CS2014Pac$Service))
sd(c(CS2014Pac$Quality,CS2014Pac$Ease.of.Use,CS2014Pac$Price,CS2014Pac$Service))
mean(c(CS2014China$Quality,CS2014China$Ease.of.Use,CS2014China$Price,CS2014China$Service))
sd(c(CS2014China$Quality,CS2014China$Ease.of.Use,CS2014China$Price,CS2014China$Service))

#3iv
CS2014Quality = data.frame('quality',CS2014$Quality)
names(CS2014Quality) = c('attribute','ratings')
CS2014EOS = data.frame('EOS',CS2014$Ease.of.Use)
names(CS2014EOS) = c('attribute','ratings')
CS2014Price = data.frame('price',CS2014$Price)
names(CS2014Price) = c('attribute','ratings')
CS2014Service = data.frame('service',CS2014$Service)
names(CS2014Service) = c('attribute','ratings')
CS2014final = rbind(CS2014Quality,CS2014EOS,CS2014Price,CS2014Service)
aovtest = aov(CS2014final$ratings~CS2014final$attribute)


#4i Mower Test Fraction
MowerTest = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetName = 'Mower Test',startRow = 4, header = TRUE)
length(which(MowerTest[,c(2:31)] == 'Fail'))/(nrow(MowerTest)*ncol(MowerTest[-1]))


#4ii Blade Weight
Bladeweight = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetName = 'Blade Weight',startRow = 3,colIndex = c(1,2), header = TRUE)
mean(Bladeweight[,2])
sd(Bladeweight[,2])/sqrt(nrow(Bladeweight))
boxplot(Bladeweight$Weight,horizontal = TRUE,main = 'Blade weight',xlab = 'Weight')


#5i on time delivery
OTD = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 11,startRow = 3, header = TRUE)
plot(OTD$Percent, type = 'l', col = 'red',axes = FALSE, main = 'On time deliveries',xlab = 'year',ylab = 'proportion of on time deliveries')
axis(1,at =1,  labels = c('2010'))
axis(1,at =13,  labels = c('2011'))
axis(1,at =25,  labels = c('2012'))
axis(1,at =37,  labels = c('2013'))
axis(1,at =49,  labels = c('2014'))
axis(2, las = 1, at=seq(0.975,1,0.005))
box()
newOTD2014 = data.frame(2014,OTD$Number.On.Time[49:60]/OTD$Number.of.deliveries[49:60])
names(newOTD2014) = c('month','NOD')
newOTD2010 = data.frame(2010,OTD$Number.On.Time[1:12]/OTD$Number.of.deliveries[1:12])
names(newOTD2010) = c('month','NOD')
newOTD = rbind(newOTD2010,newOTD2014)
t.test(newOTD$NOD~newOTD$month, alternative = 'less')


#5ii defects after delivery
DAD = read.xlsx( 'Performance Lawn Equipment Database.xlsx',sheetIndex = 12,startRow = 4,rowIndex = c(4:16), header = TRUE)
consolidated = c(DAD$X2010,DAD$X2011,DAD$X2012,DAD$X2013,DAD$X2014)
plot(consolidated,type = 'l', col = 'red', axes =FALSE, main = 'Defects after Delivery', xlab = 'Year', ylab = 'Number of Defects per million')
axis(1,at =1,  labels = c('2010'))
axis(1,at =13,  labels = c('2011'))
axis(1,at =25,  labels = c('2012'))
axis(1,at =37,  labels = c('2013'))
axis(1,at =49,  labels = c('2014'))
axis(2, las = 1, at=seq(400,900,100))
box()
newDAD2014 = data.frame(2014,DAD$X2014)
names(newDAD2014) = c('month','DAD')
newDAD2010 = data.frame(2010,DAD$X2010)
names(newDAD2010) = c('month','DAD')
newDAD = rbind(newDAD2010,newDAD2014)
t.test(newDAD$DAD~newDAD$month, alternative = 'greater')
