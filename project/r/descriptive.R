library(dplyr)
library(ggplot2)


data88 <- read.csv('data/1988.csv')
data97 <- read.csv('data/1997.csv')
data07 <- read.csv('data/2007.csv')

# mean
mean(data88$DepDelay, na.rm=TRUE)
mean(data97$DepDelay, na.rm=TRUE)
mean(data07$DepDelay, na.rm=TRUE)

# mean with NA
mean(as.numeric(is.na(data88$DepDelay)))
mean(as.numeric(is.na(data97$DepDelay)))
mean(as.numeric(is.na(data07$DepDelay)))

# median
median(c(2.3, 8.1, 5.5, 9.0, 7.8))

# trimmed mean
mean(c(0,1,2,2,2,3,4),trim=0.2)
mean(c(70,1,2,2,2,3,4),trim=0.2)

# mode
table(data88$DayOfWeek)
table(data97$DayOfWeek)
table(data07$DayOfWeek)

# variance
var(data88$DepDelay, na.rm=TRUE)

# standard deviation
sd(data88$DepDelay, na.rm=TRUE)
sd(data97$DepDelay, na.rm=TRUE)
sd(data07$DepDelay, na.rm=TRUE)

# range
range(data88$DepDelay, na.rm=TRUE)
range(data97$DepDelay, na.rm=TRUE)
range(data07$DepDelay, na.rm=TRUE)

# sample covariance
cov(data88$DepDelay, data88$ArrDelay, use="pairwise.complete.obs")

napa <- read.csv('data/napa_marathon_fm2015.csv')

ageM <- napa$Age[napa$Gender=="M"]
timeM <- napa$Hours[napa$Gender=="M"]*60
plot( ageM , timeM , pch=19 , xlab="Age in Years" , ylab="Finishing Time in Minutes" , xlim=c(10,80) , ylim=c(100,400) )
cov( ageM , timeM )

cov( ageM*12 , timeM/60 )

# correlation
cor( ageM , timeM )
cor( ageM*12 , timeM/60 )

# percentiles
quantile(data88$DepDelay,c(0.05,0.95),na.rm=TRUE)
quantile(data97$DepDelay,c(0.05,0.95),na.rm=TRUE)
quantile(data07$DepDelay,c(0.05,0.95),na.rm=TRUE)

# tabular summary
summary(data07$DepDelay)

table(data07$CancellationCode[data07$CancellationCode!=""])
table(data07$CancellationCode)
table(data07$CancellationCode[data07$UniqueCarrier=="AA"])
