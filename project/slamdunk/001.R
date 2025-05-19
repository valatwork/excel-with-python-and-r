library(tidyverse)
library(reshape2)
library(sqldf)
library(patchwork)

draft <- read.csv("draft.csv")

draft <- select(draft,-c(3,4,16:24))

draft <- draft[-c(90, 131),]

glimpse(draft)
dim(draft)

head(draft, 3)
tail(draft, 3)


draft$Year <- as.factor(draft$Year)
draft$Tm <- as.factor(draft$Tm)
draft$Born <- as.factor(draft$Born)
draft$From <- as.factor(draft$From)
draft$To <- as.factor(draft$To)

class(draft$Year)

mutate(draft, Born2 = ifelse(Born == "us", "USA", "World")) -> draft
draft$Born2 <- as.factor(draft$Born2)

draft$College[is.na(draft$College)] <- 0
mutate(draft, College2 = ifelse(College == 0, 0, 1)) -> draft
draft$College2 <- as.factor(draft$College2)


levels(as.factor(draft$Pos))

draft$Pos2 <- draft$Pos
draft$Pos2 <- recode(draft$Pos2,
 "C" = "Center",
 "C-F" = "Big",
 "F" = "Forward",
 "F-C" = "Big",
 "F-G" = "Swingman",
 "G" = "Guard",
 "G-F" = "Swingman")
draft$Pos <- as.factor(draft$Pos)
draft$Pos2 <- as.factor(draft$Pos2)
levels(draft$Pos2)

head(draft)

summary(draft)


sd(draft$G)
sd(draft$MP)
sd(draft$WS)

sqldf("SELECT min(WS), Player, Tm, Pk, Year FROM draft")
sqldf("SELECT max(WS), Player, Tm, Pk, Year FROM draft")
sqldf("SELECT min(G), Player, Tm, Pk, Year FROM draft")
sqldf("SELECT max(G), Player, Tm, Pk, Year FROM draft")
sqldf("SELECT min(MP), Player, Tm, Pk, Year FROM draft")
sqldf("SELECT max(MP), Player, Tm, Pk, Year FROM draft")
sqldf("SELECT min(Age), Player, Tm, Pk, Year FROM draft")
sqldf("SELECT max(Age), Player, Tm, Pk, Year FROM draft")


p1 <- ggplot(draft, aes(x = WS)) +
 geom_histogram(fill = "royalblue3", color = "royalblue3",
 bins = 8) +
 labs(title = "Career Win Shares Distribution of
 NBA First-Round Selections",
 subtitle = "2000-09 NBA Drafts",
 x = "Career Win Shares",
 y = "Frequency") +
 theme(plot.title = element_text(face = "bold"))
print(p1)

sqldf("SELECT COUNT (*) FROM draft WHERE WS >= 75")
sqldf("SELECT COUNT (*) FROM draft WHERE WS < 75")
sqldf("SELECT COUNT (*) FROM draft WHERE WS <= 25")

p2 <- ggplot(draft, aes(x = College2, y = WS)) +
 geom_boxplot(color = "orange4", fill = "orange1") +
 labs(title = "Career Win Shares Distribution of
 NBA First-Round Selections",
 x = "",
 y = "Career Win Shares",
 subtitle = "2000-09 NBA Drafts") +
 stat_summary(fun = mean, geom = "point", shape = 20,
 size = 8, color = "white", fill = "white") +
 theme(plot.title = element_text(face = "bold")) +
 facet_wrap(~Born2) +
 scale_x_discrete(breaks = c(0, 1),
 labels = c("No College", "College"))
print(p2)