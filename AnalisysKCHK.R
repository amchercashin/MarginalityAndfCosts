library(openxlsx)
library(lubridate)
library(dplyr)
library(ggplot2)
library(scales)
library(tidyr)
library(grid)

space <- function(x, ...) {format(x, ..., big.mark = " ", scientific = FALSE, trim = TRUE)}

###Data reading and processing
#Read fixed costs for fact 2014, select only several columns
costs_f14 <- read.xlsx("./data/КонтрольФБ.xlsx", sheet = 19)
costs_f14 <- tbl_df(data.frame(companygroup = costs_f14[,11], scenario = "Факт",
                               month = costs_f14[,15], fixed.costs = costs_f14[,10], ico = costs_f14[,18]))

#Read fixed costs for budget 2015, select only several columns
costs_bp15 <- read.xlsx("./data/ФБ2015_Консолидация_все_затраты.xlsx", sheet = 8)
costs_bp15 <- tbl_df(data.frame(companygroup = costs_bp15[,11], scenario = "Бюджет",
                                month = costs_bp15[,15], fixed.costs = costs_bp15[,10], ico = costs_bp15[,19]))
#Costs processing
costs <- rbind(costs_f14, costs_bp15) #combine two data.frames
costs <- costs[!is.na(costs$fixed.costs),] #filter out all NA in costs sums
costs <- filter(costs, companygroup == "Группа ВМУ" | companygroup == "Группа АЗОТ" | companygroup == "Минеральные удобрения, ОАО"
                | companygroup == "Группа КЧХК" | companygroup == "ПМУ" ) #take only plants costs
#costs <- filter(costs, ico == 0) #filter out all ICO costs
costs$month <- dmy("01.01.1900") + days(costs$month) - days(2) #format dates to POSIXct
#Summarise
costs <- group_by(costs, scenario, companygroup, month) %>% summarise(fixed.costs = sum(fixed.costs))
#To narrow tidy form
costs <- gather(costs, indicator, value, -(scenario:month))

#Read revenue and profit data for plants, select only several columns
marg <- read.xlsx("./data/DBmarginality.xlsx", sheet = 1)
marg <- tbl_df(data.frame(companygroup = marg[,3], scenario = marg[,5],
                          month = marg[,16], revenue = marg[,32], marginal.profit = marg[,42], ico = marg[,45]))
#Revenue and profit data processing
marg$scenario <- toupper(marg$scenario) #all scenario charachters to upper case
marg$month <- dmy("01.01.1900") + days(marg$month) - days(2) #date to POSIXct
#Take only fact for 2014 and budget for 2015
marg <- filter(marg, scenario == "БП" & month >= dmy("01.01.2015") | scenario == "ФАКТ" & month >= dmy("01.01.2014"))
marg <- filter(marg, companygroup != "SBU-Nitrotrade AG") #filter out non Uralchem company
marg$scenario[marg$scenario == "ФАКТ"] <- "Факт" #back to normal spelling
marg$scenario[marg$scenario == "БП"] <- "Бюджет" #back to normal spelling
marg$revenue[is.na(marg$revenue)] <- 0 #all NA revenue to zeros
marg$marginal.profit[is.na(marg$marginal.profit)] <- 0 #all NA marginal profit to zeros
#Summarise
marg <- group_by(marg, scenario, companygroup, month) %>% summarise(revenue = sum(revenue), marginal.profit = sum(marginal.profit))
#To narrow tidy form
marg <- gather(marg, indicator, value, -(scenario:month))
marg$value <- marg$value * 1000 #from thousands to units
#Combine two datasets
data <- rbind(costs, marg)

#Transform company group names to short names
data$companygroup <- as.character(data$companygroup)
data[grep("КЧХК", data$companygroup, ignore.case = T), "companygroup"] <- "КЧХК"
data[grep("АЗОТ", data$companygroup, ignore.case = T), "companygroup"] <- "АЗОТ"
data[grep("ВМУ", data$companygroup, ignore.case = T), "companygroup"] <- "ВМУ"
data[grep("Минеральные удобрения", data$companygroup, ignore.case = T), "companygroup"] <- "ПМУ"
#Back companygroup names to factors 
data$companygroup <- factor(data$companygroup, levels = c("АЗОТ", "КЧХК", "ПМУ", "ВМУ"))
#Indicators factor reoder
data$indicator <- factor(data$indicator, levels = c("revenue", "marginal.profit", "fixed.costs")
                         , lables = c("Выручка", "Маржинальная прибыль", "Постоянные затраты"))

##FIGURES
my_grob = grobTree(textGrob("Ф 14", x=0.2,  y=0.85, hjust=0,
                            gp=gpar(fontsize=22, col = "darkgrey", fontface = "bold")))
my_grob2 = grobTree(textGrob("БП 15", x=0.8,  y=0.85, hjust=0,
                             gp=gpar(fontsize=22, col = "darkgrey", fontface = "bold")))

png("Plants_indicators.png", width=3308/2, height=2339/2)
p <- ggplot(data, aes(x = month, y = value/10^6)) +
        facet_wrap(~ companygroup, ncol = 2, scales = "free_y") +
        geom_area(aes(fill = indicator), alpha = 0.65, position = "identity") +
        #geom_smooth(aes(colour = scenario), method="loess", se = F) +
        scale_y_continuous(labels = space) +
        ylab("Миллионов рублей") +
        xlab(NULL) +
        ggtitle("Выручка и маржинальная vs постоянные затраты") +
        theme_bw(base_size = 18) +
        theme(legend.position="right", legend.title=element_blank(), legend.key.size = unit(1.5,"cm")) +
        theme(axis.text=element_text(size=16), axis.title=element_text(size=18), strip.text.x=element_text(size=16), legend.text=element_text(size=16)) +
        theme(axis.text.x  = element_text(hjust=1,angle=90, vjust=0.5)) +
        scale_x_datetime(labels = date_format("%b.%y"), breaks = date_breaks("month"), minor_breaks = NULL) +
        scale_fill_manual(values = c("#66CC99", "#9999CC", "#CC6666")) +
        geom_vline(xintercept = as.numeric(dmy("01.01.2015")), colour = "darkgrey", size = 1) +
        annotation_custom(my_grob) +
        annotation_custom(my_grob2) +
        coord_cartesian(ylim = c(0, 4.25*10^3), xlim = c(dmy("01.01.2014"), dmy("01.12.2015")))

print(p)
dev.off()