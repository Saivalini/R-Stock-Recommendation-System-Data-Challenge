rm(list=ls())

# Please install the following packages, if not present.  
library(PortfolioAnalytics)
library(openxlsx)
library(ROI)
library(ROI.plugin.glpk)
library(ROI.plugin.quadprog)
library(ROI)
library(tidyquant) # tq_get
library(dplyr) # group_by 
library(timetk)
library(forcats)# fct_reorder
library(tidyr) # spread()

# Set the working directory, this is where the final output will be saved finally. 
setwd("C:\\Users\\saiva\\Downloads\\Appreciate")

# Read the training data given by the Appreciate Team:
# Please be watchful that the variable names created here, and the sheet numbers in the main training data dont match. 
# For instance, here data1, has data from sheet-2 of the main training data and so on. 
# Please retain the naming as it is, as these files are again referenced in the code below. 

data1 = read.xlsx("C:\\Users\\saiva\\Downloads\\20201022-Training Dataset_v1(1) (1).xlsx",sheet=2)
data2 = read.xlsx("C:\\Users\\saiva\\Downloads\\20201022-Training Dataset_v1(1) (1).xlsx",sheet=3)
data3 = read.xlsx("C:\\Users\\saiva\\Downloads\\20201022-Training Dataset_v1(1) (1).xlsx",sheet=1)

#View(data1)
#View(data2)
#View(data3)

# Extracting data from yahoo finance and calculating daily returns
tickers = c(unique(data1$Ticker))
data <- data.frame(tq_get(tickers,
                          from = "2015-01-01", 
                          to = "2020-12-01", 
                          get = "stock.prices"))
log_ret_tidy <- data %>%
  group_by(symbol) %>%
  tq_transmute(select = adjusted,
               mutate_fun = periodReturn,
               period = 'daily',
               col_rename = 'ret',
               type = 'log')

log_ret_xts <- log_ret_tidy %>%
  spread(symbol, value = ret) %>%
  tk_xts()

# Rebalancing customer's Portfolio based on the weights suggested through optimization, 
# adding one stock at a time to the existing portfolio of the customer, and calulating all the metrics. 
# Returning top-5 stocks based on Max Sharpe Ratio's 

new2 = function(Cust_1_tick){
  Cust_1_port_ret = c()
  Cust_1_port_sd = c()
  SR = c()
  best = c()
  best1 = data.frame()
  Portfolio = data.frame()
  for (i in 1:length(tickers[!tickers %in% Cust_1_tick$Ticker])){ 
    Cust_1_DR=subset(log_ret_xts, select = c(Cust_1_tick$Ticker,tickers[!tickers %in% Cust_1_tick$Ticker][i]))
    Cust_1_DR[is.na(Cust_1_DR)] <- 0
    Cust_1_MR = data.frame("Ticker" = colnames(Cust_1_DR),"Returns"= colMeans(Cust_1_DR,na.rm=TRUE))
    rownames(Cust_1_MR) <- 1:nrow(Cust_1_MR)
    Cust_1_cov_mat <- round(cov(Cust_1_DR, use = "complete.obs") * 252,5)
    init.portf <- portfolio.spec(assets=colnames(Cust_1_DR))
    init.portf <- add.constraint(portfolio=init.portf, type="full_investment") 
    init.portf <- add.constraint(portfolio=init.portf, type="long_only")
    init.portf <- add.objective(portfolio=init.portf, type="return", name="mean")
    init.portf <- add.objective(portfolio=init.portf, type="risk", name="StdDev")
    maxSR <- optimize.portfolio(R=Cust_1_DR, portfolio=init.portf, 
                                optimize_method="ROI", 
                                Rf = 0.03,
                                maxSR=TRUE, trace=TRUE)
    Cust_1_port_ret <- sum(maxSR$weights * Cust_1_MR$Returns)
    Cust_1_port_ret[i] <- ((Cust_1_port_ret + 1)^252) - 1
    Cust_1_port_sd[i] <- sqrt(t(maxSR$weights) %*% (Cust_1_cov_mat  %*% maxSR$weights))
    SR[i] <- (Cust_1_port_ret[i]-0.03)/Cust_1_port_sd[i] 
    best <-  data.frame(round(maxSR$weights,10), SR[i], Cust_1_port_ret[i],Cust_1_port_sd[i])
    best1 <- rbind(best1,best)
  }
  best1 = best1[order(best1$SR, decreasing = TRUE),]
  best1 = cbind(rownames(best1), best1, stringsAsFactors = FALSE)
  rownames(best1) = 1:nrow(best1)
  colnames(best1)=c("Ticker","Weights","SharpeRatio","PortfolioReturn","PortfolioRisk")
  A=best1[row.names(best1[seq((length(Cust_1_tick$Ticker)+1), nrow(best1), (length(Cust_1_tick$Ticker)+1)),][2]>0),]
  colnames(A)=c("Ticker","Weights","SharpeRatio","PortfolioReturn","PortfolioRisk")
  Sug_Tick=head(A[which(A$Weights > 0),])[1]
  Total=best1[best1$SharpeRatio %in% head(A[which(A$Weights > 0),])$SharpeRatio,]
  for (i in 1 : length(unique(Total$SharpeRatio))){
    Check = split(Total, fct_inorder(factor(Total$SharpeRatio),))
    assign(paste0("Portfolio",i),Check[[i]])
  }
  Final_Suggestions <- list()
  Final_Suggestions$Top_Tickers = Sug_Tick[1]
  Final_Suggestions$Portfolio = mget(ls(pattern = "^Portfolio.$"))
  cat(paste0("Top Stock Recommendations for your existing Portfolio are:", Sug_Tick[1]),sep="\n") 
  return(Final_Suggestions)
}

# storing data1 into data5 
data5 = data1

# Running the the function "new2" created above for all the customers in the data and saving the results: 
for (i in 1: length(unique(data5$Customer_ID))){
  cat(paste0("Customer:", unique(data5$Customer_ID)[i], sep = "\n"))
  cust = new2(data5[data5$Customer_ID==unique(data5$Customer_ID)[i],])
  dir.create(paste0("Customer"," ",unique(data5$Customer_ID)[i]), showWarnings = FALSE) 
  write.csv(cust$Top_Tickers, file.path(paste0("Customer"," ",unique(data5$Customer_ID)[i]), "Top-Tickers-Suggested-Market.csv"), row.names=FALSE) 
  write.csv(cust$Portfolio, file.path(paste0("Customer"," ",unique(data5$Customer_ID)[i]), "Rebalanced-Portfolio-Metrics-Market.csv"), row.names=FALSE)
}

# References: 
# 1) https://www.codingfinance.com/post/2018-05-31-portfolio-opt-in-r/ (Yahoo finance data extraction + formulas for annualized returns, risk) 
# 2) https://cran.r-project.org/web/packages/PortfolioAnalytics/PortfolioAnalytics.pdf




