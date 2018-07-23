
args = commandArgs(trailingOnly = TRUE) # allows R to get parameters

x<-sample.int( n = 4, size=1 )
pred <- c("read", "ignore", "delete", "followUp")

cat(pred[x])# converts 3 to a number (numeric)