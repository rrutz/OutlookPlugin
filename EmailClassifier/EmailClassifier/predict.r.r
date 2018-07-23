par <- as.data.frame(read.csv("C:\\Users\\Ruedi\\OneDrive\\MS\\OutlookPlugin\\EmailClassifier\\pamramters.csv"))

args = commandArgs(trailingOnly = TRUE) # allows R to get parameters
body <- args[1]
words <- as.data.frame(strsplit(body, split = " ")[[1]] )
colnames(words) <- c("word")
probs <- merge(x=words,y=par, by = "word", all.x=TRUE )[, c(3:5)]
probs[is.na(probs)] <- 1

pred <- which.max(sapply(probs, FUN=prod))
if ( length(pred ) == 0 )
{
  cat("error")

} else
{
  if ( pred == 1 )
  {
    cat("read")
  } else if ( pred == 2 )
  {
    cat("delete")
  } else if ( pred == 3 )
  {
    cat("followUp")
  } 
}