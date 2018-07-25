# very simple naive bayes model


# get data
df.emails <- as.data.frame(read.csv("C:\\Users\\Ruedi\\OneDrive\\MS\\OutlookPlugin\\EmailClassifier\\data2.csv"), header=TRUE)
df.emails$body <- tolower(df.emails$body )


# get all words
words <- c()
for ( body in df.emails$body)
{
  words <- c(words, strsplit(body, split = " ")[[1]])
}


# create frequencies
words.freq<-table(unlist(words))
words.freq = as.data.frame(cbind(names(words.freq),as.integer(words.freq)) )
words.freq$V2 <- as.numeric(words.freq$V2)

getFreq <- function( label, df_in)
{
  classEmails <- df.emails[ df.emails$class == label, ]
  words_tmp <- c()
  for ( body in classEmails$body)
  {
    words_tmp <- c(words_tmp, strsplit(body, split = " ")[[1]])
  }
  words.freq_tmp<-table(unlist(words_tmp))
  words.freq_tmp = as.data.frame(cbind(names(words.freq_tmp),as.integer(words.freq_tmp)) )
  words.freq_tmp$V2 <- as.numeric(words.freq_tmp$V2)
  words.freq[is.na(words.freq)] <- 0
  words.freq_tmp$V2 <- words.freq_tmp$V2 / sum(words.freq_tmp$V2) + 0.5
  words.freq <- merge(x=df_in, y=words.freq_tmp, by ="V1", all.x = TRUE )
  return(words.freq)
}

words.freq <- getFreq("Read", words.freq)
words.freq <- getFreq("Delete", words.freq)
words.freq <- getFreq("Follow Up", words.freq)

colnames(words.freq) <- c( "word", "totalCount", "readFreq", "deleteFreq", " followUpFreq")
words.freq[is.na(words.freq)] <- 1

write.csv(x = words.freq, file = "C:\\Users\\Ruedi\\OneDrive\\MS\\OutlookPlugin\\EmailClassifier\\pamramters.csv", row.names = FALSE)

