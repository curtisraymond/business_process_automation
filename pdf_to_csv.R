if("pacman" %in% rownames(installed.packages()) == FALSE) {install.packages("pacman")} # Check if you have universal installer package, install if not

pacman::p_load(pdftools, tidyverse)

files <- list.files(path = "C:/Users/curtis/Desktop/pdf_cvs/", pattern = "*.pdf", full.names = T)
file_names <- list.files(path = "C:/Users/curtis/Desktop/pdf_cvs/", pattern = "*.pdf", full.names = F)

for (i in 1:length(files)){
  text <- pdf_text(files[i])
  text2 <- unlist(str_split(text, "[\\r\\n]+"))
  text3 <- str_split_fixed(str_trim(text2), "\\s{2,}", Inf)
  text3 <- as.data.frame(text3)
  write.csv(text3, paste0("C:/Users/curtis/Desktop/pdf_cvs/", gsub('.{4}$', '', file_names[[i]]), ".csv"))
}