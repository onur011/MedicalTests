library(tm)
library(pdftools)
library(stringr)
library(stringi)
library(lubridate)
library(openxlsx)

# Folderul unde se afla analizele
path <- "C:/Users/user/Desktop/tests"
out_path <- "C:/Users/user/Desktop/output"

medical_file <- list.files(path)
nr_file <- length(medical_file)

# Functia care extrage textul din pdf
get_text_pdf <- function(dir) {
  return(pdf_text(paste(path,"/", dir, sep="")))
}

# Se extrage textul din pdf
text_files <- lapply(medical_file, get_text_pdf)

# Functia de eliminare a randului nou
remove_newline <- function(x) {
  x <- gsub("[\r\n]", "", x)
  return(x)
}

# Se elimina randurile noi
text_files <- lapply(text_files, remove_newline)

# Functia de spargere a textului dupa data
text_split_date <- function(x) {
  strsplit(x, split = "(?<=.)(?=[0-9]{2}/[0-9]{2}/[0-9]{4} [0-9]{2}:[0-9]{2}:[0-9]{2})",perl = TRUE)
}

# Se sparge textul dupa data
text_files<- lapply(text_files, text_split_date)

# Functie de eliminare a primului element din lista
remove_first_element <- function(x) {
  x <- x[- 1]
}

# Functie de eliminare a primului paragraf de pe fiecare pagina
remove_first_paragraph <- function(x) {
  x <- lapply(x, remove_first_element)
}

# Se elimina primul paragraf
text_files <- lapply(text_files, remove_first_paragraph)


# Functia de eliminarea a unei bucati din ultimul paragraf de
# pe pagina, in cazul in care acesta contine unul dintre delimitatori
extract_last_paragraph <- function(x) {
  i <- 1
  length_x <- length(x)
  
  # Pentru fiecare pagina, se verifica si se elimina o parte
  # din ultimul paragraf daca acesta contine delimitatori
  while (i <= length_x) {
    last_elem <- x[[i]][[length(x[[i]])]]
    if (!identical(last_elem, character(0))) {
      if (grepl("Medicul sef", last_elem, fixed = TRUE)) {
        aux <- strsplit(last_elem, split = "Medicul sef", perl = TRUE)
        x[[i]][length(x[[i]])] <- aux[[1]][1]
      
      } else if (grepl("http:", last_elem, fixed = TRUE)) {
        aux <- strsplit(last_elem, split = "http:", perl = TRUE)
        x[[i]][length(x[[i]])] <- aux[[1]][1]
        
      }
    }
    
    i <- i + 1
  }
  return(x)
}

# Se prelucreaza ultimul paragraf de pe pagina
text_files <- lapply(text_files, extract_last_paragraph)

# Functia de extragere a analizelor
get_values <- function(x) {
  
  nr_pages <- length(x)
  
  list_analyzes <- list()
  
  i <- 1
  
  while(i <= nr_pages) {
    nr_paragrafs = length(x[[i]])
    j <- 1
    
    while(j <= nr_paragrafs) {
      
      #date_string <- strsplit(x[[i]][[j]], split = "[ ,:/]+", perl = TRUE)
      
      date_string <- str_extract(x[[i]][[j]], "[0-9]{2}/[0-9]{2}/[0-9]{4} [0-9]{2}:[0-9]{2}:[0-9]{2}")
      
      # Se converteste in format UTC
      dates <- as.POSIXct(date_string,
                          format = "%d/%m/%Y %H:%M:%S", tz = "UTC")
      
      # Se determina numarul de seconde de la timestamp 01.01.1970
      seconds_timestamp <- as.numeric(as.POSIXct(dates, origin = "1970-01-01"))
      
      # Se memoreaza timpul in data frame
      #dt <- data.frame(time = seconds_timestamp,
       #                stringsAsFactors = FALSE)
      
      analyzes <- x[[i]][[j]]
      
      # Se elimina caracterele care nu sunt necesare
      analyzes <- str_remove_all(analyzes, "Hemoleucograma\\s+Completa\\s+\\(")
      
      analyzes <- str_remove_all(analyzes, "INR,PT\\(%\\),PT\\(s\\) \\(")
      
      analyzes <- str_remove_all(analyzes, "[0-9]{2}/[0-9]{2}/[0-9]{4} [0-9]{2}:[0-9]{2}:[0-9]{2}")
      
      name_value <- strsplit(analyzes, split = "[;,]", perl = TRUE)
      
      nr_analyzes <- length(name_value[[1]]) - 1
      
      dt_analyzes <- data.frame(name = "time",
                                value = seconds_timestamp,
                                stringsAsFactors = FALSE)
      
      k <- 1
      bad_analyze <- 0
      
      # Se extrage fiecare analiza
      while (k <= nr_analyzes) {
        
        # Se elimina spatiile care sunt in plus
        name_value[[1]][k] <- str_squish(name_value[[1]][k])
        
        
        aux_name_value <- strsplit(name_value[[1]][k], split = ':', fixed =TRUE)
        
        # Se elimina spatiile
        aux_name_value[[1]][2] <- str_remove_all(aux_name_value[[1]][2], " ")
        
        # Daca analiza este de tip text, nu este luata in considerare
        if(!grepl("^[a-zA-Z]+", aux_name_value[[1]][2], perl = TRUE)) {
           
          
          
          # Se memoreaza numele si valoarea
          aux_dt <- data.frame(name = aux_name_value[[1]][1],
                               value = stri_extract_first_regex(aux_name_value[[1]][2], 
                                                                "[+-]?([0-9]*[.])?[0-9]+"),
                                                                stringsAsFactors = FALSE)
          
          dt_analyzes <- rbind(dt_analyzes, aux_dt)
          
          k <- k + 1
        } else {
          k <- nr_analyzes + 1
          bad_analyze <- 1
        }
      }
      
      if (bad_analyze == 0) {
        list_analyzes <- append( list_analyzes, list(dt_analyzes))
      }
      j <- j + 1
    }
    
    i <- i + 1
  }
  
  return(list_analyzes)
}

# Se extrag analizele
text_files <- lapply(text_files, get_values)

i <- 1
length_text_files <- length(text_files)

while (i <= length_text_files) {
  id_pacient <- strsplit(medical_file[i], split = ".", fixed = TRUE)
  
  j <- 1
  length_analyzes <- length(text_files[[i]])
  
  out_file <- createWorkbook()

  addWorksheet(out_file, "Analizes")

  while (j <= length_analyzes) {
    col <- 4*j
    writeData(out_file, sheet="Analizes", x = text_files[[i]][[j]], startCol = col)
    
    j <- j + 1
  }
  saveWorkbook(out_file,(paste(out_path,"/", id_pacient[[1]],".xlsx", sep="")), overwrite = TRUE)
  
  i <- i + 1  
} 










