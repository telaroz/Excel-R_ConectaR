# path <- "Ocurrencia.xlsx"
# %>% 
#   excel_sheets() %>% 
#   map2(read_excel, path = path, sheet = excel_sheets(path))
# read_ex
# 
# lapply(excel_sheets(path), function(i) read_excel("Ocurrencia.xlsx", sheet = i))
# 
# library(dplyr)
# library(purrr)
# library(readxl)
# 
# all_sheets <-
#   "Ocurrencia.xlsx" %>%
#   excel_sheets() %>%
#   set_names() %>%
#   map(read_excel, path = "Ocurrencia.xlsx",col_names = FALSE,col_types = "text")


library(readxl)
library(tidyxl)
library(dplyr)
library(purrr)
library(tidyr)

#Cargar todas las pestañas a la misma vez:


# path to spreadsheet
# file <- "Ocurrencia.xlsx"
# # get sheet names 
# sheets <- excel_sheets("Ocurrencia.xlsx")
# 
# # read rectangular
# all_sheets <-
#   file %>%
#   excel_sheets() %>%
#   set_names() %>%
#   map(read_excel, path = file,col_names = FALSE,col_types = "text")
# 
# list2env(all_sheets,globalenv())

# path to spreadsheet
file <- "exclread.xlsx"
# get sheet names 
sheets <- excel_sheets("exclread.xlsx")

# read rectangular
all_sheets <-
  file %>%
  excel_sheets() %>%
  set_names() %>%
  map(read_excel, path = file,col_names = FALSE,col_types = "text")

# label each DF with categories, bind rows, add numbering, rename vars,
all_qs_lab <- 
  all_sheets %>% map2(names(all_sheets),~mutate(.x,name=.y)) %>%
  reduce(rbind) %>% tibble::rowid_to_column() %>% 
  select(question_number=1,question=2,a=3,b=4,c=5,d=6,category=7) %>% 
  select(question,everything()) %>% 
  group_by(category) %>% mutate(bycat_number = row_number()) %>% ungroup



all_cells <- xlsx_cells(file,sheets = NA) %>% filter(is_blank!=TRUE)
formats <- xlsx_formats(file)


isBold <- formats[["local"]][["font"]][["bold"]]

correct_answers <- all_cells[all_cells$local_format_id %in% which(isBold),
                             c("sheet","row","character","numeric")]

correct_answers$boldanswer <- 
  ifelse(is.na(correct_answers$character),
         yes = correct_answers$numeric, no= correct_answers$character )
correct_answers <- correct_answers %>% select(-character,-numeric) %>% tibble::rowid_to_column()
