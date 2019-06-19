library(tidyverse)
library(docxtractr)
doc <- read_docx("data/ТОВ КУА ГЕНЕЗІС.docx")


docx_describe_tbls(doc)


docx_extract_tbl(doc, 1, header = F)




res1 <- docx_extract_tbl(doc, header = F, 23)


extract_all <- function(doc_name){
  
  doc <- read_docx(doc_name)
  main_info <- docx_extract_tbl(doc, 1,header = F)
  code <- main_info[1,2] %>% as.numeric()
  name_of_bussines <- main_info[5,2] %>% as.character()
  
  doc %>% 
     map_dbl(.,~(docx_extract_tbl))
  
  
  main_table <- docx_extract_tbl(doc, doc$tbls %>% length()-2,header = T)
  if(nrow(main_table)>0){
    main_table <- main_table %>% 
    magrittr::set_colnames(c("Number","","Code_valute","Valute","MFO","Bank","rahunok","open","close")) %>% 
    select(MFO:close) %>% 
    mutate(close = ifelse(close=="",NA,close)) %>% 
    mutate_at(vars(open:close),lubridate::dmy) %>% 
    filter(is.na(close)) %>% 
    add_column(bus = code, .before = 1) %>% 
    add_column(name = name_of_bussines, .before = 2) %>% 
    distinct()}
  
  return(main_table)
}

path<-c(paste0(getwd(),"/data"))

data <- data_frame(filename =  list.files(path)) %>%  ## find all files in folder
  mutate(file_contents = map(paste0("data/",filename),          
                             ~extract_all(.))) %>% ## read each file without colnames
  unnest()

#extract_all("data/ПАТ НАСК ОРЕНТА 2.docx")

basic <- data %>% 
  group_by(bus,name, MFO, Bank) %>% 
  summarise(text = glue::glue_collapse(rahunok,", "))

un_mfo <- basic %>% 
  ungroup %>% 
  select(MFO,Bank) %>% distinct() 

make_c <- function(df){
  titl <- df$main_text[1]
  txt <- glue_collapse(df$full_text,sep = " \n")
  tog <- glue::glue('
                   {titl}
                   {txt}  
                   ')
  return(tog)
}


tog_list <- un_mfo$MFO %>%   
  map_chr(.,~{
   my_list <-  basic %>% filter(MFO  %in% ..1)

 loc_res <-   my_list %>% 
     mutate(main_text = glue::glue('МФО {MFO} {Bank}'),
            full_text = glue::glue('{name} ({bus}) рахунок № {text}'))
res <- make_c(loc_res)
   
   })

tog_list %>% unlist() %>% c ->qwe

