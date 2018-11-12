library(stringr)
library(readr)
library(mgsub)
library(dplyr)
library(readr)
library(readtext)
#https://regex101.com/
#codigo <- read_file("Vida_Individual_con_Ifs.txt")

VBA_to_R <- function(codigo){
  
  # Quitar el Public --------------------------------------------------------
  nuevo_texto <- mgsub(codigo,unlist(str_extract_all(codigo,"Public")),rep("",length(unlist(str_extract_all(codigo,"Public")))))
  
  # Funciones ---------------------------------------------------------------
  
  patterns = unlist(str_extract_all(nuevo_texto,"Function .+"))
  replacements <- paste0(patterns,"{")
  nuevo_texto1 <- mgsub(string = nuevo_texto, unlist(str_extract_all(nuevo_texto,"Function .+")), replacement = replacements, fixed =TRUE)
  
  patterns = unlist(str_extract_all(nuevo_texto1,"Function .+?(?=[(])"))
  replacements = paste(unlist(str_extract_all(nuevo_texto1,"(?<=Function ).+?(?=[(])")),"<- function")
  
  nuevo_texto2 <- mgsub(string = nuevo_texto1, pattern = unlist(patterns),replacement = replacements) %>% 
    mgsub("End Function","}") %>% 
    mgsub("'","#")
  # Quita los As...
  # Más adelante se podría identificar cuáles son Range para identificar cuándo poner corchetes ´[]´ donde hay parénteseis ()
  patterns <- unlist(str_extract_all(nuevo_texto2,"As.+?(?=(,|[)]))"))
  
  nuevo_texto3 <- mgsub(nuevo_texto2,patterns,rep("",length(patterns)))
  
  patterns <- unlist(str_extract_all(nuevo_texto3,"For .+"))
  replacements <- paste0(patterns,")){")
  nuevo_texto4 <- mgsub(string = nuevo_texto3, patterns , replacements, fixed =TRUE)
  
  #For por for (futuro: sapply)
  patterns <- unlist(str_extract_all(nuevo_texto4,"For "))
  nuevo_texto5 <- mgsub(string = nuevo_texto4, patterns , rep("for(",length(patterns)), fixed =TRUE)
  
  
  #Cambiar los next por }
  patterns <- unlist(str_extract_all(nuevo_texto5,"Next.*"))
  nuevo_texto6 <- mgsub(string = nuevo_texto5, patterns , rep("}",length(patterns)), fixed =TRUE)
  
  #To por :
  patterns <- unlist(str_extract_all(nuevo_texto6," To "))
  nuevo_texto7 <- mgsub(string = nuevo_texto6, patterns , rep(":(",length(patterns)), fixed =TRUE)
  # = por in
  patterns <- unlist(str_extract_all(nuevo_texto7,"(?<=for[(]).*([=])?"))
  replacements <- mgsub(string = patterns, rep("=", length(patterns)) , rep("in",length(patterns)), fixed =TRUE)
  
  
  #Ifs y operadores lógicos
  nuevo_texto8 <- mgsub(string = nuevo_texto7, patterns , replacements, fixed =TRUE) %>% 
    mgsub("If ", "if(") %>% 
    mgsub(" And ","&") %>% 
    mgsub(" Or ","|")
  
  
  #Quitar los Dim
  patterns <- unlist(str_extract_all(nuevo_texto8,"Dim.+"))
  nuevo_texto9 <- mgsub(nuevo_texto8 ,patterns,rep("",length(patterns)))
  
  # No es limpio lo del entender cuándo hay un Else y cuándo no por los If anidados. Si no hay anidados, no hay problema con este código.
  
  # patterns <- unlist(str_extract_all(nuevo_texto9,"End If"))
  # probables_if_else <- unlist(str_extract_all(nuevo_texto9,"Then[\\s\\S]*?End If"))
  # cuales <- which(c(str_detect(probables_if_else,"Else")))
  # 
  # replacements <- rep("}", length(probables_if_else))
  # replacements[cuales] <- "}\n}"
  # 
  # nuevo_texto10<- nuevo_texto9
  # for (i in 1:length(patterns)){
  #   nuevo_texto10 = sub(patterns[i], replacements[i], nuevo_texto10)
  # }
  
  nuevo_texto10 <-  mgsub(nuevo_texto9,"Else","}else{") %>% 
    mgsub(" Then", replacement = "){", fixed =TRUE) %>% 
    mgsub("End If", replacement = "}", fixed =TRUE)
  
  cat(nuevo_texto10, file = here("Codigo_Cambiado",paste0("VBA-R_",format(now(), "%Y%m%d_%H%M%S"),".txt")))
}

#readtext no necesita el encoding para las tildes y demás. read_file es de readr (tidyverse).


codigo <- readtext("Vida_Individual.txt")$text
codigo <- read_file("Vida_Individual.txt",locale = locale("es", encoding = "Cp1252"))


VBA_to_R(codigo)





# Cambios en código -------------------------------------------------------












# Código viejo ------------------------------------------------------------

VBA_to_R <- function(codigo){
  
  
  # Quitar el Public --------------------------------------------------------
  s <- codigo
  patterns <- unlist(str_extract_all(s,"Public"))
  
  nuevo_texto <- mgsub(s,patterns,rep("",length(patterns)))
  
  # Funciones ---------------------------------------------------------------
  
  patterns = unlist(str_extract_all(nuevo_texto,"Function .+"))
  replacements <- paste0(patterns,"{")
  nuevo_texto1 <- mgsub(string = nuevo_texto, pattern = patterns, replacement = replacements, fixed =TRUE)
  
  patterns = unlist(str_extract_all(nuevo_texto1,"Function .+?(?=[(])"))
  replacements = paste(unlist(str_extract_all(nuevo_texto1,"(?<=Function ).+?(?=[(])")),"<- function")
  nuevo_texto2 <- mgsub(string = nuevo_texto1, pattern = unlist(patterns),replacement = replacements)
  
  nuevo_texto3 <- mgsub(nuevo_texto2, "End Function","}")
  #cat(nuevo_texto2)
  #cat(nuevo_texto3)
  
  # Comentarios -------------------------------------------------------------
  nuevo_texto4 <- mgsub(nuevo_texto3,"'","#")
  # Las definiciones de las variables ---------------------------------------
  # Sería valioso identificar los Range antes. 
  patterns <- unlist(str_extract_all(nuevo_texto4,"As.+?(?=(,|[)]))"))
  
  nuevo_texto5 <- mgsub(nuevo_texto4,patterns,rep("",length(patterns)))
  
  
  
  patterns <- unlist(str_extract_all(nuevo_texto5,"For .+"))
  replacements <- paste0(patterns,")){")
  nuevo_texto6 <- mgsub(string = nuevo_texto5, patterns , replacements, fixed =TRUE)
  
  
  
  
  #For por for
  patterns <- unlist(str_extract_all(nuevo_texto6,"For "))
  nuevo_texto7 <- mgsub(string = nuevo_texto6, patterns , rep("for(",length(patterns)), fixed =TRUE)
  
  
  
  
  #Cambiar los next por }
  patterns <- unlist(str_extract_all(nuevo_texto7,"Next.*"))
  nuevo_texto8 <- mgsub(string = nuevo_texto7, patterns , rep("}",length(patterns)), fixed =TRUE)
  
  
  #To por :
  patterns <- unlist(str_extract_all(nuevo_texto8," To "))
  nuevo_texto9 <- mgsub(string = nuevo_texto8, patterns , rep(":(",length(patterns)), fixed =TRUE)
  
  
  
  
  patterns <- unlist(str_extract_all(nuevo_texto9,"(?<=for[(]).*([=])?"))
  replacements <- mgsub(string = patterns, rep("=", length(patterns)) , rep("in",length(patterns)), fixed =TRUE)
  
  nuevo_texto10 <- mgsub(string = nuevo_texto9, patterns , replacements, fixed =TRUE)
  
  nuevo_texto11 <- mgsub(string = nuevo_texto10, " Then", replacement = "){", fixed =TRUE)
  
  nuevo_texto12 <- mgsub(string = nuevo_texto11, "If ", "if(")
  nuevo_texto13 <- mgsub(nuevo_texto12, "End If","}")
  nuevo_texto14 <- mgsub(nuevo_texto13, " And ","&")
  nuevo_texto15 <- mgsub(nuevo_texto14, " Or ","|")
  
  
  patterns <- unlist(str_extract_all(nuevo_texto15,"Dim.+"))
  
  nuevo_texto16 <- mgsub(nuevo_texto15,patterns,rep("",length(patterns)))
  
  cat(nuevo_texto16, file = "Caroli.txt")
}

VBA_to_R_pipes(codigo)





