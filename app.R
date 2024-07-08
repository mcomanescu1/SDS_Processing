# Required Libraries
library(shiny)
library(shinyFiles)
library(pdftools)
library(data.table)
library(stringr)
library(tm)
library(readxl)
library(openxlsx)
library(polyglotr)
library(spsComps)

# UI Code
ui <- fluidPage(
  titlePanel("Aplicație de procesare a fișelor SDS"),
  sidebarLayout(
    sidebarPanel(
      shinyDirButton("dir", "Alegeți Locația Fișelor", "PDF"),
      textInput("filename", "Alegeți numele tabelului ce va fi salvat", value = "output"),
      actionButton("process", "Procesare Fișe")
    ),
    mainPanel(
      verbatimTextOutput("info"),
      uiOutput("progress_ui")
    )
  )
)

# Server Code
server <- function(input, output, session) {
  # Set up shinyFile buttons
  setwd(getSrcDirectory(function(){})[1])

  GHS <- read_xlsx(path = 'GHS.xlsx')
  GHS2 <- read_xlsx(path = 'GHS2.xlsx')
    Merck<-c("FIŞA CU DATE DE SECURITATE", "în conformitate cu Reglementările UE No. 1907/2006", "Aldrich-", "Millipore-", "SIGALD-", "Sigma-Aldrich-")
    Roth<-c("Fişa cu date de securitate", "în conformitate cu Regulamentul (CE) nr. 1907/2006 (REACH), modificat prin 2020/878/EU", "România (ro)", "număr articol:", "Pagina")
    language<-c("en","ro")
  
  setwd("C://")
  volumes<- setNames(c('.', 'c:'), c(getwd(), "home"))
   shinyDirChoose(input, id = 'dir', defaultRoot = getwd(),roots = volumes)
  
   input_folder<- reactive({
     req(input$dir)
     return(parseDirPath(volumes, input$dir))
   })
  
output$info <- renderPrint({
  list(
    Folder_SDS = input_folder(),
    Folder_Salvare = "Tabelul dumneavoastră va fi salvat în folderul conținând fișele SDS",
    Nume_Tabel = input$filename
  )
})
  
  observeEvent(input$process, {
    req(input_folder(), input$filename)
    output$info <- renderPrint({
      list(
        Folder_SDS = input_folder(),
        Folder_Salvare = "Tabelul dumneavoastră va fi salvat în folderul conținând fișele SDS",
        Nume_Tabel = input$filename
      )
    })
    # Progress Bar Setup
    progress <- shiny::Progress$new()
    on.exit(progress$close())
    progress$set(message = "Se procesează", value = 0)
    
    # Set working directory and get list of files (case insensitive for .pdf and .PDF)
    setwd(input_folder())
    files <- list.files(pattern = "\\.pdf$", ignore.case = TRUE)
    
    # Initialize the result matrix
    file.names <- sapply(files, function(f) {
      strsplit(f,".pdf")[[1]]
      })
    file.names <- sapply(file.names, function(f) {
      strsplit(f,".PDF")[[1]]
    })
  
    tabel <- matrix(nrow = length(files), ncol = 11)
    colnames(tabel) <- c("Fișă SDS", "Nr.Cas", "Proprietăți Fizio-Chimice", "Clasificare", "Riscuri pentru sănătate",
                         "Riscuri pentru mediu", "Mijloace de Protecție", "Mijloace de Intervenție", "Măsuri de Prim Ajutor", 
                         "Mijloace de Stingere", "Măsuri de Decontaminare")
    tabel[,1] <- as.character(file.names)
    
    # Function to read text from PDF and concatenate
    read_pdf_text <- function(file) {
      tryCatch({
        text <- pdf_text(file)
        return(paste(text, collapse = " "))
      }, error = function(e) {
        message(paste("Error reading file:", file, "\n", e))
        return(NULL)
      })
    }
    
    # Function to process individual file
    process_file <- function(file, GHS) {
      text <- read_pdf_text(file)
      # text<-read_pdf_text(files[1])
      if (is.null(text)) return(NULL)
      # text<-read_pdf_text(files[3])
      is_rom <- grepl('conformitate', text, ignore.case = TRUE)
      if (!is_rom) {
        text<-str_flatten(gsub("[\r\n]", " ", text))
        sections <-c(unlist(gregexpr('SECTION 1', text))[1],
                     unlist(gregexpr('SECTION 2', text))[1],
                     unlist(gregexpr('SECTION 3', text))[1],
                     unlist(gregexpr('SECTION 4', text))[1],
                     unlist(gregexpr('SECTION 5', text))[1],
                     unlist(gregexpr('SECTION 6', text))[1],
                     unlist(gregexpr('SECTION 7', text))[1],
                     unlist(gregexpr('SECTION 8', text))[1],
                     unlist(gregexpr('SECTION 9', text))[1],
                     unlist(gregexpr('SECTION 10', text))[1],
                     unlist(gregexpr('SECTION 11', text))[1])
        text2<-paste(create_translation_table(substr(text, sections[1], sections[2]), language)[,3],
                     create_translation_table(substr(text, sections[2], sections[3]), language)[,3],
                     create_translation_table(substr(text, sections[3], sections[4]), language)[,3],
                     create_translation_table(substr(text, sections[4], sections[5]), language)[,3],
                     create_translation_table(substr(text, sections[5], sections[6]), language)[,3],
                     create_translation_table(substr(text, sections[6], sections[7]), language)[,3],
                     create_translation_table(substr(text, sections[7], sections[8]), language)[,3],
                     create_translation_table(substr(text, sections[8], sections[9]), language)[,3],
                     create_translation_table(substr(text, sections[9], sections[10]), language)[,3],
                     create_translation_table(substr(text, sections[10], sections[11]), language)[,3],
                     sep=" ")
        text<-text2
        # return(rep("FIȘA SDS NU ESTE ÎN ROMÂNĂ SAU NU A PUTUT FI CITITĂ CORECT. TABELUL NECESITĂ COMPLETARE MANUALĂ", 9))
      }
      
      text<-str_flatten(gsub("[\r\n]", " ", text))
      text<-gsub("Ț", "Ţ", text)
      text<-gsub("ș","ş", text)
      ##Cleanup
      if(grepl("Roth",text, ignore.case = TRUE)){
        
        for (ii in 1:(length(Roth))){
          text<-gsub(Roth[ii], "", text, ignore.case =TRUE)}
      }else{
        for (ii in 1:(length(Merck)))
        {text<-gsub(Merck[ii],"", text, ignore.case =TRUE)
        st.pg<-unlist(gregexpr("Pagina", text, ignore.case = TRUE))
        en.pg<-unlist(gregexpr("Canada", text, ignore.case = TRUE))
        # length(en.pg)
        if(length(st.pg)>1 && length(en.pg)>1){
          for(ij in length(st.pg):1){
            asub<-substr(text, st.pg[ij], en.pg[ij]+5)
            text<-gsub(asub,"", text)
          }
        }
        }
      }
      
      sections <- unlist(gregexpr('SECŢIUNEA', text))
      
      if (length(sections)>1){
        cas_section <- gsub("\\b\\d{3}-\\d{3}-\\d{2}-\\d{1}\\b","", substr(text, unlist(gregexpr('CAS', text))[1], unlist(gregexpr('SECŢIUNEA 2', text))[1]-1))
        cas_numbers <- if(length(grep("CAS", cas_section))>0){ 
          grep("\\b\\d{3,4}-\\d{3,4}-\\d{2}-\\d{1}\\b", cas_section)
          str_extract_all(cas_section, "\\b\\d{2,7}-\\d{2}-\\d{1}\\b")[[1]]
        }else{cas_numbers<-c("Acest produs este un amestec")}
        
        hazard_section <-substr(text, unlist(gregexpr('SECŢIUNEA 2', text))[1], unlist(gregexpr('SECŢIUNEA 3', text))[1]-1)
        
        codes <- NULL
        classif <- NULL
        for (g.i in 1:nrow(GHS)) {
          if (length(grep(GHS[g.i, 1], hazard_section, ignore.case = TRUE))>0) {
            codes <- c(codes, paste(GHS[g.i, 1], GHS[g.i, 2]))
            classif <- c(classif, paste(GHS[g.i, 3]))
          }
        }
        codes <- if (is.null(codes)) "N/A" else str_flatten(unique(codes), collapse = " ")
        classif <- if (is.null(classif)) "Anodin" else str_flatten(unique(classif), collapse = " ")
      
        #Mediu
        codes.m <- NULL
        classif.m <- NULL
        for (g.i in 1:nrow(GHS2)) {
          if (length(grep(GHS2[g.i, 1], hazard_section, ignore.case = TRUE))>0) {
            codes.m <- c(codes.m, paste(GHS2[g.i, 1], GHS2[g.i, 2]))
            classif.m <- c(classif.m, paste(GHS2[g.i, 3]))
          }
        }
        
        codes.m <- if (is.null(codes.m)) "N/A" else str_flatten(unique(codes.m), collapse = " ")
        classif.m<-if (is.null(classif.m)) "" else str_flatten(unique(classif.m), collapse = " ")
        
        class2<-paste(classif, classif.m, sep=" ")
        
          
        first_aid<-substr(text, unlist(gregexpr('SECŢIUNEA 4', text))[1], unlist(gregexpr('SECŢIUNEA 5', text))[1]-1)
        first_aid<-substr(first_aid, gregexpr('4.1', first_aid)[[1]][1]+3, gregexpr('4.2', first_aid)[[1]][1]-1)
        
        
        foc<- substr(text, unlist(gregexpr('SECŢIUNEA 5', text))[1], unlist(gregexpr('SECŢIUNEA 6', text))[1]-1)
        extinguishing_media<-substr(foc, gregexpr('5.1', foc)[[1]][1]+3, gregexpr('5.3', foc)[[1]][1]-1)
        advice_for_firefighters <- substr(foc, gregexpr('5.3', foc)[[1]][1]+3, ifelse(grepl('5.4', foc), gregexpr('5.4', foc)[[1]][1]-1, nchar(foc)))
        
        
        decon<-substr(text, unlist(gregexpr('SECŢIUNEA 6', text))[1], unlist(gregexpr('SECŢIUNEA 7', text))[1])
        decontamination <-ifelse(grepl('6.3',decon), substr(decon, gregexpr('6.3', decon)[[1]][1]+3, gregexpr('6.4', decon)[[1]][1]-1), decon)
        
        ppe_section <- substr(text, unlist(gregexpr('SECŢIUNEA 8', text))[1], unlist(gregexpr('SECŢIUNEA 9', text))[1])
        ppe <- NULL
        if (grepl("goggles|ochelari|Ochelari|ochilor", ppe_section, ignore.case = TRUE)) ppe <- c(ppe, "Ochelari de Protecție")
        if (grepl("gloves|mănuși|mănuşi.|Mănuşile ", ppe_section, ignore.case = TRUE)) ppe <- c(ppe, "Mănuși")
        if (grepl("respiratory protection should|respiratorie", ppe_section, ignore.case = TRUE)) ppe <- c(ppe, "Protecția resiprației în caz de generare de aerosoli/pulberi/vapori")
        if (grepl("clothing|îmbrăcăminte|Îmbrăcăminte", ppe_section, ignore.case = TRUE)) ppe <- c(ppe, "Îmbrăcăminte de protecție")
        if (grepl("antistatic|antistatică", ppe_section, ignore.case = TRUE)) ppe <- c(ppe, " antistatică şi ignifugă")
        ppe <- str_flatten(ppe, collapse = " și ")
        
        physical_properties <- substr(text, unlist(gregexpr('SECŢIUNEA 9', text))[1], unlist(gregexpr('SECŢIUNEA 10', text))[1])
        sec1<-gregexpr(ifelse(grepl('9.1', physical_properties),'9.1', 1), physical_properties)[[1]][1]+3
        sec2<-gregexpr(ifelse(grepl('9.2', physical_properties),'9.2', nchar(physical_properties)), physical_properties)[[1]][1]-1
        physical_properties <- substr(physical_properties, sec1, sec2)
        
        
        return(c(cas_numbers, physical_properties, class2, codes, codes.m, ppe, advice_for_firefighters, first_aid, extinguishing_media, decontamination))
      }else{return(rep("FIȘA SDS NU ESTE ÎN ROMÂNĂ SAU NU A PUTUT FI CITITĂ CORECT. TABELUL NECESITĂ COMPLETARE MANUALĂ", 9))}
    }

    az<-str_trunc(gsub(":","-",Sys.time()), 19, ellipsis="")
    log.tm<-paste("Error-Log",az,".txt", sep="")
    cat("",file=log.tm)

    sink(file=log.tm, append=TRUE)
    ## Process all files and build the result table
    for (i in 1:length(files)) {
      shinyCatch({result <- process_file(files[i], GHS)
      if (!is.null(result)) {
        for (t in 1:10){
          tabel[i, t+1] <- ifelse(nchar(result[t])<32767,result[t],str_trunc(result[t], 32760))}
      } else {
        tabel[i, 2] <- c("Error processing file", 10)
      }
      
      # Update Progress Bar
      progress$inc(1 / length(files), detail = paste(":", i, "din", length(files)))
    })
    }
    
    # Save the result table to an Excel file
    # detach("package:readxl", unload=TRUE)
    # output_file <- file.path(Folder_Salvare, paste0(input$filename, ".xlsx"))
    sink()
    write.xlsx(as.data.frame(tabel), file = paste0(input$filename, ".xlsx"), asTable=TRUE)


    # progress$close()
  })
  session$onSessionEnded(function() {
    stopApp()
  })
}

# Run the Shiny App
shinyApp(ui = ui, server = server)
