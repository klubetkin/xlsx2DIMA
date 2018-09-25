#'Import from legacy excel files
#'@description Import legacy data from Excel files
#'@param excel.file An Excel of a blank DIMA template
#'@param format Format type, as identified in the reference lookup table
#'@param reference Reference lookup table filepath which identifies excel file formats and corresponding field map
#'@name legacy.format
#'@return A \code {tbl} of all method data in a directory formatted for joining to the DIMA.template.
#'@export
#'@rdname legacy.format
#'
#Make reference tall format
reference.tall<-function(reference){
  #Read reference file
  reference<-read.csv(reference, stringsAsFactors = FALSE)

  #Gather Reference tall
  reference.tall<-reference %>% tidyr::gather(key=format, value=cell.ref, -Table, -FieldName)

  return(reference.tall)
}

#'@export
#'@rdname legacy.format
#Map excel references to a given field name, given a specific reference table
map.excel<-function(excel.file, reference.tall, field, format){
  cell.ref<-reference.tall$cell.ref[reference.tall$FieldName==field &reference.tall$format==format] %>% as.character() %>% strsplit(.,split=",")%>% unlist() %>% gsub("\ ", "", .)

  #Check to see if it is a valid excel reference
  if(!any((grepl(pattern="[[:digit:]]", cell.ref)&grepl(pattern="[[:alpha:]]", cell.ref)))){
    data<-cell.ref
  } else if(any(is.na(cell.ref))){
    data<-NA
  } else if (field%in%c("Date","EstablishDate")){
    data<-tryCatch(readxl::read_excel(path=excel.file,sheet =format,range=cell.ref, col_types = "date", col_names=FALSE)[[1]]%>% as.character, error=function(e) return(NA))
  } else if (substr(field,1,6)=="Chkbox"){
    #Define rows
    rows<-gsub(pattern="[[:alpha:]]", "", x=cell.ref)%>% unique()%>%
      strsplit(., split = ":")%>% unlist()%>% as.numeric
    rows<-seq(min(rows), max(rows))

    #Define cells
    cols<-strsplit(cell.ref, split = ":")%>% unlist() %>% openxlsx::convertFromExcelRef() %>% unique()
    
    #Read in row/column specific data from R
    dataSheet<-xlsx::loadWorkbook(excel.file)%>% getSheets()
    dataSheet<-dataSheet[[format]]
    dataCells<-list()
    dataCells<-lapply(X=cols, FUN=function(X){
      dataCells<-c(dataCells,getRows(dataSheet,rows)%>% getCells(colIndex=X))
    })%>% unlist()
    data<-lapply(X=dataCells,FUN=function(X){
      getCellStyle(X)$getFont()$getBold()
     })
  } else{

    #Read in row/column specific data from R
    data<-tryCatch(lapply(X=cell.ref, FUN=function(X){readxl::read_excel(path=excel.file, sheet=format, range=X, col_names=FALSE)%>% data.frame()}), 
                   warning=function(w) return(NA))
    data<-lapply(X=data,FUN=function(X){
      if(!nrow(X)>0){X<-NA}
      else{X<-X}
    })%>% data.frame()%>% unlist()
  }

  return(data)

}

#'@export
#'@rdname legacy.format

build.header<-function(excel.file, reference.tall, format){
  #Build Header Data Frame
  header.fields<-reference.tall$FieldName[reference.tall$Table=="Header"&reference.tall$format==format]

  #Build Header Data Frame
  header.data<-data.frame(FieldName=header.fields,
                          value=sapply(X=header.fields, FUN=function(f=X){
                            print(f)
                            data<-map.excel(excel.file = excel.file, format = format, reference.tall=reference.tall,field=f)
                            assign(paste(f), data)

                          })) %>%
    #Spread so that the response values are a row
    tidyr::spread(key = FieldName, value=value)%>% dplyr::mutate(excel.file=excel.file)

  #Return Header Data
  return(header.data)

}

#'@export
#'@rdname legacy.format

#Build Detail Table
build.detail<-function(excel.file, reference.tall, format){
  detail.fields<-reference.tall$FieldName[reference.tall$Table=="Detail"&reference.tall$format==format]

  #Build data frame
  detail.data<-lapply(X=detail.fields, FUN=function(f=X){
    print(f)
    df=data.frame(data=map.excel(excel.file = excel.file,  format = format, reference.tall=reference.tall,field=f)%>% as.character(),
                  FieldName=f) %>% dplyr::mutate(id=1:n(), excel.file=excel.file)
    df})%>%
    do.call(rbind, .) %>%
    #Spread so that the response values are a row
    tidyr::spread(key = FieldName, value=data)%>% dplyr::select(-id)
}

#'@export
#'@rdname legacy.format
#'
# join.header.detail<-function(excel.file, reference.tall, format){
#   lapply(X=formats, FUN=function(X){
#   #Gather header and detail file
#   header<-build.header(excel.file=excel.file,  reference.tall=reference.tall, format = X)
#   detail<-build.detail(excel.file=excel.file,  reference.tall=reference.tall, format = X)
#   #Join header and detail
#   dplyr::left_join(header, detail)
#
# })%>% do.call(rbind, .)
# }

#Put it all together for an import function
xlsx2R<-function(folder, reference, out){

  #Build reference tall data frame
  reference.tall<-reference.tall(reference=reference)

  files<-list.files(folder, full.names = TRUE, recursive=TRUE)%>% subset(grepl(pattern = ".xlsx$", x=.)&!grepl(pattern = "~", x=.))

  if(file.exists(out)){
    read.files<-read.csv(out) %>% dplyr::select(excel.file)%>% unique()

    #Remove files that have already been read
    files<-files[!files %in% read.files$excel.file]
  }

  #For each excel file in the folder, pull the relevant data
  lapply(X=files, FUN=function(X){

    format<-openxlsx::getSheetNames(X) %>% subset(. %in% unique(reference.tall$format))

    #If the result of format is a character string, then gather header and detail file
    if(length(format)>0){
      print(X)
      #Gather header and detail file
      header<-build.header(excel.file=X,  reference.tall=reference.tall, format = format)
      detail<-build.detail(excel.file=X,  reference.tall=reference.tall, format = format)
      #Join header and detail
      header.detail<-dplyr::left_join(header, detail) %>% dplyr::mutate(Format=format)

      #save the output to a csv
      if(!file.exists(out)){
        write.table(header.detail, out, append=FALSE, col.names = TRUE, row.names = FALSE, sep=",")
      } else {
        write.table(header.detail, out, append=TRUE, col.names = FALSE, row.names = FALSE, sep=",")
      }
    } else{
      warning(paste("No valid data for import.", paste(X), "will be ignored", sep=" "))
      print(X)
    }})

  #Return the completed file from the compiled csv
  compiled.csv<-read.csv(out, colClasses = "character" ) %>% unique()

  #Fix Shrubshape into a single field (if they exist)

  if("ShrubShape1" %in% colnames(compiled.csv)){
    compiled.csv$ShrubShape<-paste(compiled.csv$ShrubShape1, compiled.csv$ShrubShape2, compiled.csv$ShrubShape3, compiled.csv$ShrubShape4, sep="")%>%
      stringr::str_replace_all("NA|Na", "")
    compiled.csv<-dplyr::select(compiled.csv, -c(ShrubShape1:ShrubShape4))
  }


  #Return
  return(compiled.csv)
}

