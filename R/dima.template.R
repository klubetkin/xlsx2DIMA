#'Build DIMA Template Files
#'@description Build a file in the format needed for DIMA
#'@param template.file An Excel of a blank DIMA template
#'@param method The file template method. Options are "LPI", "Richness", "Gap".
#'@param data The data output from xlsx2R.
#'@param  out.file The file path for the output
#'@name dima.template
#'@return A \code {tbl} of files formated for a DIMA template.

#'@export
#'@rdname dima.template

#Read DIMA.template file
build.DIMA.template<-function(data, template.file, method, out.file){

  #If it is a line based method, or plot history
  if(method %in% c("LPI", "Gap", "History")){
    #Read DIMA template
    DIMA.template<-readxl::read_excel(template.file, col_types = "text")

    #Template names
    template.names<-names(DIMA.template)

    #Update the LPI Template Names (e.g., remove spaces)
    names(DIMA.template)<-names(DIMA.template) %>% gsub(pattern=" ", replacement = "")

    #Fix fields with parentheses so they are R friendly
    if(method %in% c("LPI", "Gap")){
      DIMA.template<-dplyr::rename(DIMA.template, "Date"="Date(mm/dd/yyyy)")
      DIMA.template<-dplyr::rename(DIMA.template, "LineLength"="LineLength(m)", "InterceptInterval"="InterceptInterval(m)")
    }
    
    #Join to data
    data.formatted<-dplyr::full_join(DIMA.template, data)

    #Subset to fields in DIMA template
    data.formatted<-data.formatted[, colnames(data.formatted) %in% colnames(DIMA.template)]

    #Add the Excel formatted names back in
    names(data.formatted)<-template.names

    #Write out detail table
    openxlsx::write.xlsx(x = data.formatted, file = out.file)

  }
  #If it is the species richness method
  if(method %in% "Richness"){
    ###Read DIMA template
    DIMA.template.header<-readxl::read_excel(template.file, col_types = "text", sheet = "SpecRich SubPlots")
    DIMA.template.detail.species<-readxl::read_excel(template.file, col_types = "text", sheet = "SubPlot Species")
    DIMA.template.detail.abundance<-readxl::read_excel(template.file, col_types = "text", sheet = "Species Abundance")
    
    #Template names
    template.names.header<-names(DIMA.template.header)
    template.names.detail.species<-names(DIMA.template.detail.species)
    template.names.detail.abundance<-names(DIMA.template.detail.abundance)

    #Update the Template Names (e.g., remove spaces)
    names(DIMA.template.header)<-names(DIMA.template.header) %>% gsub(pattern=" ", replacement = "")
    names(DIMA.template.detail.species)<-names(DIMA.template.detail.species) %>% gsub(pattern=" ", replacement = "")
    names(DIMA.template.detail.abundance)<-names(DIMA.template.detail.abundance) %>% gsub(pattern=" ", replacement = "")

     #Join to data
    data.formatted.header<-dplyr::full_join(DIMA.template.header, data,
                                            by=c(`SubPlot#`="SubPlot.", colnames(data)[colnames(data) %in% colnames(DIMA.template.header)]))
    data.formatted.detail.species<-dplyr::full_join(DIMA.template.detail.species, data,
                                            by=c(`SubPlot#`="SubPlot.", colnames(data)[colnames(data) %in% colnames(DIMA.template.detail.species)]))
    data.formatted.detail.abundance<-dplyr::full_join(DIMA.template.detail.abundance, data)

    #Subset to fields in DIMA template
    data.formatted.header<-data.formatted.header[,colnames(data.formatted.header) %in% colnames(DIMA.template.header)]
    data.formatted.detail.species<-data.formatted.detail.species[, colnames(data.formatted.detail.species) %in% colnames(DIMA.template.detail.species)]
    data.formatted.detail.abundance<-data.formatted.detail.abundance[, colnames(data.formatted.detail.abundance) %in% colnames(DIMA.template.detail.abundance)]

    #Add the Excel formatted names back in
    names(data.formatted.header)<-template.names.header
    names(data.formatted.detail.species)<-template.names.detail.species
    names(data.formatted.detail.abundance)<-template.names.detail.abundance
    
    #Create list
    data.formatted<-list("SpecRich SubPlots"=data.formatted.header, "SubPlot Species"=data.formatted.detail.species, 
                         "Species Abundance"=data.formatted.detail.abundance)

    #Write out detail table
    openxlsx::write.xlsx(x = data.formatted, file=out.file)

  }
  #If it is the plot definition method
  if(method %in% "Definition"){
    ###Read DIMA template
    DIMA.template.general<-readxl::read_excel(template.file, col_types = "text", sheet = "General")
    DIMA.template.lines<-readxl::read_excel(template.file, col_types = "text", sheet = "Lines")
    DIMA.template.soilpits<-readxl::read_excel(template.file, col_types = "text", sheet = "Soil Pits")
    DIMA.template.horizons<-readxl::read_excel(template.file, col_types = "text", sheet = "Horizons")
    
    #Template names
    template.names.general<-names(DIMA.template.general)
    template.names.lines<-names(DIMA.template.lines)

    #Update the Template Names (e.g., remove spaces)
    names(DIMA.template.general)<-names(DIMA.template.general) %>% gsub(pattern=" ", replacement = "")
    names(DIMA.template.lines)<-names(DIMA.template.lines) %>% gsub(pattern=" ", replacement = "")

    #Fix fields with parentheses so they are R friendly
    DIMA.template.general<-dplyr::rename(DIMA.template.general, "Lat"="Lat(orNorthing)", "Long"="Long(or Easting)", "Zone"="Zone(ifUTM)")
    DIMA.template.lines<-dplyr::rename(DIMA.template.lines, "StartLatitude"="StartLatitude(orNorthing)", "StartLongitude"="StartLongitude(orEasting)",
                                  "EndLatitude"="EndLatitude(orNorthing)", "EndLongitude"="EndLongitude(orEasting)")     

    #Join to data
    data.formatted.general<-dplyr::full_join(DIMA.template.general, data,
                                            by=c(`SubPlot#`="SubPlot.", colnames(data)[colnames(data) %in% colnames(DIMA.template.general)]))
    data.formatted.lines<-dplyr::full_join(DIMA.template.lines, data,
                                            by=c(`SubPlot#`="SubPlot.", colnames(data)[colnames(data) %in% colnames(DIMA.template.lines)]))

    #Subset to fields in DIMA template
    data.formatted.general<-data.formatted.general[,colnames(data.formatted.general) %in% colnames(DIMA.template.general)]
    data.formatted.lines<-data.formatted.lines[,colnames(data.formatted.lines) %in% colnames(DIMA.template.lines)]
    
    #Add the Excel formatted names back in
    names(data.formatted.general)<-template.names.general
    names(data.formatted.lines)<-template.names.lines
    
    #Create list
    data.formatted<-list("General"=data.formatted.general, "Lines"=data.formatted.lines, "Soil Pits"=DIMA.template.soilpits, "Horizons"=DIMA.template.horizons)

    #Write out detail table
    openxlsx::write.xlsx(x = data.formatted, file=out.file)

  }
  
}

#'@export
#'@rdname dima.template
