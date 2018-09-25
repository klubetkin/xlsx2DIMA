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

  #If it is a line based method, fix a few other field names
  if(method %in% c("LPI", "Gap")){
    #Read DIMA template
    DIMA.template<-readxl::read_excel(template.file, col_types = "text")

    #Template names
    template.names<-names(DIMA.template)

    #Update the LPI Template Names (e.g., remove spaces)
    names(DIMA.template)<-names(DIMA.template) %>% gsub(pattern=" ", replacement = "")

    #Fix Date field so it is R friendly
    DIMA.template<-dplyr::rename(DIMA.template, "Date"="Date(mm/dd/yyyy)")
    DIMA.template<-dplyr::rename(DIMA.template, "LineLength"="LineLength(m)", "InterceptInterval"="InterceptInterval(m)")

    #Join to data
    data.formatted<-dplyr::full_join(DIMA.template, data)

    #Subset to fields in DIMA template
    data.formatted<-data.formatted[, colnames(data.formatted) %in% colnames(DIMA.template)]

    #Add the Excel formatted names back in
    names(data.formatted)<-template.names

    openxlsx::write.xlsx(x = data.formatted, file = out.file)

    }
  #If this is the method
  if(method %in% "Richness"){
    ###Build DIMA Header Table
    DIMA.template.header<-readxl::read_excel(template.file, col_types = "text", sheet = "SpecRich SubPlots")
    #Template names
    template.names.header<-names(DIMA.template.header)

    #Update the  Template Names (e.g., remove spaces)
    names(DIMA.template.header)<-names(DIMA.template.header) %>% gsub(pattern=" ", replacement = "")

     #Join to data
    data.formatted.header<-dplyr::full_join(DIMA.template.header, data,
                                            by=c(`SubPlot#`="SubPlot.", colnames(data)[colnames(data) %in% colnames(DIMA.template.header)]))

    #Subset to fields in DIMA template
    data.formatted.header<-data.formatted.header[,colnames(data.formatted.header) %in% colnames(DIMA.template.header)]%>% unique()

    #Add the Excel formatted names back in
    names(data.formatted.header)<-template.names.header

    ###Build DIMA Detail Tables
    DIMA.template.detail.species<-readxl::read_excel(template.file, col_types = "text", sheet = "SubPlot Species")
    DIMA.template.detail.abundance<-readxl::read_excel(template.file, col_types = "text", sheet = "Species Abundance")
    
    #Template names
    template.names.detail.species<-names(DIMA.template.detail.species)
    template.names.detail.abundance<-names(DIMA.template.detail.abundance)

    #Update the  Template Names (e.g., remove spaces)
    names(DIMA.template.detail.species)<-names(DIMA.template.detail.species) %>% gsub(pattern=" ", replacement = "")
    names(DIMA.template.detail.abundance)<-names(DIMA.template.detail.abundance) %>% gsub(pattern=" ", replacement = "")

    #Join to data
    data.formatted.detail.species<-dplyr::full_join(DIMA.template.detail.species, data,
                                            by=c(`SubPlot#`="SubPlot.", colnames(data)[colnames(data) %in% colnames(DIMA.template.detail.species)]))
    data.formatted.detail.abundance<-dplyr::full_join(DIMA.template.detail.abundance, data,
                                            by=c(`SubPlot#`="SubPlot.", colnames(data)[colnames(data) %in% colnames(DIMA.template.detail.abundance)]))

    #Subset to fields in DIMA template
    data.formatted.detail.species<-data.formatted.detail.species[, colnames(data.formatted.detail.species) %in% colnames(DIMA.template.detail.species)]
    data.formatted.detail.abundance<-data.formatted.detail.abundance[, colnames(data.formatted.detail.abundance) %in% colnames(DIMA.template.detail.abundance)]

    #Add the Excel formatted names back in
    names(data.formatted.detail.species)<-template.names.detail.species
    names(data.formatted.detail.abundance)<-template.names.detail.abundance

    #Create list
    data.formatted<-list("SpecRich SubPlots"=data.formatted.header, "SubPlot Species"=data.formatted.detail.species, "SpeciesAbundance"=data.formatted.detail.abundance)

    #Write out detail table
    openxlsx::write.xlsx(x = data.formatted, file=out.file)

  }



}

#'@export
#'@rdname dima.template
