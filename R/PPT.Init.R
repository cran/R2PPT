"PPT.Init" <-
function(visible=TRUE){

    if(!require(rcom)){stop("library rcom unavailable")}
    ppt<-list()
    ppt$ppt <- comCreateObject("PowerPoint.Application")
    ppt$pres <- comInvoke(comGetProperty(ppt$ppt,"Presentations"),"Add")
    if(visible){comSetProperty(ppt$ppt,"Visible",TRUE)}
    return(invisible(ppt))
    
}

