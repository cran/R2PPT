"PPT.AddTitleOnlySlide" <-
function(ppt,title=NULL,title.fontsize=NULL,title.font=NULL){

if(!comIsValidHandle(ppt$ppt))   stop("Invalid handle for powerpoint application")
if(!comIsValidHandle(ppt$pres))  stop("Invalid handle for powerpoint presentation")

ppt$Current.Slide <- comInvoke(comGetProperty(ppt$pres,"Slides"),"Add",comGetProperty(comGetProperty(ppt$pres,'Slides'),'Count')+1,11)
comInvoke(ppt$Current.Slide,'Select')

mainseg<-comGetProperty(comGetProperty(comGetProperty(comGetProperty(ppt$Current.Slide,"Shapes"),"Title"),"TextFrame"),"TextRange")

if(!is.null(title))          comSetProperty(mainseg,"Text",title) 
if(!is.null(title.fontsize)) comSetProperty(comGetProperty(mainseg,"Font"),"Size",as.numeric(title.fontsize))
if(!is.null(title.font))     comSetProperty(comGetProperty(mainseg,"Font"),"Name",as.character(title.font))



return(invisible(ppt))

}

