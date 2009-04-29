"PPT.AddTitleSlide" <-
function(ppt,title=NULL,subtitle=NULL,title.font=NULL,title.fontsize=NULL,subtitle.font=NULL,subtitle.fontsize=NULL){

if(!comIsValidHandle(ppt$ppt))   stop("Invalid handle for powerpoint application")
if(!comIsValidHandle(ppt$pres))  stop("Invalid handle for powerpoint presentation")

ppt$Current.Slide <- comInvoke(comGetProperty(ppt$pres,"Slides"),"Add",comGetProperty(comGetProperty(ppt$pres,'Slides'),'Count')+1,1)
comInvoke(ppt$Current.Slide,'Select')


mainseg<-comGetProperty(comGetProperty(comGetProperty(comGetProperty(ppt$Current.Slide,"Shapes"),"Title"),"TextFrame"),"TextRange")

if(!is.null(title))          comSetProperty(mainseg,"Text",title) 
if(!is.null(title.fontsize)) comSetProperty(comGetProperty(mainseg,"Font"),"Size",as.numeric(title.fontsize))
if(!is.null(title.font))     comSetProperty(comGetProperty(mainseg,"Font"),"Name",as.character(title.font))


subseg<-comGetProperty(comGetProperty(comInvoke(comGetProperty(ppt$Current.Slide,"Shapes"),"Item",2),"TextFrame"),"TextRange")

if(!is.null(subtitle))          comSetProperty(subseg,"Text",subtitle)
if(!is.null(subtitle.fontsize)) comSetProperty(comGetProperty(subseg,"Font"),"Size",as.numeric(subtitle.fontsize))
if(!is.null(subtitle.font))     comSetProperty(comGetProperty(subseg,"Font"),"Name",as.character(subtitle.font))


return(invisible(ppt))
}

