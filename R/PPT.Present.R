"PPT.Present" <-
function(ppt){

if(!comIsValidHandle(ppt$ppt))   stop("Invalid handle for powerpoint application")
if(!comIsValidHandle(ppt$pres))  stop("Invalid handle for powerpoint presentation")

if(!comGetProperty(ppt$ppt,"Visible")){comSetProperty(ppt$ppt,"Visible",TRUE)}
comInvoke(comGetProperty(ppt$pres,"SlideShowSettings"),"Run")

return(invisible(ppt))
}

