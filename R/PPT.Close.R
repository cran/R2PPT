"PPT.Close" <-
function(ppt){

if(!comIsValidHandle(ppt$ppt))   stop("Invalid handle for powerpoint application")
if(!comIsValidHandle(ppt$pres))  stop("Invalid handle for powerpoint presentation")

comInvoke(ppt$pres,"Close")

return(invisible(ppt))
}

