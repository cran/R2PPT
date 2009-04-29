"PPT.ApplyTemplate" <-
function(ppt,file){

if(!comIsValidHandle(ppt$ppt))   stop("Invalid handle for powerpoint application")
if(!comIsValidHandle(ppt$pres))  stop("Invalid handle for powerpoint presentation")


if(!file.exists(file)) stop(paste(file, "does not exist"))

file<-gsub("/","\\\\",as.character(file))

comInvoke(ppt$pres,"ApplyTemplate",file)

return(invisible(ppt))
}

