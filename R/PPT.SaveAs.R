"PPT.SaveAs" <-
function(ppt,file){

if(!comIsValidHandle(ppt$ppt))   stop("Invalid handle for powerpoint application")
if(!comIsValidHandle(ppt$pres))  stop("Invalid handle for powerpoint presentation")

comInvoke(ppt$pres,"SaveAs",gsub("/","\\\\",as.character(file)))

return(invisible(ppt))
}

