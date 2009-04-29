\name{PPT.SaveAs}
\alias{PPT.SaveAs}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{ Save a PowerPoint presentation to file}
\description{
Saves a PowerPoint presentation to a specified file using rcom.
}
\usage{
PPT.SaveAs(ppt,file)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
\item{ppt}{Required list of objects of class 'COMobject' as initialised by PPT.Init. See example.}
\item{file}{Required filename giving the full file path to R graphics output. See example.}

}
%\details{}
\value{

  \item{ppt}{Invisibly returns an initialised list of objects of class 'COMobject'.}
  
}
\references{ 
See \url{ http://sunsite.univie.ac.at/rcom} for more details on rcom.

}
\author{ Wayne Jones \email{wayne_betws@hotmail.com}}

\section{Warning}{You must update the list of ComObjects initialised with PPT.Init when you call this function. 
The full filepath must be specified. No overwrite warning is given.
}

%\note{}
\seealso{\code{\link{PPT.Init}},\code{\link{PPT.Close}},\code{\link{PPT.AddTitleSlide}},\code{\link{PPT.AddTextSlide}},\code{\link{PPT.AddTitleOnlySlide}},\code{\link{PPT.Present}},\code{\link{PPT.ApplyTemplate}},\code{\link{PPT.AddGraphicstoSlide}}}


\examples{

\dontrun{
#Example of generating a presentation without making it visible (can be a lot quicker!).
myPres<-PPT.Init(visible=FALSE)
myPres<-PPT.AddBlankSlide(myPres)
myPres<-PPT.AddTitleSlide(myPres,title="Title Slide",subtitle="Subtitle here")
myPres<-PPT.AddTitleOnlySlide(myPres,title="Title Only",title.fontsize=40,title.font="Arial")
myPres<-PPT.AddTextSlide(myPres,title="Text Slide",text="Text1 \rText2 \rText3",text.font="Arial")
myPres<-PPT.SaveAs(myPres,file=paste(getwd(),"test R2PPT Pres.ppt",sep="/"))
myPres<-PPT.Close(myPres)
rm(myPres)

### Open presentation
myPres<-PPT.Open(file=paste(getwd(),"test R2PPT Pres.ppt",sep="/"))
myPres<-PPT.Present(myPres)
} 

}
\keyword{IO}
\keyword{connection}
