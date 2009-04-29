\name{PPT.Init}
\alias{PPT.Init}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{ Initialises a PowerPoint presentation}
\description{
Initialises a PowerPoint presentation by using rcom.
}
\usage{
PPT.Init(visible=TRUE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{visible}{ Optional logical specifying if PowerPoint is visible from initialisation. Useful for debugging. Defaults to TRUE.}
}
%\details{}
\value{

  \item{ppt}{Invisibly returns an initialised list of objects of class 'COMobject'.}
  
}
\references{ 
See \url{ http://sunsite.univie.ac.at/rcom} for more details on rcom.

}
\author{ Wayne Jones \email{wayne_betws@hotmail.com}}
\section{Warning}{You must save the returned list of ComObjects initialised with PPT.Init to an R object. See the example. }
%\note{}
\seealso{\code{\link{PPT.Init}},\code{\link{PPT.AddTitleSlide}},\code{\link{PPT.AddTextSlide}},\code{\link{PPT.AddTitleOnlySlide}},\code{\link{PPT.Present}},\code{\link{PPT.ApplyTemplate}},\code{\link{PPT.AddGraphicstoSlide}}}


\examples{

\dontrun{
myPres<-PPT.Init(visible=TRUE)
myPres<-PPT.AddBlankSlide(myPres)
myPres<-PPT.AddTitleSlide(myPres,title="Title Slide",subtitle="Subtitle here")
myPres<-PPT.AddTitleOnlySlide(myPres,title="Title Only",title.fontsize=40,title.font="Arial")
myPres<-PPT.AddTextSlide(myPres,title="Text Slide",text="Text1 \rText2 \rText3",text.font="Arial")
rm(myPres)
} 

}
\keyword{IO}
\keyword{connection}
