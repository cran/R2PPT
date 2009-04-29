\name{PPT.AddBlankSlide}
\alias{PPT.AddBlankSlide}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{ Add a Blank Slide to current PowerPoint presentation}
\description{
Adds a slide to the current PowerPoint presentation of Microsoft type 'ppLayoutBlank'.
}
\usage{
PPT.AddBlankSlide(ppt)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{ppt}{ List of objects of class 'COMobject' as initialised by PPT.Init.}
}
%\details{}
\value{

  \item{ppt }{Invisibly returns an updated list of objects of class 'COMobject'.}
  
}
\references{ 
See \url{ http://sunsite.univie.ac.at/rcom} for more details on rcom.

}
\author{ Wayne Jones \email{wayne_betws@hotmail.com}}
\section{Warning}{You must update the list of ComObjects initialised with PPT.Init when you call this function. See the example. }
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
