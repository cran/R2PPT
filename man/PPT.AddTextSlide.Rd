\name{PPT.AddTextSlide}
\alias{PPT.AddTextSlide}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{ Add a title slide to current PowerPoint presentation}
\description{
Adds a slide to the current PowerPoint presentation of Microsoft type 'ppLayoutText'.
}
\usage{
PPT.AddTextSlide(ppt, title = NULL, title.fontsize = NULL, title.font = NULL, text = NULL, text.fontsize = NULL, text.font = NULL)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
\item{ppt}{Required list of objects of class 'COMobject' as initialised by PPT.Init. See example.}
\item{title}{Optional title text.}
\item{title.fontsize}{Optional numeric title font size.}
\item{title.font}{Optional character title font.}
\item{text}{Optional character text.}
\item{text.fontsize}{Optional numeric text font size.}
\item{text.font}{Optional character text font.}

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

