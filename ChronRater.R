#the first time, you'll need to install some packages - run this:
# install.packages(c("xlsx","gplots"))

###EDIT LINES 7 - 73 to parameterize your ChronRater experience

###INPUT FILE SECTION###
#csv is most reliable
#inFile="ChronRater_input_template.csv"

#for multiple csv files, try
inFile=c("ChronRater_input_template.csv","ChronRater_input_template2.csv")

#or try excel (Note that this is tempermental and doesn't work on all systems)
#inFile=c("ChronRater_input_template.xlsx")

###END INPUT FILE SECTION###

###SETTINGS###
ChronRaterDir="~/ChronRater/"
setwd(ChronRaterDir)
which.sheets="all" #which worksheets in the spreadsheet to use? "all" uses all, otherwise, select numbers (e.g., c(1,2,4))
agerange=c(-200,12000) #age range (in BP) over which ChronRater will operate
clamdir=paste0(ChronRaterDir,"clam/")

csv.export=TRUE #export results to a csv file? CSV is less buggy
csv.outfile="chron_ratings.csv" #

excel.export=FALSE #export results to an excel spreadsheet? CSV is less buggy
excel.outfile="chron_ratings.xlsx" #because the xlsx package is funky, this needs to not include the directory name and will be written in the chron rater directory

plot.pdfs=TRUE #make pdfs of each agemodel?
pdfdir=paste0(ChronRaterDir,"pdfs/")   #directory for agemodel pdfs - you will need a trailing slash here

plot.histogram=FALSE #want to see a histogram of all the scores
use.marine.calibration.curve=FALSE # if FALSE, it uses the standard Intcal13

###Assigned variables and weightings here###
nReps=1; # if you want to run multiple times with multiple weightings, you can do so

#if you want to vary the weightings, set as a vector with multiple entries (e.g., wD <- c(.001,.002))
#but make sure it's not longer than nReps

#Delineation Weighting
wD <- .001
#Frequency weighting
wFreq <- 2
#Regularity weighting
wReg <- .5
#Uniformity weighting
wUnif <- 3
#Spline stiffness (higher is stiffer, must be less than 1, 0.75 default in CLAM)
spline_smooth=.7
#Delineation Weighting
wQ <- 1
#proportion accepted/monotonic
wpAccMon <- 1
#material weighting
wMat <- 1
#analytical precision weighting
wS <- 200


#number of reversals > 100yr
#determine reversals?
det.rev=TRUE
rev.thresh=100 #years (threshold for considering something a reversal (won't consider less than this range))
depth.thresh=5 #cm (threshold for considering something a reversal (won't consider less than this range))

#otherwise enter manually HERE
nRev=0 # how many reversals are there (only used if det.rev=FALSE)


###STOP EDITING HERE######



###begin helper functions - do not edit
csv.chron.list <- function(inFile="ChronRater_input_template.csv"){
sheet.range=1:length(inFile)
  cl=list()
  for(s in sheet.range){
    inData=read.csv(inFile[s],stringsAsFactors=FALSE)
    short.name=as.character((inData$short_name[1]))
    DT=inData[,c(4:10)]
    #how many rows?
    min.na=c()
    for(d in 1:7){
      min.na[d]=min(which(is.na(DT[,d])))
    }
    nr=max(min.na-1)#so take all the rows that have some data
    DT=DT[1:nr,]
    #if top and bottom depths are present, and middle depth isn't, calculate middle depth
    if(all(!is.na(DT[,1:2])) & all(is.na(DT[,3]))){
      DT[,3]=rowMeans(DT[,1:2])
    }
    #assign the right names
    names(DT)=c("depth_top","depth_bot","depth","age14C","age_error","age","age_error")
    
    #assign into chron.list
    temp.list=list()
    temp.list$measurements=DT
    temp.list$units=c("cm","cm","cm","bp","1_s","bp","2_s")
    temp.list$txtdata$age_type=unlist(inData[[3]][1:nr])
    temp.list$txtdata$rejected_dates=unlist(inData[[11]][1:nr])
    temp.list$txtdata$rejected_dates[temp.list$txtdata$rejected_dates==""]=NA
    
    temp.list$material.score=inData[1,2]
    #write into bigger list
    cl[[short.name]]=temp.list
  }
  
  return(cl)
}

xls.chron.list <- function(inFile="ChronRater.xls",sheet.range="all"){
  library("xlsx")
  wb=loadWorkbook(inFile)
  if (grepl("all",sheet.range)){
    rm(sheet.range)
    sheet.range <- 1:length(getSheets(wb))
  }
  cl=list()
  for(s in sheet.range){
    inData=read.xlsx(inFile,s,colIndex=seq(1,12))
    short.name=as.character(na.exclude(inData$short_name))
    DT=inData[,c(4:10)]
    #how many rows?
    min.na=c()
    for(d in 1:7){
      min.na[d]=min(which(is.na(DT[,d])))
    }
    nr=max(min.na-1)#so take all the rows that have some data
    DT=DT[1:nr,]
    #if top and bottom depths are present, and middle depth isn't, calculate middle depth
    if(all(!is.na(DT[,1:2])) & all(is.na(DT[,3]))){
      DT[,3]=rowMeans(DT[,1:2])
    }
    #assign the right names
    names(DT)=c("depth_top","depth_bot","depth","age14C","age_error","age","age_error")
    
    #assign into chron.list
    temp.list=list()
    temp.list$measurements=DT
    temp.list$units=c("cm","cm","cm","bp","1_s","bp","2_s")
    temp.list$txtdata$age_type=unlist(inData[[3]][1:nr])
    temp.list$txtdata$rejected_dates=unlist(inData[[11]][1:nr])
    
    temp.list$material.score=inData[1,2]
    #write into bigger list
    cl[[short.name]]=temp.list
  }
  
  return(cl)
}

rmse.from.spline <- function(age,depth,error=NA,smooth=.75,plotopt=FALSE,pdfopt=plot.pdfs,pdfname="~/ChronRater/pdfs/out.pdf",curvetype="pspline",polyorder=2,polydf=0,title.string=NA,cex.title=1.3,ylims=c(-200,12000)){
#library("pspline")
it=0;
too.many.its=FALSE
if (all(is.na(error))){
  w <- c()
}else if(any(is.na(error))){
  error[which(is.na(error))]=5
 w <- 1/error^2
}else{
  w <- 1/error^2
}

depthseq <- data.frame(use.depth=seq(min(depth),max(depth)))
while(1){
if (curvetype=="loess"){
resids <- age-predict(loess(age ~ depth, weights=w, span=smooth), depth)
}else if(curvetype=="orthpoly"){
  #needs work!
  resids <- age-predict(age ~ poly(depth,2), depth)
}else if(curvetype=="smoothspline"){
 # smth.spl <- smooth.spline(depth,age,
  stop("doesn't work yet")
}else if(curvetype  =="pspline"){
  if(polydf==0){polydf=4}
   if(length(age)<4){
     resids <- age-predict(lm(use.age ~ use.depth))
     chron <- predict(lm(use.age ~ use.depth),depthseq)
   }else{
    chronl <- predict(smooth.spline(use.depth,use.age,df=polydf),depth)
    resids <- age-chronl$y
    chronl <- predict(smooth.spline(use.depth,use.age,df=polydf),depthseq[[1]])
    chron <- chronl$y
 
}
  #check to make sure no reversals
  if(all(chron == cummax(chron)) | too.many.its){
    break
  }else{
    polydf=polydf-.1
    it=it+1;
    if (it>30){
      too.many.its==TRUE
      print("Reached iteration limit")
      break
    }
  }
}
}


rmse <- sqrt(mean(resids^2))


if (plotopt){
  library(gplots)
 x11()
plot.new()
plotCI(depth,age,uiw=error)
lines(depthseq[[1]],chron,col="red")
}

if (pdfopt){
  library(gplots)
 pdf(file=pdfname,height=5,width=5)
plotCI(depth/10,age,uiw=error,yaxs='i',ylim=ylims,cex.axis=1.3,cex=1.3,cex.lab=1.3,xlab="Depth (cm)",ylab="Age (cal yr BP)",cex.main=cex.title,ylims=agerange)
  lines(depthseq[[1]]/10,chron,col="red")
  pti <- max(unlist(gregexpr("[/]",pdfname)))
  pti2 <- max(unlist(gregexpr("[.pdf]",pdfname)))
    plottitlename <- substr(pdfname,pti+1,pti2-4)
  if (is.na(title.string)){
  plottitle <- plottitlename
}else{
    plottitle <- paste0(plottitlename," - ",title.string)
}
  title(plottitle,cex.main=cex.title)
  dev.off()
}

return(rmse)
}



#function to get the median calibrated age
median.cal.age <- function(age14c,age14cerr=50,calcurve=1){
   current.dir=getwd()
   setwd(clamdir)
   if (length(age14cerr)==0){ age14cerr=50}

if (is.na(age14cerr)){
   age14cerr=50
 }
if (age14c>0){
calibrate(cage=age14c,error=age14cerr,cc=calcurve,storedat=TRUE,graph=FALSE)
}else{
  calibrate(cage=age14c,error=age14cerr,postbomb=1,storedat=TRUE,graph=FALSE)
}

   setwd(current.dir)
   cumprob=cumsum(dat$calib[,2])
   medage=dat$calib[which(abs(0.5-cumprob)==min(abs(0.5-cumprob))),1]
   age95=dat$calib[which(abs(0.95-cumprob)==min(abs(0.95-cumprob))),1]
   age05=dat$calib[which(abs(0.05-cumprob)==min(abs(0.05-cumprob))),1]
   age1sp=dat$calib[which(abs(0.683-cumprob)==min(abs(0.683-cumprob))),1]
   age1sn=dat$calib[which(abs(0.317-cumprob)==min(abs(0.317-cumprob))),1]

   calageerr2s=abs(age95-age05)/2
   calageerr1s=abs(age1sp-age1sn)/2
   outage=list(medage,calageerr2s,calageerr1s,age95,age05)
   names(outage)=c("medage","calageerr2s","calageerr1s","age95","age05")
   return(outage)

 }
##end helper functions
#start chron rating

#if inFile is one or more csv files, use csv.chron.list
if(all(grepl(".csv",inFile))){
  chron.list=csv.chron.list(inFile=inFile)
}else if(grepl(".xls",inFile) & length(grepl(".xls",inFile))==1){
chron.list=xls.chron.list(inFile=inFile,sheet.range=which.sheets)
}else{
  stop("inFile options are multiple .csv files, or a single .xls or .xlsx file with (optionally) multiple sheets")
}
#load in clam.R
source(paste0(clamdir,"clam.R"))

#loads in chronology and identifies age and depth
Acc.Score <- matrix(NA,length(chron.list),nReps)
site.name=vector("character",length(chron.list))

#setup parameters to store
Unif=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
def=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
quality=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
Freq=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
Reg=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
pAccMon=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
Mat=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
an.prec=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
mean.calageerror1s=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
mean.calageerror2s=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
mean.error=matrix(data=NA,nrow=length(chron.list),ncol=nReps)
ms2 <- c()

#go through each chronology
for(i in 1:length(chron.list)){
site.chron <- chron.list[[i]]
site.name[i] <- names(chron.list)[i]
site.chron$material.score=chron.list[[i]]$material.score
if (is.na(site.chron$material.score)){
  stop("No material score present")
}

ms2[i] <- site.chron$material.score

#calibrate appropriately
age14C=c()
ageCal=c()
q14c=FALSE
qcal=FALSE
qcalerr=FALSE
ageerror=c()
 C14error1s=c()
 C14error2s=c()

C14i=which(grepl("14C",site.chron$txtdata$age_type))

depth <- site.chron$measurements$depth
if (length(grep("age14C",names(site.chron$measurements)))>0){
   age14C=site.chron$measurements$age14C
   q14c=TRUE
   if(length(grep("age_error",names(site.chron$measurements)))>0){
      age14Cerr=site.chron$measurements$age_error
    }
 }

if (!is.na(match("age",names(site.chron$measurements)))){
   ageCal=site.chron$measurements$age
   qcal=TRUE
      if(length(grep("cal_age_error",names(site.chron$measurements)))>0){
      ageCalerr=site.chron$measurements$cal_age_error
      ageerror2s=ageCalerr
      qcalerr=TRUE
    }
   if(!qcalerr & any(grepl("cal_age_range_old",names(site.chron$measurements))) & any(grepl("cal_age_range_young",names(site.chron$measurements)))){
     ageerror2s=abs(site.chron$measurements$cal_age_range_old-site.chron$measurements$cal_age_range_young)/2
     ageerror1s=ageerror2s/2
     ageCalerr=ageerror2s
     qcalerr=TRUE
   }
 }
no.age.q=FALSE
if (!qcal & !q14c){
 no.age.q=TRUE
}

if (!no.age.q){

#calibrate ages if needed
q.wecal=FALSE
if(q14c){
     if(use.marine.calibration.curve){
       cc=2 #use marine curve
    }else{
      cc=1 #
    }
  calibrated=age14C


  for(iii in which(!is.na(age14C))){

    if((age14C[iii] < 50000) & !(cc == 2 & age14C[iii]<400)){
    outage <- median.cal.age(age14C[iii],age14Cerr[iii],calcurve=cc)
    calibrated[iii]=outage[[1]]
    ageerror[iii]=outage$calageerr1s
    C14error2s[iii]=outage$calageerr2s
    C14error1s[iii]=outage$calageerr1s
  }else{
     calibrated[iii]=NA
    ageerror[iii]=NA
      C14error2s[iii]=NA
    C14error1s[iii]=NA
   }
  }
  q.wecal=TRUE
   }

#setup the age vector, starting with the ages we calibrated, and filling in blanks with ageCal
if (q.wecal){
   age=calibrated
 }else if(qcal){
   age=ageCal
 }

if(qcal & q.wecal & any(is.na(age) & !is.na(ageCal))){
age[is.na(age)] <- ageCal[is.na(age) & !is.na(ageCal)]
}

 #do the same for error
age.error.q=FALSE

if (q.wecal){
   error=ageerror
   age.error.q=TRUE
 }else if(qcalerr){
   error=ageCalerr
   age.error.q=TRUE
    ageerror2s=ageCalerr
    ageerror1s=ageCalerr/2

 }
if(qcalerr){
if(qcalerr & q.wecal & any(is.na(error) & !is.na(ageCalerr))){
error[is.na(error)] <- ageCalerr[is.na(error) & !is.na(ageCalerr)]
age.error.q=TRUE
}}

error=C14error2s

good.index <- which(!is.na(age) & !is.na(depth) & age>agerange[1] & age<=agerange[2])

age <- age[good.index]
depth <- depth[good.index]


if(any(!is.na(error))){
error <- error[good.index]
}


#find rejected ages
nrej=0
rej.q=c()
rej <- NA
if(length(grep("txtdata",names(site.chron))) > 0){
if(length(grep("reject",names(site.chron$txtdata))) > 0){
 rej <- ((site.chron$txtdata$rejected_dates!="NA"))
}}
rej <- rej[good.index]
#and sort by depth
age <- age[order(depth)]
rej <- rej[order(depth)]
error <- error[order(depth)]
depth <- depth[order(depth)] # make sure this one is last

rej.q <- which(rej)
nrej=length(rej.q)
if (nrej==0){
rej.q=NULL
}

for(j in 1:nReps){

#variable quantities:
#for each of the variable quantities, make sure that they are all nReps long
wFreq=rep(wFreq,length.out=nReps)
wReg=rep(wReg,length.out=nReps)
wUnif=rep(wUnif,length.out=nReps)
spline_smooth=rep(spline_smooth,length.out=nReps)
wpAccMon=rep(wpAccMon,length.out=nReps)
wMat=rep(wMat,length.out=nReps)
wS=rep(wS,length.out=nReps)
wD=rep(wD,length.out=nReps)
wQ=rep(wQ,length.out=nReps)

#Maybe Assigned, maybe calculated variables here
#indices for which there are both ages and depths
#total number of ages
nAgesTotal <- length(age)
#number of accepted ages
nAgesAcc <- nAgesTotal-nrej

###REVERSALS###


#identify and calculate age reversals
if(det.rev){
  rev.q=vector("logical",nAgesAcc)
  rev.q[1]=FALSE
  #only look at non-rejected ages
  if(!is.null(rej.q)){
    age.no.rej=age[-rej.q]
    depth.no.rej=depth[-rej.q]
        error.no.rej=error[-rej.q]

  }else{
    age.no.rej=age
    depth.no.rej=depth
    error.no.rej=error
  }
  for(ar in 2:nAgesAcc){
    #reversal if the age difference is negative, and more than the age threshold, and the depth difference is greater than the depth threshold
    if(((age.no.rej[ar]-age.no.rej[ar-1])<(-rev.thresh))&((depth.no.rej[ar]-depth.no.rej[ar-1])>=depth.thresh )){
      rev.q[ar] <- TRUE
    }else{
      rev.q[ar] <- FALSE
    }
  }

 nRev=length(which(rev.q))

}

###END REVERSALS

#calculated variables

use.age=age.no.rej # include reversals, throw out rejections
use.depth=depth.no.rej
use.error=error.no.rej

##Frequency (ages/kyr)
Freq[i,j]=(max(use.age)-min(use.age))/nAgesAcc

##Regularity
age.diff=use.age[-1]-use.age[-length(use.age)]
depth.diff=use.depth[-1]-use.depth[-length(use.depth)]

#deal with mult samples from same depth?
use.in=1:length(age.diff)
depth.thresh <- (max(use.depth)-min(use.depth))/100# one percent
use.in <- which(depth.diff > depth.thresh)
Reg[i,j]=sd(age.diff[use.in])

##Uniformity
Unif[i,j] <- rmse.from.spline(use.age,use.depth,error=use.error,smooth=spline_smooth[j],polydf=4)


##Definition
def[i,j]=wD[j]*(Freq[i,j]*wFreq[j]+Reg[i,j]*wReg[j]+Unif[i,j]*wUnif[j])

##proportion accepted and monotonic
pAccMon[i,j]=1-((nRev+nrej)/nAgesTotal)

##Material type and reference ages
#this is a subjective score that will be included in the site_meta data.

#needs to updated get.chron.R to get this info once its there.
Mat[i,j]=site.chron$material.score

## Quality score
quality[i,j]=wQ[j]*(pAccMon[i,j]*wpAccMon[j]*Mat[i,j]*wMat[j])

## Analytical precision
an.prec[i,j] = wS[j] * mean(error[which(!is.na(error))])^-1
mean.error[i,j]=mean(error,na.rm=TRUE)
mean.calageerror1s[i,j] = mean(C14error1s,na.rm=TRUE)
mean.calageerror2s[i,j] = mean(C14error2s,na.rm=TRUE)

if (!no.age.q){
## and finally
Acc.Score[i,j] <- -def[i,j] + quality[i,j] + an.prec[i,j]
}else{
 Acc.Score[i,j] <- NA
}
roundnum=1
print(paste("Accuracy score for",site.name[i],"=",as.character(Acc.Score[i,j])))

#add into chron.list
chron.list[[i]]$Chron.Score=Acc.Score[i,]

outmat=cbind(Freq,Reg,Unif,pAccMon,Mat,an.prec,Acc.Score,mean.error,mean.calageerror1s,mean.calageerror2s)
dum <- rmse.from.spline(use.age,use.depth,error=use.error,smooth=spline_smooth[j],pdfname=paste0(pdfdir,site.name[i],".pdf"),polydf=4,cex.title=.8,title.string=paste0("AS = ",as.character(round(Acc.Score[i,j],roundnum)),"\n D = ",as.character(round(def[i,j],roundnum)),"; Q  = ",as.character(round(quality[i,j],roundnum)),"; P = ",as.character(round(an.prec[i,j],roundnum)),"\n Freq = ",as.character(round(Freq[i,j],roundnum)),"; Reg = ",as.character(round(Reg[i,j],roundnum)),"\n Unif(RMSE) = ",as.character(round(Unif[i,j],roundnum)),"; pAccMon = ",as.character(round(pAccMon[i,j],roundnum)),"\n Mat = ",as.character(round(Mat[i,j],roundnum))))

}}}
if(plot.histogram){
ncells=21

hist.data.x=matrix(NA,ncells-1,nReps)
hist.data.y=matrix(NA,ncells-1,nReps)

for(j in 1:nReps){
  hist.data=hist(Acc.Score[,j],breaks=seq(min(Acc.Score[!is.na(Acc.Score[,j]),j]),max(Acc.Score[!is.na(Acc.Score[,j]),j]),length.out=ncells),plot=FALSE)
  hist.data.x[,j] <- hist.data$mids
  hist.data.y[,j] <-  hist.data$counts
}
x11()
plot.new()
plot(hist.data.x[,1],hist.data.y[,1],type="l",col="black")
if(nReps>1){
for(j in 2:nReps){
  lines(hist.data.x[,j],hist.data.y[,j],type="l",col=j*5)
}}
}
if(excel.export){
#export to excel
outmat=cbind(Freq,Reg,Unif,def,pAccMon,Mat,quality,an.prec,Acc.Score,mean.error,mean.calageerror1s,mean.calageerror2s)
row.names(outmat)=names(chron.list)
colnames(outmat)=cbind(rep("Frequency",times=nReps),rep("Regularity",times=nReps),rep("Uniformity (RMSE)",times=nReps),rep("D",times=nReps),rep("Proportion",times=nReps),rep("Material",times=nReps),rep("Q",times=nReps),rep("Precision",times=nReps),rep("Total Chronology Score",times=nReps),rep("mean used error",times=nReps),rep("1 sigma calibrated range",times=nReps),rep("2-sigma calibrated range",times=nReps))

voi=wFreq
vois="wFreq"
wordir=getwd()
setwd(ChronRaterDir)
write.xlsx(outmat,excel.outfile,sheetName=vois)
setwd(wordir)
}
if(csv.export){
  #export to csv
  outmat=cbind(Freq,Reg,Unif,def,pAccMon,Mat,quality,an.prec,Acc.Score,mean.error,mean.calageerror1s,mean.calageerror2s)
  row.names(outmat)=names(chron.list)
  colnames(outmat)=cbind(rep("Frequency",times=nReps),rep("Regularity",times=nReps),rep("Uniformity (RMSE)",times=nReps),rep("D",times=nReps),rep("Proportion",times=nReps),rep("Material",times=nReps),rep("Q",times=nReps),rep("Precision",times=nReps),rep("Total Chronology Score",times=nReps),rep("mean used error",times=nReps),rep("1 sigma calibrated range",times=nReps),rep("2-sigma calibrated range",times=nReps))
  
  voi=wFreq
  vois="wFreq"
  wordir=getwd()
  setwd(ChronRaterDir)
  write.csv(outmat,csv.outfile)
  setwd(wordir)
}
