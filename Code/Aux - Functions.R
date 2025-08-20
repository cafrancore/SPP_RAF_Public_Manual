##Profiling functions

findingX=function(data, distMAX,discCut=3){
  goodnessFit=function(clusterCut){
    if(min(table(clusterCut))<basketLimit){
      return(-2)
    }else{
      
      #Calculate Silhouette
      dis = as.dist(distMAX)
      res = clusterCut
      #res= kmeans(data,5)$cluster
      # sil = silhouette (res, dis)
      # fitValue=mean(sil[,3])
       fitValue=cluster.stats(d = dis, res)$dunn
      return(fitValue)
    }
  }
  #X Variables
  result=NULL
  for(i in 1:dim(data)[2]){
    si = Sys.time()
    print(paste0("Checking X",i))
    eval(parse(text=paste0("
    if(is.numeric(data$X",i,")){
    
    
optimizerFunction=function(numberChosen){
  clusterCut=data%>%
    mutate(clusterCut=if_else(X",i,">numberChosen,1,0))
  clusterCut=clusterCut$clusterCut
  if(length(unique(clusterCut))==1){
    y=-2
  }else{
    y=-goodnessFit(clusterCut)
  }
  return(y)
}
if(min(data$X",i,")<max(data$X",i,")){

    temp=optimize(optimizerFunction, c(min(data$X",i,"),max(data$X",i,")))
    result=rbind(result,cbind('X",i,"',temp$minimum,-temp$objective))
}else{
result=rbind(result,cbind('X",i,"',min(data$X",i,"),-2))
}
    }else{
    
selectionVector=unique(data$X",i,")
n <- length(selectionVector)
if(n<2){
result=rbind(result,cbind('X",i,"','No Condition',-2))
}else{
l <- rep(list(0:1), n)
optionMatrix=expand.grid(l)
optionMatrix=optionMatrix[-c(1,dim(optionMatrix)[1]),]

powerMat=t(matrix(rep(2^(0:(dim(optionMatrix)[2]-1)), dim(optionMatrix)[1]),dim(optionMatrix)[2],dim(optionMatrix)[1]))
test1=rowSums(optionMatrix*powerMat)
optionMatrix2=1-optionMatrix
test2=rowSums(optionMatrix2*powerMat)
duplicateLine=apply(cbind(test1,test2), c(1), min)
test1=rowSums(optionMatrix)
optionMatrix2=1-optionMatrix
test2=rowSums(optionMatrix2)
sizeLine=apply(cbind(test1,test2), c(1), min)

optionMatrix=optionMatrix%>%mutate(dupFinder=duplicateLine, sizeCat=sizeLine)%>%
  filter(!duplicated(dupFinder)&sizeCat<discCut )%>%dplyr::select(-c('dupFinder','sizeCat'))

optim=-2
condition='None'
for(j in 1:dim(optionMatrix)[1]){
#print(paste0('Observation ',j,' out of ',dim(optionMatrix)[1]))
  selectVariables=selectionVector[as.vector(optionMatrix[j,])==1]
  clusterCut=data%>%
    mutate(clusterCut=if_else(X",i,"%in%selectVariables,1,0))
   clusterCut=clusterCut$clusterCut
  y=goodnessFit(clusterCut)
  if(y>=optim){
    optim=y
    condition=paste0('c(XXTX',paste(selectVariables, collapse = 'XXTX, XXTX'),'XXTX)')
 
  }
}
result=rbind(result,cbind('X",i,"',condition,optim))
}
}
")))
    ei = Sys.time()
    print(ei - si)
  }

result[,2]=gsub("XXTX", "'", result[,2]) 
result=data.frame(Var=result[,1],cutPoint=as.character(result[,2]),optimal=as.numeric(result[,3]))
result=result%>%arrange(desc(optimal))
return(result[1,])
}



#Profiling functions
treeGraph=function(result, cutsIteration, counter){
  reviewResult=1
  pMatrix=matrix(rep(1:(counter+1),counter+1),counter+1, counter+1)
  for(i in 2:(counter)){
    modifValue=result[counter+2-i,4]
    pMatrix[pMatrix[,i-1]>modifValue,i]=pMatrix[pMatrix[,i-1]>modifValue,i-1]-1
    pMatrix[,i+1]=pMatrix[,i]
    if(i==counter){
      pMatrix[,i+1]=0
    }
  }
  pMatrix=matrix(paste0(pMatrix,"s"),counter+1, counter+1)
  for(i in (counter+1):1){
    pMatrix[,i]=paste0(pMatrix[,i],counter+1-i)
  }
  linkMatrix=NULL
  for(i in (counter+1):2){
    linkMatrix=rbind(linkMatrix,cbind(pMatrix[,i],pMatrix[,i-1]))
  }
  
  linkMatrix=as.data.frame(linkMatrix)%>% distinct()
  net <- graph_from_data_frame(d=linkMatrix, directed=T) 
  lo=t(matrix(as.numeric(unlist(str_split(V(net)$name, "s"))), nrow = 2))
  p=plot(net, layout=lo, vertex.size=0, edge.arrow.size=0, vertex.label=NA) 
  return(list(net=net, lo=lo))
}

#Profiling functions
clusteringFunction=function(data, distanceData, basketLimit, PolicyGroups){
  originalData=data
  
  colnames(data)=paste0("X",1:dim(data)[2])
  colnames(distanceData)=paste0("X",1:dim(distanceData)[2])
  ## Clean Data
  split <- splitmix(distanceData)  
  colnames(split$X.quanti)=gsub("\\.", " ", colnames(split$X.quanti))
  
  if(is.null(colnames(split$X.quali))){
    cleanData=distanceData
  }else{
  colnames(split$X.quali)=gsub("\\.", " ", colnames(split$X.quali))
  cleanData=distanceData%>% 
    dummy_cols(select_columns = colnames(split$X.quali), 
               remove_first_dummy = T,
               remove_selected_columns = T)
  }
  # res.pcamix <- PCAmix(X.quanti=split$X.quanti,  
  #                      X.quali=split$X.quali, 
  #                      rename.level=TRUE, 
  #                      graph=FALSE,
  #                      ndim=dim(data)[2])
  cleanData=as.data.frame(cleanData)
  colnames(cleanData)=paste0("X",1:dim(cleanData)[2])
  cleanData=scale(cleanData)
  cleanData=as.data.frame(cleanData)
  #data=cleanData
  gc()
  #Start Clustering
  #print("Start calculating distance")
  distMAX=as.matrix(dist(cleanData)^2)
  #print("Finished calculating distance")
  result=findingX(data, distMAX)
  result$Guidance=0
  counter=0
  eval(parse(text=paste0("
      if(is.numeric(data$",result[1,1],")){
temp=data%>%
  mutate(clusterCut=10^counter*if_else(",result[1,1],">",result[1,2],",2,1))
      }else{
  temp=data%>%
  mutate(clusterCut=10^counter*if_else(",result[1,1],"%in%",result[1,2],",2,1))    
      }
temp=temp$clusterCut
")))
  cutsIteration=temp
  iterationStep=function(cutsIteration,counter,distMAX){
    gc()
    vecList=list(NULL)
    sVector=NULL
    sResult=list(NULL)
    miniCounter=1
    track=FALSE
    for(i in unique(cutsIteration)){
      miniData=data[cutsIteration==i,]
      distMini=distMAX[cutsIteration==i,cutsIteration==i]
      if(dim(miniData)[1]>4*basketLimit){
        track=TRUE
        result=findingX(miniData,distMini)
        result$Guidance=i
        sResult[[miniCounter]]=result
        eval(parse(text=paste0("
        if(is.numeric(data$",result[1,1],")){
        temp=data%>%mutate(clusterCut = case_when(
                       cutsIteration!=",i," ~ 0,
                       cutsIteration==",i,"&",result[1,1],">",result[1,2]," ~ 0.5,
                       cutsIteration==",i,"&",result[1,1],"<=",result[1,2]," ~ -0.5))
        }else{
        temp=data%>%mutate(clusterCut = case_when(
                       cutsIteration!=",i," ~ 0,
                       cutsIteration==",i,"&",result[1,1],"%in%",result[1,2]," ~ 0.5,
                       cutsIteration==",i,"&!",result[1,1],"%in%",result[1,2]," ~ -0.5))
        }
                       temp=temp$clusterCut"
                               
        )))
        cutsIterationT=cutsIteration+temp
        cutsIterationT=as.numeric(factor(cutsIterationT))
        #table(cutsIterationT)
        
        goodnessFit=function(clusterCut){
          if(min(table(clusterCut))<basketLimit){
            return(-2)
          }else{
            
            #Calculate Silhouette
            dis = distMAX
            res = clusterCut
            #res= kmeans(data,5)$cluster
            # sil = silhouette (res, dis)
            # fitValue=mean(sil[,3])
             fitValue=cluster.stats(d = dis, res)$dunn
            return(fitValue)
          }
        }
        s=goodnessFit(as.numeric(factor(cutsIterationT)))
        vecList[[miniCounter]]=cutsIterationT
        sVector=c(sVector,s)
        miniCounter=miniCounter+1
      }
    }
    if(track){
      cutsIteration=vecList[[which.max(sVector)[1]]]
      result=sResult[[which.max(sVector)[1]]]
      result$optimal=max(sVector)[1]
      return(list(cutsIteration=cutsIteration, result=result))
    }else{
      return(list(cutsIteration=NULL, result=NULL))
    }
  }
  
  Track=TRUE
  while(Track){
    counter=counter+1
    temp=iterationStep(cutsIteration,counter, distMAX)
    if(is.null(temp$result)){
      Track=FALSE
    }else{
      cutsIteration=temp$cutsIteration
      result=rbind(result,temp$result)
      print(result)
      print(table(temp$cutsIteration))
    }
  }
  #print(table(cutsIteration))
  tG=treeGraph(result, cutsIteration, counter)
  
  #Reduce
  dataN=cleanData
  dataN$Clusters=cutsIteration
  summarizedData=dataN%>%
    group_by(Clusters)%>%
    summarize_all(mean)
  
  clust=agnes(summarizedData%>%dplyr::select(-c("Clusters")), method = "ward")
  pltree(clust)
  
  PolicyGroups=min(PolicyGroups,length(unique(summarizedData$Clusters))) # policy groups will give the max available option
  #rect.hclust(clust, k = PolicyGroups, border = 2:(PolicyGroups))
  LargeClusters <- cutree(clust, k = PolicyGroups)
  
  #Re-do
  dataN$Clusters=factor(cutsIteration, levels=summarizedData$Clusters, labels=LargeClusters)
  summarizedDataN=dataN%>%
    group_by(Clusters)%>%
    summarize_all(mean)
  
  originalData$Clusters=factor(cutsIteration, levels=summarizedData$Clusters, labels=LargeClusters)
  
  control <- rpart.control(maxdepth = PolicyGroups*2,
                           cp = 0.000000000000000000001)
  decision = rpart(Clusters ~., data = originalData, method = "class", control=control)
  newClusters=predict(decision, originalData, type = 'class')
  originalData$Clusters=newClusters
  control <- rpart.control(minbucket =  2,
                           cp = 0.000000000000000000001)
  #decision = rpart(Clusters ~., data = originalData, method = "class", control=control)
  rpart.plot(decision)
  table(newClusters)
  
  resultLIst=list(disaggregatedCluster=cutsIteration,aggregatedCluster=dataN$Clusters, editedClusters=newClusters, initialClusteringResult=result, initialTreeConfig=tG, dataSummary1=summarizedData, dataSummary2=summarizedDataN, clusterCluster=clust, decisionTree=decision)
  return(resultLIst)
  
}



#pending chapter
##### density overlap
commonArea=function(scores, beneficiaries){
  # case Yes
  a <- scores[beneficiaries%in%c("Yes")]
  #case No
  b <- scores[beneficiaries%in%c("No")]
  # define limits of a common grid, adding a buffer so that tails aren't cut off
  lower <- min(c(a, b)) - 1 
  upper <- max(c(a, b)) + 1
  # generate kernel densities
  da <- density(a, from=lower, to=upper)
  db <- density(b, from=lower, to=upper)
  d <- data.frame(x=da$x, a=da$y, b=db$y)
  # calculate intersection densities
  d$w <- pmin(d$a, d$b)
  # integrate areas under curves
  total <- integrate.xy(d$x, d$a) + integrate.xy(d$x, d$b)
  intersection <- integrate.xy(d$x, d$w)
  # compute overlap coefficient
  overlap <- 2 * intersection / total
  return(overlap)
}



