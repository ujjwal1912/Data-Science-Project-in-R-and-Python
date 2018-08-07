# Clearing the environment
rm(list=ls(all=T))
# Setting working directory
setwd("C:/Users/Ujjwal/Desktop/Project 1")

# Loading libraries
library(ggplot2)
library(corrgram)
library(DMwR)
library(caret)
library(randomForest)
library(unbalanced)
library(dummies)
library(e1071)
library(Information)
library(MASS)
library(rpart)
library(gbm)
library(ROSE)
library(xlsx)
library(DataCombine)
library(rpart)

## Reading the data
df = read.xlsx('Absenteeism_at_work_Project.xls', sheetIndex = 1)


#----------------------------------------------------Exploratory Data Analysis------------------------------------------------------
# Shape of the data
dim(df)
# Viewing data
# View(df)
# Structure of the data
str(df)
# Variable namesof the data
colnames(df)
# From the above EDA and problem statement categorising data in 2 category "continuous" and "catagorical"
continuous_vars = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
            'Work.load.Average.day.', 'Transportation.expense',
            'Hit.target', 'Weight', 'Height', 
            'Body.mass.index', 'Absenteeism.time.in.hours')

catagorical_vars = c('ID','Reason.for.absence','Month.of.absence','Day.of.the.week',
                     'Seasons','Disciplinary.failure', 'Education', 'Social.drinker',
                     'Social.smoker', 'Son', 'Pet')



#------------------------------------Missing Values Analysis---------------------------------------------------#
#Creating dataframe with missing values present in each variable
missing_val = data.frame(apply(df,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"

#Calculating percentage missing value
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(df)) * 100

# Sorting missing_val in Descending order
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL

# Reordering columns
missing_val = missing_val[,c(2,1)]

# Saving output result into csv file
write.csv(missing_val, "Missing_perc_R.csv", row.names = F)

# # Plot
# ggplot(data = missing_val[1:18,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
# geom_bar(stat = "identity",fill = "grey")+xlab("Variables")+
# ggtitle("Missing data percentage") + theme_bw()

# Actual Value = 23
# Mean = 26.68
# Median = 25
# KNN = 23


#Mean Method
# df$Body.mass.index[is.na(df$Body.mass.index)] = mean(df$Body.mass.index, na.rm = T)

#Median Method
# df$Body.mass.index[is.na(df$Body.mass.index)] = median(df$Body.mass.index, na.rm = T)

# kNN Imputation
df = knnImputation(df, k = 3)

# Checking for missing value
sum(is.na(df))


#-------------------------------------Outlier Analysis-------------------------------------#
# BoxPlots - Distribution and Outlier Check

# Boxplot for continuous variables
for (i in 1:length(continuous_vars))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (continuous_vars[i]), x = "Absenteeism.time.in.hours"), data = subset(df))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=continuous_vars[i],x="Absenteeism.time.in.hours")+
           ggtitle(paste("Box plot of absenteeism for",continuous_vars[i])))
}

# ## Plotting plots together
gridExtra::grid.arrange(gn1,gn2,ncol=2)
gridExtra::grid.arrange(gn3,gn4,ncol=2)
gridExtra::grid.arrange(gn5,gn6,ncol=2)
gridExtra::grid.arrange(gn7,gn8,ncol=2)
gridExtra::grid.arrange(gn9,gn10,ncol=2)


# #Remove outliers using boxplot method

# #loop to remove from all variables
for(i in continuous_vars)
{
  print(i)
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  #print(length(val))
  df = df[which(!df[,i] %in% val),]
}

#Replace all outliers with NA and impute
for(i in continuous_vars)
{
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  #print(length(val))
  df[,i][df[,i] %in% val] = NA
}

# Imputing missing values
df = knnImputation(df,k=3)


#-----------------------------------Feature Selection------------------------------------------#

## Correlation Plot 
corrgram(df[,continuous_vars], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

## ANOVA test for Categprical variable
summary(aov(formula = Absenteeism.time.in.hours~ID,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Reason.for.absence,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Month.of.absence,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Day.of.the.week,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Seasons,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Disciplinary.failure,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Education,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Social.drinker,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Social.smoker,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Son,data = df))
summary(aov(formula = Absenteeism.time.in.hours~Pet,data = df))


## Dimension Reduction
df = subset(df, select = -c(Weight))


#--------------------------------Feature Scaling--------------------------------#
#Normality check
hist(df$Absenteeism.time.in.hours)

# Updating the continuous and catagorical variable
continuous_vars = c('Distance.from.Residence.to.Work', 'Service.time', 'Age',
                    'Work.load.Average.day.', 'Transportation.expense',
                    'Hit.target', 'Height', 
                    'Body.mass.index')

catagorical_vars = c('ID','Reason.for.absence','Disciplinary.failure', 
                     'Social.drinker', 'Son', 'Pet', 'Month.of.absence', 'Day.of.the.week', 'Seasons',
                     'Education', 'Social.smoker')


# Normalization
for(i in continuous_vars)
{
  print(i)
  df[,i] = (df[,i] - min(df[,i]))/(max(df[,i])-min(df[,i]))
}

# Creating dummy variables for categorical variables
library(mlr)
df = dummy.data.frame(df, catagorical_vars)


#------------------------------------------Model Development--------------------------------------------#
#Cleaning the environment
rmExcept("df")

#Divide data into train and test using stratified sampling method
set.seed(123)
train.index = sample(1:nrow(df), 0.8 * nrow(df))
train = df[ train.index,]
test  = df[-train.index,]

##Decision tree for classification
#Develop Model on training data
fit_DT = rpart(Absenteeism.time.in.hours ~., data = train, method = "anova")

#Summary of DT model
summary(fit_DT)

#write rules into disk
write(capture.output(summary(fit_DT)), "Rules.txt")

#Lets predict for training data
pred_DT_train = predict(fit_DT, train[,names(test) != "Absenteeism.time.in.hours"])

#Lets predict for training data
pred_DT_test = predict(fit_DT,test[,names(test) != "Absenteeism.time.in.hours"])


# For training data 
print(postResample(pred = pred_DT_train, obs = train[,107]))

# For testing data 
print(postResample(pred = pred_DT_test, obs = test[,107]))


#------------------------------------------Linear Regression-------------------------------------------#
set.seed(123)

#Develop Model on training data
fit_LR = lm(Absenteeism.time.in.hours ~ ., data = train)

#Lets predict for training data
pred_LR_train = predict(fit_LR, train[,names(test) != "Absenteeism.time.in.hours"])

#Lets predict for testing data
pred_LR_test = predict(fit_LR,test[,names(test) != "Absenteeism.time.in.hours"])

# For training data 
print(postResample(pred = pred_LR_train, obs = train[,107]))

# For testing data 
print(postResample(pred = pred_LR_test, obs = test[,107]))


#-----------------------------------------Random Forest----------------------------------------------#

set.seed(123)

#Develop Model on training data
fit_RF = randomForest(Absenteeism.time.in.hours~., data = train)

#Lets predict for training data
pred_RF_train = predict(fit_RF, train[,names(test) != "Absenteeism.time.in.hours"])

#Lets predict for testing data
pred_RF_test = predict(fit_RF,test[,names(test) != "Absenteeism.time.in.hours"])

# For training data 
print(postResample(pred = pred_RF_train, obs = train[,107]))

# For testing data 
print(postResample(pred = pred_RF_test, obs = test[,107]))


#--------------------------------------------XGBoost-------------------------------------------#

set.seed(123)

#Develop Model on training data
fit_XGB = gbm(Absenteeism.time.in.hours~., data = train, n.trees = 500, interaction.depth = 2)

#Lets predict for training data
pred_XGB_train = predict(fit_XGB, train[,names(test) != "Absenteeism.time.in.hours"], n.trees = 500)

#Lets predict for testing data
pred_XGB_test = predict(fit_XGB,test[,names(test) != "Absenteeism.time.in.hours"], n.trees = 500)

# For training data 
print(postResample(pred = pred_XGB_train, obs = train[,107]))

# For testing data 
print(postResample(pred = pred_XGB_test, obs = test[,107]))



#----------------------Dimensionality Reduction using PCA-------------------------------#


#principal component analysis
prin_comp = prcomp(train)

#compute standard deviation of each principal component
std_dev = prin_comp$sdev

#compute variance
pr_var = std_dev^2

#proportion of variance explained
prop_varex = pr_var/sum(pr_var)

#cumulative scree plot
plot(cumsum(prop_varex), xlab = "Principal Component",
     ylab = "Cumulative Proportion of Variance Explained",
     type = "b")

#add a training set with principal components
train.data = data.frame(Absenteeism.time.in.hours = train$Absenteeism.time.in.hours, prin_comp$x)

# From the above plot selecting 45 components since it explains almost 95+ % data variance
train.data =train.data[,1:45]

#transform test into PCA
test.data = predict(prin_comp, newdata = test)
test.data = as.data.frame(test.data)

#select the first 45 components
test.data=test.data[,1:45]


#------------------------------------Model Development after Dimensionality Reduction--------------------------------------------#

#-------------------------------------------Decision tree for classification-------------------------------------------------#

#Develop Model on training data
fit_DT = rpart(Absenteeism.time.in.hours ~., data = train.data, method = "anova")


#Lets predict for training data
pred_DT_train = predict(fit_DT, train.data)

#Lets predict for training data
pred_DT_test = predict(fit_DT,test.data)


# For training data 
print(postResample(pred = pred_DT_train, obs = train$Absenteeism.time.in.hours))

# For testing data 
print(postResample(pred = pred_DT_test, obs = test$Absenteeism.time.in.hours))



#------------------------------------------Linear Regression-------------------------------------------#

#Develop Model on training data
fit_LR = lm(Absenteeism.time.in.hours ~ ., data = train.data)

#Lets predict for training data
pred_LR_train = predict(fit_LR, train.data)

#Lets predict for testing data
pred_LR_test = predict(fit_LR,test.data)

# For training data 
print(postResample(pred = pred_LR_train, obs = train$Absenteeism.time.in.hours))

# For testing data 
print(postResample(pred = pred_LR_test, obs =test$Absenteeism.time.in.hours))


#-----------------------------------------Random Forest----------------------------------------------#

#Develop Model on training data
fit_RF = randomForest(Absenteeism.time.in.hours~., data = train.data)

#Lets predict for training data
pred_RF_train = predict(fit_RF, train.data)

#Lets predict for testing data
pred_RF_test = predict(fit_RF,test.data)

# For training data 
print(postResample(pred = pred_RF_train, obs = train$Absenteeism.time.in.hours))

# For testing data 
print(postResample(pred = pred_RF_test, obs = test$Absenteeism.time.in.hours))


#--------------------------------------------XGBoost-------------------------------------------#

#Develop Model on training data
fit_XGB = gbm(Absenteeism.time.in.hours~., data = train.data, n.trees = 500, interaction.depth = 2)

#Lets predict for training data
pred_XGB_train = predict(fit_XGB, train.data, n.trees = 500)

#Lets predict for testing data
pred_XGB_test = predict(fit_XGB,test.data, n.trees = 500)

# For training data 
print(postResample(pred = pred_XGB_train, obs = train$Absenteeism.time.in.hours))

# For testing data 
print(postResample(pred = pred_XGB_test, obs = test$Absenteeism.time.in.hours))







