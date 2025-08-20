

# -------------------------------------------------------------
# SCRIPT- SOCIAL PROTECTION PROGRAMME RAPID ASSESSMENT FRAMEWORK
# -------------------------------------------------------------

#Date: August 2025
#Version 1.0

# 1| Preparation of the code --------------------------------------------------------

# Clear and start with a new project
rm(list = ls())
graphics.off()
gc()

# Paste start time - to estimate the whole time it takes to run
startTime  = Sys.time()

# 1.1| Libraries
# List of packages
myPackages <- c(
  'broom','caret','cluster','clValid','cobalt','colorspace','data.table','descr',
  'dplyr','extrafont','factoextra','FactoMineR','fastDummies','foreign','fpc','gbm',
  'geosphere','ggdendro','ggparty','ggplot2','ggpubr','ggspatial','ggmap','glmnet',
  'gridExtra','gtools','haven','here','Hmisc','igraph','Metrics','openxlsx','partykit',
  'PCAmixdata','ppcor','purrr','questionr','raster','RColorBrewer','readr','readxl',
  'reshape2','rpart','rpart.plot','scales','sf','shadowtext','spatstat','stars',
  'StatMatch','stringr','survey','tidyr','tidyverse','treemapify','writexl','eeptools',
  'lubridate','lattice','sfsmisc'
)

# Identify missing packages
notInstalled <- myPackages[!(myPackages %in% rownames(installed.packages()))]

# Install missing packages with dependencies
if (length(notInstalled)) {
  install.packages(notInstalled, dependencies = TRUE)
}

# Load packages one-by-one with feedback
for (pkg in myPackages) {
  message("Loading: ", pkg)
  tryCatch({
    library(pkg, character.only = TRUE, quietly = TRUE)
  }, error = function(e) {
    warning("Failed to load package: ", pkg, "\n", e$message)
  })
}

# Load fonts only if extrafont is available and on Windows
if ("extrafont" %in% loadedNamespaces() && .Platform$OS.type == "windows") {
  try(loadfonts(device = "win", quiet = TRUE), silent = TRUE)
}

# Disable scientific notation
options(scipen = 999)


# 1.2| Initial values locations and folders
# 1: English.
# 0: Arabic.
language = 0

# Specify location of file
userLocation   = enc2native(here()) # Replace by your own path.

## Create output folders if not available to save the output
if(!file.exists("Output")){
  dir.create("Output")  
}
if(!file.exists("Output/Clustering")){
  dir.create("Output/Clustering")  
}
if(!file.exists("Output/Targeting Assessment")){
  dir.create("Output/Targeting Assessment")  
}
if(!file.exists("Output/Coverage Evaluation")){
  dir.create("Output/Coverage Evaluation")  
}

# Location of Input and Code folders
scriptLocation = paste0(userLocation, '/Code/')
inputLocation  = paste0(userLocation, '/Input/')


# 1.3| Upload main datasets: Individual, Household Data, and beneficiaries  
IndividualData=read_excel(paste0(inputLocation, 'Bases/DATA_BASES.xlsx'), sheet = 'Individuals',na = "NULL")
HH_Data=read_excel(paste0(inputLocation, 'Bases/DATA_BASES.xlsx'), sheet = 'Households',na = "NULL")

IndividualData = IndividualData %>% 
  mutate(`Schooling level`=if_else(`Schooling level`%in%c("illiterat"), "illiterate",`Schooling level`))

IndividualData = IndividualData %>%
  mutate(Age = trunc((as.Date(`Date of birth`) %--% as.Date(Sys.Date()) / years(1) )))

Beneficiaries_Data=read.csv(paste0(inputLocation, 'Bases/Beneficiaries.csv'), na = "NULL")


# 1.4| Upload dictionary - Labels so that the graphs read both English and Arabic
dataPlots <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), 
                        sheet = 'Labels')
dataPlots[is.na(dataPlots)] <- '' #Replacing NAs with characters


# 1.5| Refer each graph to its specific labels in the dictionary

# The main objective is linking the labels to its specific location in the Dictionary (Excel file) sheet named "Category"
IndVarNames <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'A8:D33')) #From A8 to D33
HHVarNames <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'F8:I51')) 
Benificiaries_VarNames<- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'P14:S23'))


GenderLabels <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'A1:D3')) 
MarriageLabel<- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'U14:X18')) 
YesNoLabels <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'F1:I3')) 
EducationAttainment <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'K1:N5')) 
HighestEducation <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'P1:S9')) 
EmploymentStatus <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'U1:X4')) 
EmploymentType <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'Z1:AC7'))
EmploymentRegularity <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'AE1:AH5'))
EmploymentSector <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'AJ1:AM4')) 

DwellingType <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'AO1:AR12')) 
DwellingOwnership <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'AT1:AW5')) 
DwellingWalls <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'AY1:BB7')) 
DamageLevel <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'BD1:BG4')) 
DwellingRoof <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'BI1:BL6')) 
DwellingFloor <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'BN1:BQ6')) 
CookingFuel <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'BS1:BV5')) 
ToiletType <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'BX1:CA6')) 

WaterSource <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'CC1:CF6')) 
ElectricitySource <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'CH1:CK4')) 
SewageWater <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'CM1:CP4')) 
HHProperties <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'CR1:CU3')) 
AgriculturalLand <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'CW1:CZ5')) 

LocationLabels <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'DB1:DE4')) 
GovernorateLabels <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'DG1:DJ9')) 


# 1.6| Upload Weights for each variable from a different sheet named "Weights"
DwellingTypeWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'A2:C46') %>%  
  rename(`House type`=Dwelling_type, `Contract type`= Dwelling_ownership, var_Dwelling_type=weight)
WallTypeWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'E2:G20')%>% 
  rename(`Wall material`=Wall_type, `Wall quality`= Damage_level, var_Wall_type=weight)
RoofTypeWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'I2:K17')%>% 
  rename(`Roof material`=Roof_type, `Roof quality`= Damage_level, var_Roof_type=weight)
FloorTypeWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'M2:O17')%>% 
  rename(`Floor material`=Floor_type, `Floor quality`= Damage_level, var_Floor_type=weight)
CookingfuelWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'Q2:R6')%>% 
  rename(`Cooking source`=Cooking_fuel, var_Cooking_fuel=weight)
ToiletTypeWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'T2:U7')%>% 
  rename(`Toilet type`=Toilet_type, var_Toilet_Type=weight)
WaterWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'W2:X7')%>% 
  rename(`Water`=Water_source, var_Drinking_water=weight)
DrinkingWaterWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'W9:X14')%>% 
  rename(`Drinking water`=Drinking_water_source, var_water=weight)
SewageWaterWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'Z9:AA12')%>% 
  rename(`Sewers`=Sewage_water, var_sewage=weight)
ElectricitySourceWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'Z2:AA5')%>% 
  rename(`Electricity source`=Electricity_source, var_electricity=weight)
LivestockWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'AC2:AE8')%>% 
  rename(`Livestock`=Livestock,`Location`= Location, var_livestock=weight)
ReaslestateWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'AG2:AI8')%>% 
  rename(`Realstate`=Realestate, `Location`= Location, var_Realestate=weight)
AgriLandWeights <- read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Weights', range = 'AK2:AM14')%>% 
  rename(`Agriculture land`=Agricultural_land,`Location`= Location, var_AgriLand=weight)


# 1.7| Renaming variables in each dataset
colnames(IndividualData)=as.character(factor(colnames(IndividualData),
                                             levels = IndVarNames[,2],
                                             labels = IndVarNames[,4-1]))
colnames(HH_Data)=as.character(factor(colnames(HH_Data),
                                      levels = HHVarNames[,2],
                                      labels = HHVarNames[,4-1]))

colnames(Beneficiaries_Data)=as.character(factor(colnames(Beneficiaries_Data),
                                                 levels = Benificiaries_VarNames[,2],
                                                 labels = Benificiaries_VarNames[,4-1]))
HH_Data=HH_Data[,!is.na(colnames(HH_Data))]


# 1.8| Create Indicators for Household Dataset
# Include specific variables
Indicators=HH_Data%>%
  left_join(DwellingTypeWeights) %>%  
  left_join(WallTypeWeights) %>% 
  left_join(FloorTypeWeights) %>% 
  left_join(RoofTypeWeights) %>% 
  left_join(CookingfuelWeights) %>% 
  left_join(ToiletTypeWeights) %>% 
  left_join(WaterWeights) %>% 
  left_join(DrinkingWaterWeights) %>% 
  left_join(ElectricitySourceWeights) %>% 
  left_join(SewageWaterWeights) %>% 
  left_join(LivestockWeights) %>% 
  left_join(ReaslestateWeights) %>% 
  left_join(AgriLandWeights) %>% 
  mutate(peoplePerRoom=`Household members`/`Number of rooms`,
         var_overcrowdingIssue=if_else(peoplePerRoom>3,1,0),
         var_duration_market=if_else(`Time to the market`>60,100,0),
         var_duration_school=if_else(`Time to school`>60,100,0),
         var_duration_hospital=if_else(`Time to hospital`>60,100,0),
         kitchen=if_else(`Kitchen`%in%c("Yes"),0,100),
         television=if_else(`TV`>0,0,90),
         phone=if_else(`Mobile phone`>0,0,90),
         carr=if_else(`Car`>0,0,80),
         fridge=if_else(`Fridge`>0,0,70),
         washing_machine=if_else(`Washing machine`>0,0,60),
         AC=if_else(`AC`>0,0,60),
         computer=if_else(`Computer`>0,0,50),
         internet=if_else(`Internet`%in%c("Yes"),0,50),
         heater=if_else(`Heater`%in%c("Yes"),0,50),
         sewing_machine=if_else(`Sewing Machine`%in%c("Yes"),0,40),
         solar=if_else(`Solar Panel`%in%c("Yes"),0,10),
         bike=if_else(`Bike`>0,0,10),
         auto_wash_machine=if_else(`Automatic washing machine`>0,0,10),
         var_Property_score= 100*(kitchen+television+phone+carr+fridge+washing_machine+AC+computer+internet+heater+sewing_machine+solar+bike+auto_wash_machine)/770,
  )

# 1.9| Create Indicators for Individual Dataset
t1=IndividualData%>%group_by(HHID)%>%
  mutate(hhSize=length(HHID),
         var_insurance=if_else(`Health insurance`%in%c("No"),1,0),
         var_chronic=if_else(`Chronic disease`%in%c("Yes"),1,0),
         var_nutrition=if_else(`Nutrition Disease`%in%c("Yes"),1,0),
         var_common=if_else(`Common Disease`%in%c("Yes"),1,0),
         var_tot_disab=if_else(`Total disability`%in%c("Yes"),1,0),
         var_part_disab=if_else(`Partial disability`%in%c("Yes"),1,0),
         var_other_disab=if_else(`Other disability`%in%c("Yes"),1,0),
         var_polio=if_else(`Polio`%in%c("Yes"),1,0),
         var_diphtheria=if_else(`Diphtheria`%in%c("Yes"),1,0),
         var_measels=if_else(`Measels` %in%c("Yes"), 1, 0),
         #Age=trunc((as.Date(`Date of birth`) %--% as.Date("2023-08-28")) / years(1)),
         school_attendance=if_else(Age>6 & `School attendance`%in%c("was enrolled and dropped out","haven't been enrolled before"),1,0),
         members_15above=if_else(Age>15 ,1,0),
         illiterate=if_else(Age>15 & `Schooling level`%in%c("illiterate"),1,0),
         high_education=if_else(Age>18 & `Schooling level`%in%c("post-secondary diploma","university and above","high-school"),1,0),
         kids13_below_enrolled=if_else(Age<13 & `School attendance`%in%c("currently enrolled"),1,0),
         kids13_below=if_else(Age<13 ,1,0),
         inactive=if_else(Age>15 & Age <25 & `Employment Status`%in%c("inactive") & `School attendance`%in%c("was enrolled and dropped out","haven't been enrolled before"),1,0),
         members15_25= if_else(Age>15 & Age <25,1,0),
         members18_65= if_else(Age>18 & Age <65,1,0),
         econ_dependency=if_else(`Job Type`%in%c("paid in cash")|`Job Type`%in%c("paid both cash and in-kind"),1,0),
         workers18_65=if_else((`Employment Status`%in%c("employed") & Age>18 & Age<65),1,0),
         workers=if_else((`Employment Status`%in%c("employed")),1,0),
         unemployment= if_else((`Employment Status`%in%c("unemployed") & Age>18 & Age<65),1,0),
         temporary=if_else(`Contract type`%in%c("temporary","seasonal"),1,0),
         child_labour=if_else(`Employment Status`%in%c("employed") & Age<18,1,0),
         members18_below=if_else(Age<18,1,0),
         female=if_else(`Gender`%in%c("Female"),1,0),
         old_dependency= if_else(Age>65,1,0),
         young_dependency=if_else(Age<18,1,0)
  )%>%
  summarise(hhSize=length(HHID),
            var_insurance=100*mean(var_insurance, rm.na=F),
            var_chronic=100*mean(var_chronic, rm.na=F),
            var_nutrition=100*mean(var_nutrition, rm.na=F),
            var_common=100*mean(var_common, rm.na=F),
            var_tot_disab=100*mean(var_tot_disab, rm.na=F),
            var_part_disab=100*mean(var_part_disab, rm.na=F),
            var_other_disab=100*mean(var_other_disab, rm.na=F),
            var_polio=100*mean(var_polio, rm.na=F),
            var_diphtheria=100*mean(var_diphtheria, rm.na=F),
            var_measels=100*mean(var_measels, rm.na=F),
            var_enrolment=100*mean(school_attendance, rm.na=F),
            var_illiteracy_rate=100*sum(illiterate)/(sum(members_15above)),
            var_high_education_rate=100 - 100*mean(high_education, rm.na=F),
            var_kids_outside_school=100-100*(sum(kids13_below_enrolled))/(sum(kids13_below)),
            var_NEET=100*sum(inactive)/sum(members15_25),
            var_econ_dependency=100-(100*mean(econ_dependency)),
            var_unemployment_rate= 100*sum(unemployment)/(sum(unemployment)+sum(workers18_65)),
            var_temporary_rate=100*sum(temporary)/(sum(workers)),
            var_child_labour=100*sum(child_labour)/(sum(members18_below)),
            var_female_ratio=100*mean(female),
            var_old_dependency_ratio=(100*sum(old_dependency)/sum(members18_65))/2,
            var_young_dependency_ratio=(100*sum(young_dependency)/sum(members18_65))/3) %>% 
  mutate(
    var_illiteracy_rate=if_else(is.nan(var_illiteracy_rate),100,var_illiteracy_rate),
    var_kids_outside_school=if_else(is.nan(var_kids_outside_school),0,var_kids_outside_school),
    var_NEET=if_else(is.nan(var_NEET),0,var_NEET),
    var_unemployment_rate=if_else(is.nan(var_unemployment_rate),100,var_unemployment_rate),
    var_temporary_rate=if_else(is.nan(var_temporary_rate),100,var_temporary_rate),
    var_child_labour=if_else(is.nan(var_child_labour),0,var_child_labour),
    var_female_ratio=if_else(is.nan(var_female_ratio),100,var_female_ratio),
    var_old_dependency_ratio=if_else(is.nan(var_old_dependency_ratio) | var_old_dependency_ratio>100,100,var_old_dependency_ratio),
    var_young_dependency_ratio=if_else(is.nan(var_young_dependency_ratio) | var_young_dependency_ratio>100,100,var_young_dependency_ratio),
  )


# 1.10| Convert categorical variables of individuals dataset into factors, with labels translated according to the language
# Also, in case another language wants to be used as a label for the data, this code will update to the language selected (in this case, language 1 is English, 0 is Arabic).
IndividualData$`Gender`=as.character(factor(IndividualData$`Gender`,
                                            levels = GenderLabels[,2],
                                            labels = GenderLabels[,4-language]))
IndividualData$`School attendance`=as.character(factor(IndividualData$`School attendance`,
                                                       levels = EducationAttainment[,2],
                                                       labels = EducationAttainment[,4-language]))
IndividualData$`Schooling level`=as.character(factor(IndividualData$`Schooling level`,
                                                     levels = HighestEducation[,2],
                                                     labels = HighestEducation[,4-language]))

IndividualData$`Marital status`=as.character(factor(IndividualData$`Marital status`,
                                                    levels = MarriageLabel[,2],
                                                    labels = MarriageLabel[,4-language]))

IndividualData$Age=as.numeric(IndividualData$Age)
IndividualData$`Chronic disease`=as.character(factor(IndividualData$`Chronic disease`,
                                                     levels = YesNoLabels[,2],
                                                     labels = YesNoLabels[,4-language]))
IndividualData$`Nutrition Disease`=as.character(factor(IndividualData$`Nutrition Disease`,
                                                       levels = YesNoLabels[,2],
                                                       labels = YesNoLabels[,4-language]))
IndividualData$`Common Disease`=as.character(factor(IndividualData$`Common Disease`,
                                                    levels = YesNoLabels[,2],
                                                    labels = YesNoLabels[,4-language]))
IndividualData$`Total disability`=as.character(factor(IndividualData$`Total disability`,
                                                      levels = YesNoLabels[,2],
                                                      labels = YesNoLabels[,4-language]))
IndividualData$`Partial disability`=as.character(factor(IndividualData$`Partial disability`,
                                                        levels = YesNoLabels[,2],
                                                        labels = YesNoLabels[,4-language]))
IndividualData$`Other disability`=as.character(factor(IndividualData$`Other disability`,
                                                      levels = YesNoLabels[,2],
                                                      labels = YesNoLabels[,4-language]))
IndividualData$`Polio`=as.character(factor(IndividualData$`Polio`,
                                           levels = YesNoLabels[,2],
                                           labels = YesNoLabels[,4-language]))
IndividualData$`Measels`=as.character(factor(IndividualData$`Measels`,
                                             levels = YesNoLabels[,2],
                                             labels = YesNoLabels[,4-language]))
IndividualData$`Diphtheria`=as.character(factor(IndividualData$`Diphtheria`,
                                                levels = YesNoLabels[,2],
                                                labels = YesNoLabels[,4-language]))
IndividualData$`Health insurance`=as.character(factor(IndividualData$`Health insurance`,
                                                      levels = YesNoLabels[,2],
                                                      labels = YesNoLabels[,4-language]))
IndividualData$`Employment Status`=as.character(factor(IndividualData$`Employment Status`,
                                                       levels = EmploymentStatus[,2],
                                                       labels = EmploymentStatus[,4-language]))
IndividualData$`Job Type`=as.character(factor(IndividualData$`Job Type`,
                                              levels = EmploymentType[,2],
                                              labels = EmploymentType[,4-language]))
IndividualData$`Contract type`=as.character(factor(IndividualData$`Contract type`,
                                                   levels = EmploymentRegularity[,2],
                                                   labels = EmploymentRegularity[,4-language]))
IndividualData$`Sector`=as.character(factor(IndividualData$`Sector`,
                                            levels = EmploymentSector[,2],
                                            labels = EmploymentSector[,4-language]))

#For HH dataset
HH_Data$`Household members`=as.numeric(HH_Data$`Household members`)
HH_Data$`Number of rooms`=as.numeric(HH_Data$`Number of rooms`)

HH_Data$`House type`=as.character(factor(HH_Data$`House type`,
                                         levels = DwellingType[,2],
                                         labels = DwellingType[,4-language]))
HH_Data$`Contract type`=as.character(factor(HH_Data$`Contract type`,
                                            levels = DwellingOwnership[,2],
                                            labels = DwellingOwnership[,4-language]))
HH_Data$`Wall material`=as.character(factor(HH_Data$`Wall material`,
                                            levels = DwellingWalls[,2],
                                            labels = DwellingWalls[,4-language]))
HH_Data$`Wall quality`=as.character(factor(HH_Data$`Wall quality`,
                                           levels = DamageLevel[,2],
                                           labels = DamageLevel[,4-language]))
HH_Data$`Roof material`=as.character(factor(HH_Data$`Roof material`,
                                            levels = DwellingRoof[,2],
                                            labels = DwellingRoof[,4-language]))
HH_Data$`Roof quality`=as.character(factor(HH_Data$`Roof quality`,
                                           levels = DamageLevel[,2],
                                           labels = DamageLevel[,4-language]))
HH_Data$`Floor material`=as.character(factor(HH_Data$`Floor material`,
                                             levels = DwellingFloor[,2],
                                             labels = DwellingFloor[,4-language]))
HH_Data$`Floor quality`=as.character(factor(HH_Data$`Floor quality`,
                                            levels = DamageLevel[,2],
                                            labels = DamageLevel[,4-language]))
HH_Data$`Cooking source`=as.character(factor(HH_Data$`Cooking source`,
                                             levels = CookingFuel[,2],
                                             labels = CookingFuel[,4-language]))
HH_Data$`Toilet type`=as.character(factor(HH_Data$`Toilet type`,
                                          levels = ToiletType[,2],
                                          labels = ToiletType[,4-language]))
HH_Data$`Drinking water`=as.character(factor(HH_Data$`Drinking water`,
                                             levels = WaterSource[,2],
                                             labels = WaterSource[,4-language]))
HH_Data$`Water`=as.character(factor(HH_Data$`Water`,
                                    levels = WaterSource[,2],
                                    labels = WaterSource[,4-language]))
HH_Data$`Electricity source`=as.character(factor(HH_Data$`Electricity source`,
                                                 levels = ElectricitySource[,2],
                                                 labels = ElectricitySource[,4-language]))

HH_Data$`Sewers`=as.character(factor(HH_Data$`Sewers`,
                                     levels = SewageWater[,2],
                                     labels = SewageWater[,4-language]))
HH_Data$`Time to the market`=as.numeric(HH_Data$`Time to the market`)
HH_Data$`Time to school`=as.numeric(HH_Data$`Time to school`)
HH_Data$`Time to hospital`=as.numeric(HH_Data$`Time to hospital`)


HH_Data$`Kitchen`=as.character(factor(HH_Data$`Kitchen`,
                                      levels = HHProperties[,2],
                                      labels = HHProperties[,4-language]))
HH_Data$`Internet`=as.character(factor(HH_Data$`Internet`,
                                       levels = HHProperties[,2],
                                       labels = HHProperties[,4-language]))
HH_Data$`Heater`=as.character(factor(HH_Data$`Heater`,
                                     levels = HHProperties[,2],
                                     labels = HHProperties[,4-language]))
HH_Data$`Sewing Machine`=as.character(factor(HH_Data$`Sewing Machine`,
                                             levels = HHProperties[,2],
                                             labels = HHProperties[,4-language]))
HH_Data$`Solar Panel`=as.character(factor(HH_Data$`Solar Panel`,
                                          levels = HHProperties[,2],
                                          labels = HHProperties[,4-language]))
HH_Data$`TV`=as.numeric(HH_Data$`TV`)
HH_Data$`Mobile phone`=as.numeric(HH_Data$`Mobile phone`)
HH_Data$`Car`=as.numeric(HH_Data$`Car`)
HH_Data$`Fridge`=as.numeric(HH_Data$`Fridge`)
HH_Data$`Washing machine`=as.numeric(HH_Data$`Washing machine`)
HH_Data$`AC`=as.numeric(HH_Data$`AC`)
HH_Data$`Computer`=as.numeric(HH_Data$`Computer`)
HH_Data$`Bike`=as.numeric(HH_Data$`Bike`)
HH_Data$`Automatic washing machine`=as.numeric(HH_Data$`Automatic washing machine`)

HH_Data$`Livestock`=as.character(factor(HH_Data$`Livestock`,
                                        levels = YesNoLabels[,2],
                                        labels = YesNoLabels[,4-language]))
HH_Data$`Location`=as.character(factor(HH_Data$`Location`,
                                       levels = LocationLabels[,2],
                                       labels = LocationLabels[,4-language]))
HH_Data$`Realstate`=as.character(factor(HH_Data$`Realstate`,
                                        levels = YesNoLabels[,2],
                                        labels = YesNoLabels[,4-language]))
HH_Data$`Agriculture land`=as.character(factor(HH_Data$`Agriculture land`,
                                               levels = AgriculturalLand[,2],
                                               labels = AgriculturalLand[,4-language]))
HH_Data$`Governorate`=as.character(factor(HH_Data$`Governorate`,
                                          levels = GovernorateLabels[,2],
                                          labels = GovernorateLabels[,4-language]))


# 1.11| Select variables for the estimation of the Proxy Menas Test (PMT) - As provided

completeData= Indicators %>% 
  left_join(t1) %>%
  dplyr::select(c("HHID"), starts_with("var_"))

completeData_Vuln= Indicators %>% 
  left_join(t1) %>%
  dplyr::select(c("HHID", "Governorate", "Location"), starts_with("var_"))

PMTIndicators=completeData %>% 
  dplyr::select(c("HHID"), "var_overcrowdingIssue", "var_Wall_type","var_Floor_type", "var_Toilet_Type", "var_Drinking_water", "var_sewage",
                "var_Property_score", "var_AgriLand",
                "var_chronic","var_common", "var_measels", "var_tot_disab", "var_other_disab", "var_duration_hospital",
                "var_enrolment", "var_high_education_rate", "var_duration_school",
                "var_NEET", "var_unemployment_rate", "var_child_labour",
                "var_female_ratio", "var_young_dependency_ratio")





# 2| PMT estimation --------------------------------------------------------
# Merging all the variables included in the PMT - mainly are weighted variables and sum up to the SCORE
SCORING = PMTIndicators %>% 
  mutate(
    s_var_overcrowdingIssue=(1/6)*(1/7)*var_overcrowdingIssue, # the value next to each variable is its weight. in this case we specified the weight but in other cases we have to stick to the weight given
    s_var_Wall=(1/6)*(1/7)*var_Wall_type,
    s_var_Floor=(1/6)*(1/7)*var_Floor_type,
    s_var_Toilet=(1/6)*(1/7)*var_Toilet_Type,
    s_var_DrinkWater=(1/6)*(1/7)*var_Drinking_water,
    s_var_Sewage=(1/6)*(1/7)*var_sewage,
    s_var_Property=(1/2)*(1/7)*var_Property_score,
    s_var_agriLand=(1/2)*(1/7)*var_AgriLand,
    s_var_chronic=(1/6)*(1/7)*var_chronic,
    s_var_common=(1/6)*(1/7)*var_common,
    s_var_measels=(1/6)*(1/7)*var_measels,
    s_var_totdisab=(1/6)*(1/7)*var_tot_disab,
    s_var_otherdisab=(1/6)*(1/7)*var_other_disab,
    s_var_durationHospital=(1/6)*(1/7)*var_duration_hospital,
    s_var_enrolment=(1/3)*(1/7)*var_enrolment,
    s_var_higheducation=(1/3)*(1/7)*var_high_education_rate,
    s_var_durationSchool=(1/3)*(1/7)*var_duration_school,
    s_var_NEET=(1/3)*(1/7)*var_NEET,
    s_var_unemp=(1/3)*(1/7)*var_unemployment_rate,
    s_var_childLabour=(1/3)*(1/7)*var_child_labour,
    s_var_female=(1/2)*(1/7)*var_female_ratio,
    s_var_youngdependency=(1/2)*(1/7)*var_young_dependency_ratio,
    score = s_var_overcrowdingIssue  + s_var_Wall + s_var_Floor +  s_var_Toilet + s_var_DrinkWater + s_var_Sewage +
      s_var_Property + s_var_agriLand + s_var_chronic + s_var_common + s_var_measels + s_var_totdisab + s_var_otherdisab +
      s_var_durationHospital + s_var_enrolment+ s_var_higheducation + s_var_durationSchool + s_var_NEET +  s_var_unemp +
      s_var_childLabour +s_var_female +  s_var_youngdependency )%>%
  drop_na() #remove NA if available


# 2.1| Add the new SCORING variable to the different datasets
#important note is to consider the head of the household not the whole family
filtered_Head = IndividualData%>%filter(`Relationship with HH`%in% c(1)) %>% 
  dplyr::select(c("HHID", "Gender", "Marital status", "Schooling level", "Employment Status"))

completeData_Vuln= completeData_Vuln %>% 
  left_join(SCORING) %>%
  dplyr::select(c("HHID", "Governorate", "Location","score"), starts_with("var_"))

completeData_Vuln= completeData_Vuln %>% 
  left_join(filtered_Head) %>%
  dplyr::select(c("HHID", "Governorate", "Location","score", "Gender", "Marital status", "Schooling level", "Employment Status"), starts_with("var_"))

MergedData= merge(IndividualData,HH_Data,by="HHID")



# 3| Clustering --------------------------------------------------------

#Specify location to paste output for this exercise
outputLocation = paste0(userLocation, '/Output/Clustering/')

# 3.1| Filters and data preparation
# Filter only beneficiaries 
completeData_Vuln= completeData_Vuln %>% 
  left_join(Beneficiaries_Data) %>%filter(Beneficiaries_Data$Beneficiaries%in% c("Yes")) %>% 
  dplyr::select(c("HHID", "Governorate", "Location","score", "Gender", "Marital status", "Schooling level"), starts_with("var_"), -c('Beneficiaries')) 

# Specify the Eligibilible beneficiaries
Eligibles=HH_Data%>%left_join(Beneficiaries_Data%>%dplyr::select(c('HHID','Beneficiaries')))%>%
  group_by(HHID)

# Clean data for PMT:
t1= SCORING %>% 
  drop_na() %>% dplyr::select(c('HHID')) 
# Clean data for vulnerabilities
t2= completeData_Vuln %>% filter(Governorate%in%c("A")) %>% 
  drop_na() %>% dplyr::select(c('HHID')) # Governorate A is chosen as it has a decent size for the sake of the exercise

keyData = intersect(t1,t2) # output is HHID that are common in both t1 and t2

table(completeData_Vuln$Governorate)

# 3.2| Prepare the data for clustering
BeneficiariesSelectionData=keyData%>%left_join(completeData_Vuln )%>% dplyr::select(-c("HHID"))

BeneficiariesDistanceData=keyData%>%left_join(SCORING) %>% dplyr::select(starts_with("s_var_"))

BeneficiariesData=cbind(BeneficiariesDistanceData,BeneficiariesSelectionData)


# 3.3| Call customized functions and apply them

# Run a separate file contains all the functions needed
source(paste0(scriptLocation, 'Aux - Functions.R'), encoding = 'UTF-8')

# Limit the clustering to 200 with a final cluster number/group of 5
#⚠️ this could be modified by the user
basketLimit=max(15,ceiling(dim(BeneficiariesData)[1]*0.025))
PolicyGroups=5


# 3.4| Test the profiling as specified previously
resultList=clusteringFunction(BeneficiariesSelectionData, BeneficiariesDistanceData, basketLimit, PolicyGroups) #run it once with save below then comment both later
# save(resultList, file = paste0(outputLocation,Gov,Program,"ResultTree.RData"))
#load(file = paste0(outputLocation,Gov,Program,"ResultTree.RData")) # uncomment load after result and save

# Show Clustering 
net=resultList$initialTreeConfig #start with original tree
#plot(net$net, layout=net$lo, vertex.size=0, edge.arrow.size=0, vertex.label=NA) 

clust=resultList$clusterCluster

# Recreating
originalData=BeneficiariesSelectionData

# Paste clusters column 
# Assigns the 'aggregatedCluster' results from resultList into the 'Clusters' column of originalData  
originalData$Clusters=resultList$aggregatedCluster
# Create control parameters for the decision tree, maxdepth = 4
control <- rpart.control(maxdepth = 4,
                         cp = 0.000000000000000000001)
# Build a decision tree classifier to predict 'Clusters' based on all other variables in originalData  
decision = rpart(Clusters ~., data = originalData, method = "class", control=control)
newClusters=predict(decision, originalData, type = 'class')
# Replace the Clusters column in originalData with the new predicted cluster labels  
originalData$Clusters=newClusters
control <- rpart.control(minbucket = 2,
                         cp = 0.000000000000000000001)
# Build another decision tree with the updated Clusters  
decision = rpart(Clusters ~., data = originalData, method = "class", control=control)
# Plot the final decision tree
rpart.plot(decision)
table(newClusters)



# 3.5| Graphs
wb = openxlsx::createWorkbook(creator = 'ESCWA')

# 3.5.1| Dendrogram 
rowLine = 1
jumpOfRows = 18
colPlots = 1
colTables = 12

# Specify an indicatorCOde to be used as the name of the graph
indicatorCode <- '1_0'
part <- ''
labCodes <- dataPlots %>% subset(Code == paste0(indicatorCode, part))
newSheet <- addWorksheet(wb, sheetName = indicatorCode)


# To export:
base_exp = 1
heightExp = 1.5
widthExp = 1.2
scale_factor = base_exp/widthExp

# Cut same as number of maximum clusters
cutNumber=PolicyGroups
newPlot = fviz_dend(clust, k = cutNumber, color_labels_by_k = T, palette = as.character(labCodes[, 12]), show_labels = F,
                    rect = T, lower_rect = -0.02, rect_border = as.character(labCodes[, 12]), rect_fill = T) +
  labs(title = paste0(as.character(labCodes[, 5*(1-language)+2])),
       subtitle = as.character(labCodes[, 5*(1-language)+3]),
       caption = as.character(labCodes[, 5*(1-language)+6]),
       x = as.character(labCodes[, 5*(1-language)+4]),
       y = as.character(labCodes[, 5*(1-language)+5])) +
  scale_y_continuous(position = if (language == 0) {'right'} else {'left'}) + 
  theme_classic() +
  theme(legend.position = 'bottom',
        text = element_text(family = 'Georgia'),
        axis.line.x = element_blank(),
        axis.line.y = element_blank(),
        axis.text.x = element_blank(),
        axis.ticks.x = element_blank(),
        axis.text.y = element_text(size = scale_factor * 10, family = 'Georgia'),
        axis.title.x = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
        axis.title.y = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
        legend.text = element_text(size = scale_factor * 10, family = 'Georgia'),
        strip.text = element_text(size = scale_factor * 10, family = 'Georgia'),
        plot.title = element_text(face = 'bold', size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
        plot.subtitle = element_text(size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
        legend.key.size = unit(0.5 * scale_factor, 'cm'))

# Export graph and table to excel:
fileName = paste0(fileName = paste0(outputLocation, indicatorCode, part, 
                                    if(language == 1) {' (English)'} else {' (Arab)'}, '.png'))

ggsave(fileName, plot = newPlot, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
insertImage(wb, file = fileName, sheet = indicatorCode, startRow = rowLine, startCol = colPlots, width = 6 * widthExp, height = 4 * heightExp * widthExp)

newData=as.data.frame(as.matrix(table(resultList$aggregatedCluster)))%>%
  mutate(Clusters=rownames(table(resultList$aggregatedCluster)))%>%
  rename(Observations=V1)%>%
  dplyr::select(Clusters,Observations)

writeDataTable(wb, sheet = indicatorCode, x = newData, startRow = rowLine, startCol = colTables)
rowLine = rowLine + jumpOfRows


# 3.5.2| Decision tree
rowLine = 1
jumpOfRows = 38
colPlots = 1
colTables = 12

indicatorCode <- '1_1' # A different indicatorCode
part <- ''
labCodes <- dataPlots %>% subset(Code == paste0(indicatorCode, part))
newSheet <- addWorksheet(wb, sheetName = indicatorCode)

# To export:
base_exp = 1
heightExp = 1.5
widthExp = 1.2
scale_factor = base_exp/widthExp

# Summary
pct <- as.party(decision)

g <- ggparty(pct, terminal_space = 0.5)+
  labs(
    title = as.character(labCodes[, 5*(1-language)+2]),
    subtitle = as.character(labCodes[, 5*(1-language)+3]),
    caption = as.character(labCodes[, 5*(1-language)+6]),
    fill = '')+
  theme(
    plot.title = element_text(face = 'bold', size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
    plot.subtitle = element_text(size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}))

g <- g + 
  geom_edge(size = 1, colour = 'grey')
g <- g + 
  geom_edge_label(colour = 'black', size = 2)
g <- g + 
  geom_node_plot(gglist = list(geom_bar(aes(x = '', fill = Clusters), color = 'black', position = position_fill()), 
                               scale_y_continuous(expand = expansion(mult = c(0, 0)),
                                                  labels = scales::percent_format(scale = 100, suffix = '')),
                               scale_x_discrete(expand = expansion(mult = c(0, 0))),
                               theme_classic(),
                               scale_fill_brewer(palette = as.character(labCodes[, 12])),
                               labs(
                                 x = as.character(labCodes[, 5*(1-language)+4]),
                                 y = as.character(labCodes[, 5*(1-language)+5]),
                                 fill = ''),
                               theme(legend.position = 'bottom',
                                     text = element_text(family = 'Georgia'),
                                     axis.text.y = element_text(size = scale_factor * 10, family = 'Georgia'),
                                     axis.title.x = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
                                     axis.title.y = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
                                     axis.ticks.x = element_blank(),
                                     legend.text = element_text(size = scale_factor * 10, family = 'Georgia'),
                                     strip.text = element_text(size = scale_factor * 10, family = 'Georgia'),
                                     plot.title = element_text(face = 'bold', size = 10, family = 'Georgia'),
                                     plot.subtitle = element_text(size = 10, family = 'Georgia'),
                                     legend.key.size = unit(0.5 * scale_factor, 'cm'))),
                 scales = 'fixed', # y and x axis fixed for better comparability.
                 ids = 'terminal', # Labels for leaves.
                 shared_axis_labels = TRUE,
                 shared_legend = TRUE,
                 legend_separator = FALSE)
g <- g + 
  geom_node_label(aes(col = splitvar),
                  line_list = list(aes(label = splitvar)),
                  line_gpar = list(list(size = 12,
                                        fontface = 'bold')),
                  ids = 'inner')
g <- g + 
  geom_node_label(aes(label = paste0('Node ', id, '\nN = ', nodesize)),
                  fontface = 'bold',
                  ids = 'terminal',
                  size = 4,
                  nudge_y = 0.02)
newPlot <- g + theme(legend.position = 'none')

newPlot

# Export graph and table to excel:
fileName = paste0(fileName = paste0(outputLocation, indicatorCode, part,
                                    if(language == 1) {' (English)'} else {' (Arab)'}, '.png'))
ggsave(fileName, plot = newPlot, width = 25 * widthExp, height = 6 * heightExp * widthExp, scale = scale_factor)
insertImage(wb, file = fileName, sheet = indicatorCode, startRow = rowLine, startCol = colPlots, width = 25 * widthExp, height = 4 * heightExp * widthExp)
rowLine = rowLine + jumpOfRows

newData=as.data.frame(as.matrix(table(newClusters)))%>%
  mutate(Clusters=rownames(table(newClusters)))%>%
  rename(Observations=V1)%>%
  dplyr::select(Clusters,Observations)

writeDataTable(wb, sheet = indicatorCode, x = newData, startRow = rowLine, startCol = colTables)
rowLine = rowLine + jumpOfRows*1.5


# 3.5.3| Descriptive graphs by cluster
# Split the data and select continuous and discrete graphs 
mainData=BeneficiariesSelectionData[,] #if there were a score we add it to column # changed into data selection for all of variables not weighted

# Before all is in English, translate anything here if needed to Arabic or any other language based on qualy variables

# Split quanty / qualy
# Distinguishing qualitative and quantitative vars for plotting and reporting
split <- splitmix(mainData) 
colnames(split$X.quali)=gsub("\\.", " ", colnames(split$X.quali))
colnames(split$X.quanti)=gsub("\\.", " ", colnames(split$X.quanti))

qualiMainData=mainData%>% dplyr::select(colnames(split$X.quali))%>%
  mutate(Clusters=factor(originalData$Clusters))

quantiMainData=mainData%>% dplyr::select(colnames(split$X.quanti))%>%
  mutate(Clusters=factor(originalData$Clusters))

# '1_X: Qualitative variables plot production
counter=1
numberGraphs=dim(qualiMainData)[2]-1
i=1
indicatorCoder<- "Qualitative variables"
newSheet <- addWorksheet(wb, sheetName = indicatorCoder) #new sheet
rowLine = 1
jumpOfRows = 20
colPlots = 1
colTables = 12

# To know what to add to the dictionary graphs for qualy
for(i in 1:length(colnames(split$X.quali))){
  counter=counter+1
  
  indicatorCode <- paste0('1_',counter)
  part <- ''
  labCodes <- dataPlots %>% subset(Code == paste0(indicatorCode, part))
  
  # To export qualitative graphs:
  base_exp = 1
  heightExp = 1
  widthExp = 1
  scale_factor = base_exp/widthExp
  
  temp=qualiMainData%>% dplyr::select(c(colnames(split$X.quali)[i],"Clusters"))%>%
    mutate(ones=1)
  colnames(temp)=c("Variable", "Clusters","Values")
  
  colourCount = length(unique(temp$Variable))
  getPalette = colorRampPalette(brewer.pal(9, as.character(labCodes[, 12])))
  
  newPlot=temp%>%group_by(Variable, Clusters)%>%
    summarise(Values=sum(Values, rm.na=F))%>%
    ggplot(aes(x=Clusters, y=Values, fill=Variable))+
    geom_bar(position = "fill", stat = "identity", color = 'black')+
    labs(title = paste0(as.character(labCodes[, 5*(1-language)+2])),
         subtitle = as.character(labCodes[, 5*(1-language)+3]),
         caption = as.character(labCodes[, 5*(1-language)+6]),
         x = as.character(labCodes[, 5*(1-language)+4]),
         y = as.character(labCodes[, 5*(1-language)+5])) +
    scale_fill_manual(values = getPalette(colourCount))+
    scale_x_discrete(position = if (language == 0) {'top'} else {'bottom'})+
    theme_classic() +coord_flip()+
    theme(legend.position = 'bottom',
          text = element_text(family = 'Georgia'),
          axis.text.x = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          legend.text = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title = element_text(face = 'bold', size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
          plot.subtitle = element_text(size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill = guide_legend(title = ''))
  
  if(language==0){
    newPlot=newPlot+scale_y_reverse(expand = expansion(mult = c(0.05, 0)),
                                    labels = scales::percent_format(scale = 100, accuracy = 0.1, suffix = ""))
    
  }else{
    newPlot=newPlot+scale_y_continuous(expand = expansion(mult = c(0, .05)),
                                       labels = scales::percent_format(scale = 100, accuracy = 0.1, suffix = ""))
  }
  newData=temp%>%group_by(Variable, Clusters)%>%
    summarise(Values=sum(Values, rm.na=F))%>%ungroup()%>%
    group_by(Clusters)%>%
    mutate(Values=round(100*Values/sum(Values, rm.na=F),2))%>%
    spread(Clusters, Values)%>%
    rename(Percentage=Variable)
  
  # Export graph and table to excel:
  fileName = paste0(fileName = paste0(outputLocation, indicatorCode, part, 
                                      if(language == 1) {' (English)'} else {' (Arab)'}, '.png'))
  
  ggsave(fileName, plot = newPlot, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  insertImage(wb, file = fileName, sheet = indicatorCoder, startRow = rowLine, startCol = colPlots, width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, sheet = indicatorCoder, x = newData, startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
}

# '1_X:Quantitative plot production
quantiMainData=quantiMainData
# here no need to specify since already selected properly this way
colnames(quantiMainData)

numberGraphs=dim(quantiMainData)[2]-1
i=1
indicatorCoder<- "Quantitative variables"
newSheet <- addWorksheet(wb, sheetName = indicatorCoder)
rowLine = 1
jumpOfRows = 20
colPlots = 1
colTables = 12

for(i in 1:numberGraphs){
  counter=counter+1
  
  indicatorCode <- paste0('1_',counter)
  part <- ''
  labCodes <- dataPlots %>% subset(Code == paste0(indicatorCode, part))
  
  # To export quantitative graphs:
  base_exp = 1
  heightExp = 1
  widthExp = 1
  scale_factor = base_exp/widthExp
  
  temp=quantiMainData%>% dplyr::select(c(colnames(quantiMainData)[i],"Clusters"))
  
  colnames(temp)=c("Values", "Clusters")
  newPlot=temp%>%group_by(Clusters)%>%
    summarise(Values=mean(Values, rm.na=F))%>%
    ggplot(aes(x=Clusters, y=Values, fill=factor(1)))+
    geom_bar(position = position_dodge(), stat = "identity", color = 'black')+
    labs(title = paste0(as.character(labCodes[, 5*(1-language)+2])),
         subtitle = as.character(labCodes[, 5*(1-language)+3]),
         caption = as.character(labCodes[, 5*(1-language)+6]),
         x = as.character(labCodes[, 5*(1-language)+4]),
         y = as.character(labCodes[, 5*(1-language)+5])) +
    scale_fill_brewer(palette = as.character(labCodes[, 12])) + 
    scale_x_discrete(position = if (language == 0) {'top'} else {'bottom'})+
    theme_classic() +coord_flip()+
    theme(legend.position = 'bottom',
          text = element_text(family = 'Georgia'),
          axis.text.x = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          legend.text = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title = element_text(face = 'bold', size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
          plot.subtitle = element_text(size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill = "none")
  if(language==0){
    newPlot=newPlot+scale_y_reverse(expand = expansion(mult = c(0.05, 0)),
                                    labels = scales::percent_format(scale = 1, accuracy = 0.01, suffix = ""))
    
  }else{
    newPlot=newPlot+scale_y_continuous(expand = expansion(mult = c(0, .05)),
                                       labels = scales::percent_format(scale = 1, accuracy = 0.01, suffix = ""))
  }
  
  newData=temp%>%group_by(Clusters)%>%
    summarise(Values=mean(Values, rm.na=F))%>%ungroup()
  
  # Export graph and table to excel:
  fileName = paste0(fileName = paste0(outputLocation, indicatorCode, part, 
                                      if(language == 1) {' (English)'} else {' (Arab)'}, '.png'))
  
  ggsave(fileName, plot = newPlot, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  insertImage(wb, file = fileName, sheet = indicatorCoder, startRow = rowLine, startCol = colPlots, width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, sheet = indicatorCoder, x = newData, startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
}
# Final step is to create excel files of the output in EN and AR
saveWorkbook(wb,
             file = paste0(outputLocation,  if(language == 1) {'CH 1 EN.xlsx'} else {'CH 1 AR.xlsx'}),  
             overwrite = TRUE)


# 3.6| Heat map, main characteristics by cluster
# Calculate and display the percentage distribution of each categorical variable within clusters
proportion_table <- qualiMainData %>%
  dplyr::select(Clusters, where(is.character)) %>%  # Keep only character (categorical) variables
  pivot_longer(
    cols = -Clusters,
    names_to = "Variable",
    values_to = "Category"
  ) %>%
  count(Clusters, Variable, Category) %>%  # Count occurrences
  group_by(Clusters, Variable) %>%
  mutate(
    Proportion = n / sum(n) * 100  # Convert to percentage within each cluster
  ) %>%
  arrange(Clusters, Variable, Category) %>%
  arrange(desc(Proportion))

# Convert the proportion table to wide format for each cluster
proportion_wide <- proportion_table %>%
  mutate(Variable_Category = paste(Variable, Category, sep = "_")) %>%
  dplyr::select(Clusters, Variable_Category, Proportion) %>%
  pivot_wider(
    names_from = Clusters,
    values_from = Proportion,
    names_prefix = "Cluster.",
    names_sort = TRUE
  ) %>%
  mutate(across(starts_with("Cluster"), ~replace_na(., 0)))  # Fill missing with 0

# Select only relevant columns for reporting
quali_summary <- proportion_wide %>%
  ungroup() %>%
  dplyr::select(Variable_Category, starts_with("Cluster"))

# 3.6.1| SUMMARY STATISTICS FOR NUMERIC VARIABLES

# Compute mean for numeric variables by cluster
numeric_stats <- data.frame(t(quantiMainData %>%
                                group_by(Clusters) %>%
                                summarise(
                                  across(where(is.numeric), list(
                                    Mean = ~mean(., na.rm = TRUE)
                                  ))
                                )))

# Prepare summary for numeric variables
quanti_summary <- data.frame(numeric_stats[-1,] %>%  # Drop cluster row (transposed summary)
                               rename_with(
                                 ~ ifelse(startsWith(., "X"), 
                                          paste("Cluster", gsub("X", "", .)), 
                                          .)
                               ) %>%
                               mutate(Variable_Category = row.names(numeric_stats[-1,])) %>%
                               arrange(desc(`Cluster 1`), asc(`Cluster 2`)) %>%  # Example ordering
                               dplyr::select(Variable_Category, starts_with("Cluster")))

# 3.6.2| Bind categorical and numeric variables

summary_table = data.frame(rbind(quanti_summary, quali_summary))  # Combine both summaries

# Clean up any list-columns
summary_table <- summary_table %>% 
  mutate(across(where(is.list), ~unlist(.x))) %>%
  arrange(Variable_Category)


# 3.6.3| Excel report output
# Create a new Excel workbook
wb1 <- createWorkbook()

# Write summary table to "Cluster Heatmap" sheet
addWorksheet(wb1, "Cluster Heatmap")
writeData(wb1, sheet = "Cluster Heatmap", summary_table, startRow = 1, startCol = 1)

# Highlight max and min values 

highlight_max_values <- function(wb, sheet, data) {
  cluster_cols <- grep("^Cluster", names(data), value = TRUE)
  
  for (i in 1:nrow(data)) {
    row_values <- as.numeric(data[i, cluster_cols])
    max_val <- max(row_values)
    max_cols <- which(row_values == max_val)
    min_val <- min(row_values)
    min_cols <- which(row_values == min_val)
    
    # Highlight max values (green)
    for (col in max_cols) {
      addStyle(wb, sheet = sheet,
               style = createStyle(fgFill = "#A1D99B", textDecoration = "bold"),
               rows = i + 1,
               cols = which(names(data) == cluster_cols[col]))
    }
    # Highlight min values (yellow)
    for (col in min_cols) {
      addStyle(wb, sheet = sheet,
               style = createStyle(fgFill = "yellow", textDecoration = "bold"),
               rows = i + 1,
               cols = which(names(data) == cluster_cols[col]))
    }
  }
}

# Apply highlight function to the heatmap
highlight_max_values(wb1, "Cluster Heatmap", summary_table)

# Auto-resize columns
setColWidths(wb1, sheet = "Cluster Heatmap", cols = 1:ncol(summary_table), widths = "auto")

# Add a second sheet showing the full beneficiary list with assigned clusters
addWorksheet(wb1, "List")
writeData(wb1, sheet = "List", mainData, startRow = 1, startCol = 1)

setColWidths(wb1, sheet = "List", cols = 1:(ncol(mainData) + 1), widths = "auto")


output_file <- file.path(outputLocation, paste0("cluster_heatmap",  ".xlsx"))

# Save the new workbook
saveWorkbook(wb1, output_file, overwrite = TRUE)


# 4| Targeting Assessment  -----------------------------------------------------------
# This section analyzes the targeting performance of the social assistance programme by:
        # 1. Assessing variable importance in the PMT model
        # 2. Identifying redundant variables using LASSO regression
        # 3. Visualizing the results for decision-making



# Set output location for targeting assessment results
outputLocation = paste0(userLocation, '/Output/Targeting Assessment/')


# 4.1| Load labels and ranking priorities from dictionary
# These define how variables should be categorized and prioritized
RankingPriorities <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'K8:N12')) #from to
RankNames <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'K14:N36')) #from to

# Create a new workbook
wb2 = openxlsx::createWorkbook(creator = 'ESCWA')

# 4.2| Prepare data for targeting analysis
# Select relevant variables from scoring data
cleanData=SCORING%>% 
  dplyr::select("HHID", "score", starts_with("s_var_"))

cleanData = cleanData %>% rename(X0=`score`) %>%
  mutate(X0=X0+rnorm(dim(cleanData)[1],0,1)) #singular matrix, data might have mistakes so by introducing a systematical error its better
TestingData=cleanData%>% dplyr::select(-c("X0","HHID")) #remove score then put it back again at the end of the dataset
TestingData$score=cleanData$X0

# 4.3| Analyze variable importance using semi-partial correlation
# This measures each variable's unique contribution to explaining the score
sp = spcor(TestingData, method = c('pearson')) # Semi-partial correlation.
reg=lm(score~., TestingData)
summary(reg)$r.squared
summary(reg)

# Extract and process important weights

weights           = sp$estimate^2
weights           = data.frame(Values = weights[dim(weights)[1],-dim(weights)[2]])
weights$names     = TestingData %>% dplyr::select(-c('score')) %>% names() %>% .[1:nrow(weights)]
colnames(weights) = c('Values','Variable')
# Calculate relative importance metrics
weights$Variable = factor(weights$Variable)
R2 = summary(lm(score~., TestingData))$r.squared
summary(lm(score~., TestingData))

# Clean variable names by extracting the core component after underscores
result=NULL
tempRow=strsplit(as.character(weights$Variable),"_")
for(i in 1:dim(weights)[1]){
  result=c(result,tempRow[[i]][3])  # it was 1, in this particular case we have 3 components of the name
}
result=result[!result%in%c("Constant")]

# Create treemap visualization of variable importance
weights1=weights %>%
  mutate(Values   = Values/R2,
         Variable = result  #fct_reorder(result, Values)
  )%>% drop_na()%>%
  group_by(Variable)%>%
  summarise(Values=sum(Values))%>%
  arrange(-Values)%>%
  mutate(
    Values=Values/sum(Values),
    Variable = fct_reorder(Variable, Values),
    AggregateSum=cumsum(Values)/sum(Values), #cumsum because we have biggest to smallest ranked variables
    ranking=case_when(
      AggregateSum>0.95&AggregateSum<=1.01 ~ 4,
      AggregateSum>0.90&AggregateSum<=0.95 ~ 3,
      AggregateSum>0.80&AggregateSum<=0.90 ~ 2,
      AggregateSum>0.00&AggregateSum<=0.80 ~ 1),
    ranking=factor(ranking,levels = RankingPriorities[,2],
                   labels = RankingPriorities[,4-language]))

# 4.4| Visualize variable importance
# Set up plotting parameters

# Plot dendrogram
rowLine    = 1
jumpOfRows = 18
colPlots   = 1
colTables  = 20

indicatorCode <- '2_1'
part          <- ''
labCodes      <- dataPlots %>% subset(Code == paste0(indicatorCode, part))
newSheet      <- addWorksheet(wb2, sheetName = indicatorCode)

base_exp     = 1
heightExp    = 1.5
widthExp     = 1.2
scale_factor = base_exp/widthExp
weights1=weights1%>%
  mutate(Variable=as.character(factor(Variable,
                                      levels = RankNames[,2],
                                      labels = RankNames[,4-language])))

# Graph for distribution of the influence
newPlot=weights1 %>%
  ggplot(aes(area = Values, fill = ranking,
             subgroup = Variable)) +
  geom_treemap() +
  geom_treemap_subgroup_text(place = "centre", grow = TRUE, alpha = 0.5)+
  labs(title    = paste0(as.character(labCodes[, 5*(1-language)+2])), 
       subtitle = as.character(labCodes[, 5*(1-language)+3]),
       caption  = as.character(labCodes[, 5*(1-language)+6])) +
  #x        = as.character(labCodes[, 5*(1-language)+4]),
  #y        = as.character(labCodes[, 5*(1-language)+5]))
  
  scale_fill_brewer(palette = as.character(labCodes[, 12])) + 
  theme_classic() +
  theme(legend.position = 'bottom',
        text            = element_text(family = 'Georgia'),
        axis.line.x     = element_blank(),
        axis.line.y     = element_blank(),
        axis.text.x     = element_blank(),
        axis.ticks.x    = element_blank(),
        axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
        axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
        axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
        legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
        strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
        plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
        plot.subtitle   = element_text(size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
        legend.key.size = unit(0.5 * scale_factor, 'cm'))+
  guides(fill  = guide_legend(title = ''))

# Export graph and table to excel:
fileName = paste0(fileName = paste0(outputLocation, indicatorCode, part, 
                                    if(language == 1) {' (English)'} else {' (Arab)'}, '.png'))

ggsave(fileName, plot = newPlot, width = 12 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
insertImage(wb2, file = fileName, sheet = indicatorCode, startRow = rowLine, startCol = colPlots, width = 12 * widthExp, height = 4 * heightExp * widthExp)

newData=weights1%>%dplyr::select(-c("AggregateSum"))%>% 
  spread(ranking,Values)

writeDataTable(wb2, sheet = indicatorCode, x = newData, startRow = rowLine, startCol = colTables)
rowLine = rowLine + jumpOfRows




# 4.5| Redundant Variable Analysis Using Lasso Regression 

# Fit a linear model with all predictors to predict 'score'
regression = lm(score ~ ., TestingData)

# Extract regression coefficients, excluding the intercept, and append 1 for alignment
multiples = c(regression$coefficients[-1], 1)

# Replace NA coefficients (in case some predictors were dropped in lm) with 0
multiples[is.na(multiples)] = 0

# Scale each predictor column by its corresponding coefficient (used as a weight)
for(i in 1:length(multiples)){
  TestingData[,i] = TestingData[,i] * multiples[i]
}


# Create design matrix X for predictors (automatically handles factors/dummies)
X <- model.matrix(score ~ ., TestingData)

y <- TestingData$score 

# Set unconstrained bounds for LASSO coefficients
lb <- rep(-Inf, length(colnames(X)))
ub <- rep(Inf, length(colnames(X)))

cv_las1 = cv.glmnet(x = X, y = y, lower.limits = lb, upper.limits = ub)
lambda = cv_las1$lambda.min  # Best lambda minimizing CV error


# Fit LASSO model with optimal lambda 

las1 = glmnet(x = X, y = y, lower.limits = lb, upper.limits = ub, lambda = lambda)

# Extract coefficients and zero out small ones (< 0.05)
c.fit1 = coef(las1) %>% as.matrix() %>% as.data.frame()
colnames(c.fit1) = c('Coefficient')
namesVar = rownames(c.fit1)
c.fit1$Coefficient[abs(c.fit1$Coefficient) < 0.05] = 0  # Thresholding

# Compute RMSE adjusted by number of predictors
RMSE = sum((y - predict(las1, newx = X))^2 / (length(y) - (dim(X)[2] - 1)))

# Combine RMSE and coefficients into one result data frame
result = as.data.frame(c(RMSE, c.fit1$Coefficient))
namesChosen = c("RMSE", rownames(c.fit1))
result$Variable = namesChosen
result$Lambda = lambda
colnames(result)[1] = "Values"
result$Variable[2] = "Intercept"
result = result[!result$Variable %in% result$Variable[3],]  # Remove duplicate

# Clean and standardize variable names 
result$Variable = substring(result$Variable, 1, nchar(result$Variable))  # Safe trimming
fixingNames = NULL
tempRow = strsplit(as.character(result$Variable), "_")
for(i in 1:dim(result)[1]){
  fixingNames = c(fixingNames, tempRow[[i]][3])  # Take third part if split by "_"
}
result$Variable = gsub('X.', '', fixingNames)  # Remove R-style dummy prefixes
result$Variable = gsub('\\.', ' ', result$Variable)  # Replace dots with spaces


# LASSO across a wider lambda range 

lMin = 0.01 * lambda  # Lower bound: 1% of optimal lambda
lMax = 500 * lambda   # Upper bound: 500x of optimal lambda

stopCondition = TRUE
i = lMin
resu = NULL  # To store number of active variables and lambda

while(stopCondition){
  # Fit LASSO at current lambda
  las1 = glmnet(x = X, y = y, lower.limits = lb, upper.limits = ub, lambda = i)
  
  # Extract and threshold coefficients
  c.fit1 = coef(las1) %>% as.matrix() %>% as.data.frame()
  colnames(c.fit1) = c('Coefficient')
  namesVar = rownames(c.fit1)
  c.fit1$Coefficient[abs(c.fit1$Coefficient) < 0.05] = 0
  
  # Compute RMSE at current lambda
  RMSE = sum((y - predict(las1, newx = X))^2 / (length(y) - (dim(X)[2] - 1)))
  
  # Store result with variable names and lambda
  resultT = as.data.frame(c(RMSE, c.fit1$Coefficient))
  namesChosen = c("RMSE", rownames(c.fit1))
  resultT$Variable = namesChosen
  resultT$Lambda = i
  colnames(resultT)[1] = "Values"
  resultT$Variable[2] = "Intercept"
  resultT = resultT[!resultT$Variable %in% resultT$Variable[3],]
  
  # Clean variable names again
  resultT$Variable = substring(resultT$Variable, 1, nchar(resultT$Variable))
  fixingNames = NULL
  tempRow = strsplit(as.character(resultT$Variable), "_")
  for(j in 1:dim(resultT)[1]){
    fixingNames = c(fixingNames, tempRow[[j]][3])  
  }
  resultT$Variable = gsub('X.', '', fixingNames)
  resultT$Variable = gsub('\\.', ' ', resultT$Variable)
  
  # Count number of non-zero variables
  temp = resultT[-c(1,2),] %>% group_by(Variable) %>%
    summarise(Values = max(abs(Values))) %>% filter(!Values == 0)
  
  # Store number of active variables and current lambda
  resu = rbind(resu, cbind(dim(temp)[1], i))
  
  # If all coefficients are zero (except RMSE and intercept), stop
  if(sum(resultT$Values == 0) == (dim(resultT)[1] - 2)){
    stopCondition = FALSE
  }
  
  # Append to cumulative results
  result = rbind(result, resultT)
  
  print(i)  # Monitor progress
  i = i + (lMax - lMin) / 200  # Increment lambda in 200 steps
}

# Process summary of variable retention 
colnames(resu) = c("Cases", "Values")
varNumber = length(unique(result$Variable)) - 2  # Total number of variables used

# Track baseline (lowest lambda)
resuT = as.data.frame(resu) %>% filter(Values %in% lMin) %>%
  mutate(Variable = 1 - Cases / varNumber)

# For each unique number of variables retained, get max lambda that allows it
resu = as.data.frame(resu) %>%
  group_by(Cases) %>%
  summarise(Values = max(Values)) %>%
  mutate(Variable = 1 - Cases / varNumber)

# Combine with baseline
resu = rbind(resuT, resu)

# Apply optional transformation to lambda values 
lambdaMax = max(resu$Values)
lMax = max(resu$Values)
matR = rbind(cbind(lambda, 1), cbind(lambdaMax, 1))
coef = solve(matR) %*% rbind(log(0.5), log(1))  # (Log transform, but unused)

# Apply linear transformation to lambdas (here just identity)
tempL = resu$Values
coef = c(1, 0)
tempL = as.vector(tempL) * coef[1] + coef[2]
resu$Values = tempL

# Final result aggregation 
# For each variable and lambda, keep mean absolute value of coefficient
result = result %>%
  group_by(Variable, Lambda) %>%
  summarise(Values = mean(abs(Values)))


# 4.6| Create visualizations of variable retention
# Graph 1: Variable retention curve
# Plot dendrogram
rowLine    = 1
jumpOfRows = 18
colPlots   = 1
colTables  = 20

indicatorCode <- '2_2'
newSheet      <- addWorksheet(wb2, sheetName = indicatorCode)

base_exp     = 1
heightExp    = 1
widthExp     = 1
scale_factor = base_exp/widthExp

# Graph1
part          <- 'a'
labCodes      <- dataPlots %>% subset(Code == paste0(indicatorCode, part))

graph1=resu%>%
  ggplot(aes(x=Values, y=Variable))+ 
  geom_line()+
  geom_vline(xintercept = lambda, linetype = 'dashed', color = 'grey') +
  labs(title    = paste0(as.character(labCodes[, 5*(1-language)+2])),
       subtitle = as.character(labCodes[, 5*(1-language)+3]),
       caption  = as.character(labCodes[, 5*(1-language)+6]),
       x        = as.character(labCodes[, 5*(1-language)+4]),
       y        = as.character(labCodes[, 5*(1-language)+5])) +
  scale_y_continuous(expand = expansion(mult = c(0, 0)), limits = c(0,1),
                     labels = scales::percent_format(scale = 100, suffix = '%'), position = if (language == 0) {'right'} else {'left'}) +
  theme_classic() +
  theme(legend.position = 'bottom',
        text            = element_text(family = 'Georgia'),
        axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
        axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
        axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
        axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
        legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
        strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
        plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
        plot.subtitle   = element_text(size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
        legend.key.size = unit(0.5 * scale_factor, 'cm')) 

if (language == 0) {
  graph1 <- graph1 +
    scale_x_continuous(
      trans = trans_new(
        name      = "log-reverse",
        transform = function(x) -log(x),
        inverse   = function(x) exp(-x),
        domain    = c(1e-100, Inf)
      ),
      expand = expansion(mult = c(0.05, 0)),
      limits = c(lMax, lMin),
      labels = scales::percent_format(scale = 1000, suffix = '')
    )
} else {
  graph1 <- graph1 +
    scale_x_continuous(
      trans   = "log",
      expand  = expansion(mult = c(0, .05)),
      limits  = c(lMin, lMax),
      labels  = scales::percent_format(scale = 1000, suffix = '')
    )
}

# Save the graph
ggsave(fileName, plot = graph1, width = 12 * widthExp, height = 6 * heightExp * widthExp, scale = scale_factor)

# Graph 2: Coefficient paths
part          <- 'b'
labCodes      <- dataPlots %>% subset(Code == paste0(indicatorCode, part))

dataFT=result%>%filter(!Values==0 & !Variable%in%c("RMS","Intercep", NA))%>% # removed NA because here intercept is replaced by NA
  group_by(Variable)%>%
  summarise(Lambda=max(Lambda))%>%
  arrange(Lambda)

dataF=result  %>%
  filter(!Variable %in%c('RMS',"Intercep", NA))%>%
  mutate(Variable=factor(Variable,level=as.matrix(dataFT$Variable),label=as.matrix(dataFT$Variable)))%>%
  drop_na()

dataF=dataF%>%
  mutate(Variable=as.character(factor(Variable,
                                      levels = RankNames[,2],
                                      labels = RankNames[,4-language])))

colourCount = length(unique(dataF$Variable))
getPalette = colorRampPalette(brewer.pal(9, as.character(labCodes[, 12])))

graph2=dataF %>%
  ggplot(aes(x = Lambda, y = Values, color = Variable)) +
  geom_line(size = 1) +
  geom_vline(xintercept = lambda, linetype = 'dashed', color = 'grey') +
  geom_hline(yintercept = 0) +
  scale_y_continuous(expand = expansion(mult = c(.05, .05)),position = if (language == 0) {'right'} else {'left'}) +
  scale_colour_manual(values = getPalette(colourCount))+
  labs(title    = paste0(as.character(labCodes[, 5*(1-language)+2])),
       subtitle = as.character(labCodes[, 5*(1-language)+3]),
       caption  = as.character(labCodes[, 5*(1-language)+6]),
       x        = as.character(labCodes[, 5*(1-language)+4]),
       y        = as.character(labCodes[, 5*(1-language)+5])) +
  theme_classic() +
  theme(legend.position = 'bottom',
        text            = element_text(family = 'Georgia'),
        axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
        axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
        axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
        axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
        legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
        strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
        plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
        plot.subtitle   = element_text(size = 10, family = 'Georgia', hjust = if (language == 0) {1} else {0}),
        legend.key.size = unit(0.5 * scale_factor, 'cm')) +
  guides(color  = guide_legend(title = '',nrow=6,byrow=TRUE))

if (language == 0) {
  graph2 <- graph2 +
    scale_x_continuous(
      trans = trans_new(
        name      = "log-reverse",
        transform = function(x) -log(x),
        inverse   = function(x) exp(-x),
        domain    = c(1e-100, Inf)
      ),
      expand = expansion(mult = c(0.05, 0)),
      limits = c(lMax, lMin),
      labels = scales::percent_format(scale = 1000, suffix = '')
    )
} else {
  graph2 <- graph2 +
    scale_x_continuous(
      trans   = "log",
      expand  = expansion(mult = c(0, .05)),
      limits  = c(lMin, lMax),
      labels  = scales::percent_format(scale = 1000, suffix = '')
    )
}

# Save the graph
ggsave(fileName, plot = graph2, width = 12 * widthExp, height = 6 * heightExp * widthExp, scale = scale_factor)

# Merged graph
part          <- ''
base_exp     = 1
heightExp    = 2
widthExp     = 1
scale_factor = base_exp/widthExp


newPlot=ggarrange(graph1,                                                 
                  graph2, 
                  nrow = 2,
                  heights = c(1,2)
                  ) 


newData1=dataF%>%
  filter(Values>0)%>%
  group_by(Variable)%>%
  summarise(Lambda=max(Lambda))

newData2=dataF%>%
  filter(Values==0)%>%
  group_by(Variable)%>%
  summarise(Lambda=min(Lambda))
newData=rbind(newData1,newData2)%>%
  group_by(Variable)%>%
  summarise(Lambda=max(Lambda))%>%
  arrange(desc(Lambda))

# Export to excel the grapha and the table
fileName = paste0(fileName = paste0(outputLocation, indicatorCode, part, 
                                    if(language == 1) {' (English)'} else {' (Arab)'}, '.png'))

ggsave(fileName, plot = newPlot, width = 12 * widthExp, height = 6 * heightExp * widthExp, scale = scale_factor)
insertImage(wb2, file = fileName, sheet = indicatorCode, startRow = rowLine, startCol = colPlots, width = 12 * widthExp, height = 6 * heightExp * widthExp)
writeDataTable(wb2, sheet = indicatorCode, x = newData, startRow = rowLine, startCol = colTables)
rowLine = rowLine + jumpOfRows


output_file <- file.path(outputLocation, paste0("Targeting_assessment",    if(language == 1) {' (English)'} else {' (Arab)'}, ".xlsx"))

# Save the workbook
saveWorkbook(wb2, output_file, overwrite = TRUE)



# 5| Coverage Evaluation:Inclusion/ Exclusion errors ---------------------------------------------------------
# This section evaluates the coverage performance of the social protection program by:
#  Analyzing mismatches between actual and theoretical beneficiaries
#  Visualizing inclusion/exclusion errors across demographic groups


# 5.1| Specify output location for coverage evaluation results
outputLocation = paste0(userLocation, '/Output/Coverage Evaluation/')
source(paste0(scriptLocation, 'Aux - Functions.R'), encoding = 'UTF-8')

# 5.2| Load inclusion/exclusion labels from dictionary
# Retrieve labels for inclusion/exclusion categories from Excel
IncexcLabels <- as.matrix(read_excel(paste0(inputLocation, 'Dictionary/Dictionary.xlsx'), sheet = 'Category', range = 'Z14:AC16'))

# 5.3| Data preparation 
# 5.3.1| Combine vulnerability indicators with scoring and household head data
# Create dataset with demographics + eligibility data for coverage evaluation
completeData_Vuln = Indicators %>% 
  left_join(t1) %>%
  left_join(SCORING) %>%
  left_join(filtered_Head) 



# Define theoretical beneficiaries using cut-off
# Create final dataset with beneficiary information
finalData = completeData_Vuln %>% 
  left_join(Beneficiaries_Data) %>% 
  dplyr::select(c("HHID", "Governorate", "Location","score", "Gender", "Marital status", "Schooling level", "Employment Status", "Beneficiaries"), starts_with("var_")) 

# prepare eligibility data
Eligibles = HH_Data %>% 
  left_join(Beneficiaries_Data %>% dplyr::select(c('HHID','Beneficiaries'))) %>%
  group_by(HHID)

# First Method
# 5.4| Validate beneficiaries based on score threshold (40) ⚠️

finalData= finalData %>% 
  mutate(
    theo_beneficiary = if_else(score>= 40, "Yes", "No")
  )

# 5.4.1| Mismatch analysis between actual and theoretical beneficiaries

temp = finalData %>% 
  group_by(Beneficiaries, theo_beneficiary) %>% 
  summarise(cases = length(HHID)) %>% 
  drop_na()


# 5.4.2| Create plots for comparing real vs theoretical beneficiaries by main variables
wb3 = openxlsx::createWorkbook(creator = 'ESCWA')


#Plot results
rowLine    = 1
jumpOfRows = 20
colPlots   = 1
colTables  = 12

indicatorCode <- '3_1'
part          <- ''
labCodes      <- dataPlots %>% subset(Code == paste0(indicatorCode, part))
newSheet      <- addWorksheet(wb3, sheetName = indicatorCode)

# Plot dendrogram
# To export:
base_exp     = 1
heightExp    = 1
widthExp     = 1
scale_factor = base_exp/widthExp

create_and_export_plot <- function(data, facet_var, part_code, indicator_code, rowLine, wb3, outputLocation, language, colPlots = 1, colTables = 12) {
  # Summarise data
  temp <- data %>% 
    group_by(Beneficiaries, theo_beneficiary, !!sym(facet_var)) %>% 
    summarise(cases = n(), .groups = "drop") %>% 
    drop_na()
  
  # Labels
  labCodes <- dataPlots %>% filter(Code == paste0(indicator_code, part_code))
  
  # Plot parameters
  base_exp <- 1
  widthExp <- 1
  heightExp <- if (facet_var == "Schooling level") 1.5 else 1
  scale_factor <- base_exp / widthExp
  jumpOfRows <- if (facet_var == "Schooling level") 30 else 20
  
  # Plots real vs theoretical beneficiaries
  newPlot <- temp %>%
    ggplot(aes(x = theo_beneficiary, y = cases, fill = Beneficiaries)) +
    facet_wrap(as.formula(paste0("~`", facet_var, "`")), ncol = 2) +
    geom_bar(position = "fill", stat = "identity", color = 'black') +
    labs(
      title    = as.character(labCodes[, 5*(1 - language) + 2]),
      subtitle = as.character(labCodes[, 5*(1 - language) + 3]),
      caption  = as.character(labCodes[, 5*(1 - language) + 6]),
      x        = as.character(labCodes[, 5*(1 - language) + 4]),
      y        = as.character(labCodes[, 5*(1 - language) + 5])
    ) +
    scale_fill_brewer(palette = as.character(labCodes[, 12])) +
    scale_x_discrete(position = if (language == 0) "top" else "bottom") +
    coord_flip() +
    theme_classic() +
    theme(
      legend.position = 'bottom',
      text = element_text(family = 'Georgia'),
      axis.text = element_text(size = scale_factor * 10, family = 'Georgia'),
      axis.title = element_text(size = scale_factor * 10, family = 'Georgia'),
      legend.text = element_text(size = scale_factor * 10, family = 'Georgia'),
      strip.text = element_text(size = scale_factor * 10, family = 'Georgia'),
      plot.title = element_text(face = 'bold', size = 10, family = 'Georgia', hjust = if (language == 0) 1 else 0),
      plot.subtitle = element_text(size = 10, family = 'Georgia', hjust = if (language == 0) 1 else 0),
      legend.key.size = unit(0.5 * scale_factor, 'cm')
    ) +
    guides(fill = guide_legend(title = ''))
  

  # Ensure numeric x values
  resu$Values <- as.numeric(resu$Values)
  
  # Define reversed log transformation
  log_rev_trans <- trans_new(
    name      = "log-reverse",
    transform = function(x) -log(x),
    inverse   = function(x) exp(-x),
    domain    = c(1e-100, Inf)
  )
  
  if (language == 0) {
    graph1 <- graph1 +
      scale_x_continuous(
        trans   = log_rev_trans,
        expand  = expansion(mult = c(0.05, 0)),
        limits  = c(lMax, lMin),
        labels  = scales::percent_format(scale = 1000, suffix = '')
      )
  } else {
    graph1 <- graph1 +
      scale_x_continuous(
        trans   = "log",
        expand  = expansion(mult = c(0, .05)),
        limits  = c(lMin, lMax),
        labels  = scales::percent_format(scale = 1000, suffix = '')
      )
  }
  
  # Prepare reshaped data for Excel export
  newData <- temp %>% spread(!!sym(facet_var), cases)
  
  # Export plot
  fileName <- paste0(outputLocation, indicator_code, part_code, 
                     if (language == 1) " (English)" else " (Arab)", ".png")
  ggsave(fileName, plot = newPlot, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  
  # Insert into Excel
  insertImage(wb3, file = fileName, sheet = indicator_code, startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb3, sheet = indicator_code, x = newData, startRow = rowLine, startCol = colTables)
  
  return(rowLine + jumpOfRows)
}


# Variable list with associated part codes
facet_vars <- list(
  "Gender"             = "a",
  "Schooling level"    = "b",
  "Employment Status"  = "c",
  "Marital status"     = "d"
)

# Apply the function and plot.
for (v in names(facet_vars)) {
  part_code <- facet_vars[[v]]
  rowLine <- create_and_export_plot(finalData, facet_var = v, part_code = part_code, 
                                    indicator_code = indicatorCode, rowLine = rowLine,
                                    wb = wb3, outputLocation = outputLocation, language = language)
}


# 5.4.3| Mapping for comparing real vs theoretical beneficiaries under  threshold selected beneficairies.
# Process both inclusion and exclusion in one pipeline
error_metrics <- finalData %>% 
  drop_na(Beneficiaries, theo_beneficiary, Governorate) %>%
  group_by(Governorate, Beneficiaries, theo_beneficiary) %>% 
  summarise(cases = n(), .groups = "drop") %>%
  group_by(Governorate) %>%
  summarise(
    # Inclusion error metrics
    incl_cases = sum(cases[theo_beneficiary == "No" & Beneficiaries == "Yes"]),
    benefiting = sum(cases[Beneficiaries == "Yes"]),
    
    # Exclusion error metrics
    excl_cases = sum(cases[theo_beneficiary == "Yes" & Beneficiaries == "No"]),
    theo_benefiting = sum(cases[theo_beneficiary == "Yes"]),
    
    .groups = "drop"
  ) %>%
  # Compute error rates as percentages of relevant totals
  mutate(
    incl_error = incl_cases/benefiting,
    exc_error = excl_cases/theo_benefiting
  ) %>%
  # Keep only relevant columns for mapping
  dplyr::select(Governorate, incl_error, exc_error)

# Split into separate datasets
temp_inc_final <- error_metrics %>% dplyr::select(Governorate, incl_error)
temp_exc_final <- error_metrics %>% dplyr::select(Governorate, exc_error)

# 5.4.4|  GIS Data Processing - Optimized 

# Load and prepare shapefile of governorates with labels in selected language
NEMEY <- st_read(paste0(inputLocation, 'GIS/Nemey.shp')) %>%
  mutate(
    Governorate = factor(NAME_1, levels = GovernorateLabels[,2], labels = GovernorateLabels[,2]),
    textNames = Governorate
  )

# Create bounding box matrix directly
b <- matrix(st_bbox(NEMEY), nrow = 2)

# Generate bounding box and centroids for map labeling
world_points <- st_centroid(NEMEY$geometry) %>%
  st_coordinates() %>%
  as_tibble() %>%
  mutate(
    Names = NEMEY$textNames,
    Label = NEMEY$Governorate
  ) %>%
  rename(X = 1, Y = 2)

# Define dynamic color bins and palette for mapping
nBins <- max(10, length(unique(NEMEY$Governorate))) # Dynamic bin count
mycolors1 <- colorRampPalette(brewer.pal(8, "YlGnBu"))(nBins)
colors_map <- c("#f7fbff", "#deebf7", "#c6dbef", "#9ecae1", "#6baed6", "#4292c6", "#2171b5", "#08519c", "#08306b")

# 5.4.5| Map Generation - Parameterized Function 

generate_error_map <- function(part, error_data, error_var, wb3, rowLine) {
  # Get labels
  labCodes <- dataPlots %>% filter(Code == paste0(indicatorCode, part))
  
  # Prepare map data
  map_data <- NEMEY %>%
    left_join(error_data) %>%
    mutate(!!error_var := if_else(is.na(get(error_var)), 0, get(error_var)))
  
  # Plot polygons of governorates, filled by selected error variable
  # Create plot
  p <- ggplot() +
    
    geom_sf(data = map_data, aes_string(fill = error_var), 
            colour = 'black', inherit.aes = FALSE, size = 0.3) +
    geom_text(data = world_points, 
              aes(x = X, y = Y, label = Label),
              size = 3, color = "black", fontface = "bold",
              check_overlap = TRUE) +
    
    coord_sf(xlim = c(b[1,1]-0.25, b[1,2]+0.25), 
             ylim = c(b[2,1]-0.25, b[2,2]+0.25), 
             expand = FALSE) +
    
    scale_fill_gradientn(colours = colors_map, name = NULL, limits = c(0, NA), 
                         labels = scales::percent_format(scale = 100, accuracy = 1, suffix = "%")) +
    labs(
      title = as.character(labCodes[, 5*(1-language)+2]),
      subtitle = as.character(labCodes[, 5*(1-language)+3]),
      caption = as.character(labCodes[, 5*(1-language)+6]),
      x = "", y = ""
    ) +
    theme_classic() +
    theme(
      plot.title = element_text(face = 'bold', size = 10, family = 'Georgia', 
                                hjust = ifelse(language == 0, 1, 0)),
      plot.subtitle = element_text(size = 10, family = 'Georgia', 
                                   hjust = ifelse(language == 0, 1, 0))
    )
  
  # Prepare export data
  export_data <- error_data %>%
    ungroup() %>%
    mutate(Governorate = factor(Governorate, 
                                levels = GovernorateLabels[,2], 
                                labels = GovernorateLabels[,4-language]))
  
  # Export files
  fileName <- paste0(outputLocation, indicatorCode, part, 
                     ifelse(language == 1, ' (English)', ' (Arab)'), '.png')
  
  ggsave(fileName, plot = p, 
         width = 6 * widthExp, height = 4 * heightExp * widthExp, 
         scale = scale_factor)
  
  insertImage(wb3, fileName, sheet = indicatorCode, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  
  writeDataTable(wb3, sheet = indicatorCode, x = export_data, 
                 startRow = rowLine, startCol = colTables)
  
  return(rowLine + jumpOfRows)
}


# Initialize parameters
base_exp <- 1
heightExp <- 1
widthExp <- 1
scale_factor <- base_exp/widthExp
jumpOfRows <- 20

# Generate maps sequentially
rowLine <- generate_error_map('a1', temp_exc_final, 'exc_error', wb3, rowLine)
rowLine <- generate_error_map('a2', temp_inc_final, 'incl_error', wb3, rowLine)


#Save the first technique in the first sheet of the workbook
output_file <- file.path(outputLocation, paste0("Inclusion_exclusion_cagetories",  ".xlsx"))
saveWorkbook(wb3, output_file, overwrite = TRUE)


# Second Method
# 5.5| CONVERAGE EVALUATION WHEN BENEFICIARIES SELECTION  BASED ON RANKING AND FIXED AMOUNT OF HH

rowLine    = 1
jumpOfRows = 20
colPlots   = 1
colTables  = 12

indicatorCode <- '3_2'
part          <- ''
labCodes      <- dataPlots %>% subset(Code == paste0(indicatorCode, part))
newSheet      <- addWorksheet(wb3, sheetName = indicatorCode)

# Plot dendrogram
# To export:
base_exp     = 1
heightExp    = 1
widthExp     = 1
scale_factor = base_exp/widthExp

# Prepare data
data_scores <- Beneficiaries_Data %>%
  left_join(SCORING, by = "HHID") %>%
  dplyr::select(HHID, Beneficiaries, score) %>%
  drop_na(score)


# Add disaggregates
data_scores_disagg <- data_scores %>%
  left_join(filtered_Head %>% dplyr::select(HHID, Gender, `Marital status`, `Schooling level`, `Employment Status`),
            by = "HHID")

# Variables to disaggregate by
disagg_vars <- c("Gender", "Schooling level", "Employment Status", "Marital status")

# Generate and save plots to Excel workbook

# Distribution plots by aggregates (Gender, Schooling, Employment, Marital status)
for (var in disagg_vars) {
  p <- ggplot(data_scores_disagg, aes(x = score, fill = Beneficiaries)) +
    geom_density(alpha = 0.4) +
    facet_wrap(vars(!!sym(var)), scales = "free") +
    labs(title = paste("Score Distribution by", var), x = "Score", y = "Density") +
    theme_minimal()
  
  # Save plot to a temporary image file
  
  safe_var_name <- gsub(" ", "_", var)
  
  fileName <- paste0(outputLocation, indicatorCode,  part,  "rank",
                     safe_var_name, ifelse(language == 1, '_English', '_Arab'), '.png')
  
  
  ggsave(fileName, plot = p, 
         width = 10, height = 6, dpi = 300)
  
  
  # Add image to Excel
  # Insert into Excel
  insertImage(wb3, file = fileName, sheet = indicatorCode, 
              startRow = rowLine, startCol = colPlots,
              width = 10, height = 6, units = "in")
  
  tableData <- data_scores_disagg %>%
    group_by(!!sym(var), Beneficiaries) %>%
    summarise(count = n(), .groups = "drop") %>%
    tidyr::spread(key = Beneficiaries, value = count, fill = 0)
  
  # 4. Insert table into Excel
  writeDataTable(wb3, sheet = indicatorCode, x = tableData, 
                 startRow = rowLine, startCol = colTables)
  
  # Move rowLine down for the next plot
  rowLine <- rowLine + 40
  
  #  safe_var_name <- gsub(" ", "_", var)
  # addWorksheet(wb3, sheetName = safe_var_name)
  # insertImage(wb3, sheet = safe_var_name, file = tmp_file, width = 10, height = 6, units = "in")
  # insertImage(wb3, file = fileName, sheet = indicatorCode, startRow = rowLine, startCol = colPlots, 
  #             width = 6 * widthExp, height = 4 * heightExp * widthExp)
  # writeDataTable(wb3, sheet = indicatorCode, x = newData, startRow = rowLine, startCol = colTables)
  # 
}

#  Create Maps
# Process both inclusion and exclusion in one pipeline
error_metrics <- finalData %>%
  drop_na(Beneficiaries, Governorate) %>%
  group_by(Governorate) %>%
  summarise(
    # Define thresholds
    min_inc = min(score[Beneficiaries == "Yes"], na.rm = TRUE),
    max_excl = max(score[Beneficiaries == "No"], na.rm = TRUE),
    
    # Inclusion error: beneficiaries with score below max score of non-beneficiaries
    incl_cases = sum(Beneficiaries == "Yes" & score < max_excl, na.rm = TRUE),
    benefiting = sum(Beneficiaries == "Yes", na.rm = TRUE),
    
    # Exclusion error: non-beneficiaries with score above min score of beneficiaries
    excl_cases = sum(Beneficiaries == "No" & score > min_inc, na.rm = TRUE),
    no_benefiting = sum(Beneficiaries == "No", na.rm = TRUE),
    
    .groups = "drop"
  ) %>%
  mutate(
    incl_error = incl_cases / benefiting,
    exc_error = excl_cases / no_benefiting
  ) %>%
  dplyr::select(Governorate, incl_error, exc_error)


# Split into separate datasets
temp_inc_final <- error_metrics %>% dplyr::select(Governorate, incl_error)
temp_exc_final <- error_metrics %>% dplyr::select(Governorate, exc_error)

#  GIS Data Processing 
NEMEY <- st_read(paste0(inputLocation, 'GIS/Nemey.shp')) %>%
  mutate(
    Governorate = factor(NAME_1, levels = GovernorateLabels[,2], labels = GovernorateLabels[,2]),
    textNames = Governorate
  )

# Create bounding box matrix directly
b <- matrix(st_bbox(NEMEY), nrow = 2)

# Create centroids data in one step
world_points <- st_centroid(NEMEY$geometry) %>%
  st_coordinates() %>%
  as_tibble() %>%
  mutate(
    Names = NEMEY$textNames,
    Label = NEMEY$Governorate
  ) %>%
  rename(X = 1, Y = 2)

# Define dynamic color bins and palette for mapping
nBins <- max(10, length(unique(NEMEY$Governorate))) # Dynamic bin count
mycolors1 <- colorRampPalette(brewer.pal(8, "YlGnBu"))(nBins)
colors_map <- c("#ffffe5", "#fff7bc", "#fee391",
                "#fec44f", "#fe9929", "#ec7014",
                "#cc4c02", "#993404", "#662506")

# 5.5.3| Map Generation - Parameterized Function 

generate_error_map <- function(part, error_data, error_var, wb3, rowLine) {
  # Get labels
  labCodes <- dataPlots %>% filter(Code == paste0(indicatorCode, part))
  
  # Prepare map data
  map_data <- NEMEY %>%
    left_join(error_data) %>%
    mutate(!!error_var := if_else(is.na(get(error_var)), 0, get(error_var)))
  
  # Create plot
  p <- ggplot() +
    
    geom_sf(data = map_data, aes_string(fill = error_var), 
            colour = 'black', inherit.aes = FALSE, size = 0.3) +
    geom_text(data = world_points, 
              aes(x = X, y = Y, label = Label),
              size = 3, color = "black", fontface = "bold",
              check_overlap = TRUE) +
    
    coord_sf(xlim = c(b[1,1]-0.25, b[1,2]+0.25), 
             ylim = c(b[2,1]-0.25, b[2,2]+0.25), 
             expand = FALSE) +
    
    scale_fill_gradientn(colours = colors_map, name = NULL, limits = c(0, NA), 
                         labels = scales::percent_format(scale = 100, accuracy = 1, suffix = "%")) +
    labs(
      title = as.character(labCodes[, 5*(1-language)+2]),
      subtitle = as.character(labCodes[, 5*(1-language)+3]),
      caption = as.character(labCodes[, 5*(1-language)+6]),
      x = "", y = ""
    ) +
    theme_classic() +
    theme(
      plot.title = element_text(face = 'bold', size = 10, family = 'Georgia', 
                                hjust = ifelse(language == 0, 1, 0)),
      plot.subtitle = element_text(size = 10, family = 'Georgia', 
                                   hjust = ifelse(language == 0, 1, 0))
    )
  
  # Prepare export data
  export_data <- error_metrics  %>%
    ungroup() %>%
    mutate(Governorate = factor(Governorate, 
                                levels = GovernorateLabels[,2], 
                                labels = GovernorateLabels[,4-language]))
  
  # Export files
  fileName <- paste0(outputLocation, indicatorCode, part,
                     ifelse(language == 1, ' (English)', ' (Arab)'), '.png')
  
  ggsave(fileName, plot = p, 
         width = 6 * widthExp, height = 4 * heightExp * widthExp, 
         scale = scale_factor)
  
  insertImage(wb3, fileName, sheet = indicatorCode, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  
  writeDataTable(wb3, sheet = paste0(indicatorCode), x = export_data, 
                 startRow = rowLine, startCol = colTables)
  
  return(rowLine + jumpOfRows)
}



# 5.5.4| Execute Map Generation -------------------------------------------

# Initialize parameters
base_exp <- 1
heightExp <- 1
widthExp <- 1
scale_factor <- base_exp/widthExp
jumpOfRows <- 20

# Generate maps sequentially
rowLine <- generate_error_map('a1', temp_exc_final, 'exc_error', wb3, rowLine)
rowLine <- generate_error_map('a2', temp_inc_final, 'incl_error', wb3, rowLine)


output_file <- file.path(outputLocation, paste0("Inclusion_exclusion_cagetories",  ".xlsx"))

# Save the workbook
saveWorkbook(wb3, output_file, overwrite = TRUE)


# -----------------------------------------------------------
# Annex Missing values 
# -----------------------------------------------------------

PMTIndicators1=PMTIndicators # to create
PMTIndicators2=PMTIndicators # to update

var_types <- list(
  discrete = c(
    "var_overcrowdingIssue", "var_Wall_type", "var_Floor_type", 
    "var_Toilet_Type", "var_Drinking_water", "var_sewage",
    "var_AgriLand", "var_chronic", "var_duration_hospital",
    "var_duration_school", "var_NEET"
  ),
  continuous = c(
    "var_Property_score", "var_common", "var_measels",
    "var_tot_disab", "var_other_disab", "var_enrolment",
    "var_high_education_rate", "var_unemployment_rate",
    "var_child_labour", "var_female_ratio", 
    "var_young_dependency_ratio"
  )
)

predict_variable <- function(var_name, data, new_data) {
  # Remove HHID and the target variable from predictors
  predictors <- data %>% dplyr::select(-HHID, -all_of(var_name))
  
  if (var_name %in% var_types$discrete) {
    # For discrete variables - classification tree
    formula <- as.formula(paste("factor(", var_name, ") ~ ."))
    tree <- rpart(formula, data = data %>% dplyr::select(-HHID), 
                  method = "class")
    pred <- as.numeric(as.character(predict(tree, new_data, type = "class")))
  } else {
    # For continuous variables - regression tree
    formula <- as.formula(paste(var_name, "~ ."))
    tree <- rpart(formula, data = data %>% dplyr::select(-HHID), 
                  method = "anova")
    pred <- predict(tree, new_data)
  }
  
  return(pred)
}


for (var in c(var_types$discrete, var_types$continuous)) {
  PMTIndicators2[[var]] <- predict_variable(var, PMTIndicators1, PMTIndicators1)
}

SCORING2 = PMTIndicators2 %>% 
  mutate(
    s_var_overcrowdingIssue=(1/6)*(1/7)*var_overcrowdingIssue, # the value next to each variable is its weight. in this case we specified the weight but in other cases we have to stick to the weight given
    s_var_Wall=(1/6)*(1/7)*var_Wall_type,
    s_var_Floor=(1/6)*(1/7)*var_Floor_type,
    s_var_Toilet=(1/6)*(1/7)*var_Toilet_Type,
    s_var_DrinkWater=(1/6)*(1/7)*var_Drinking_water,
    s_var_Sewage=(1/6)*(1/7)*var_sewage,
    
    s_var_Property=(1/2)*(1/7)*var_Property_score,
    s_var_agriLand=(1/2)*(1/7)*var_AgriLand,
    
    s_var_chronic=(1/6)*(1/7)*var_chronic,
    s_var_common=(1/6)*(1/7)*var_common,
    s_var_measels=(1/6)*(1/7)*var_measels,
    s_var_totdisab=(1/6)*(1/7)*var_tot_disab,
    s_var_otherdisab=(1/6)*(1/7)*var_other_disab,
    s_var_durationHospital=(1/6)*(1/7)*var_duration_hospital,
    
    s_var_enrolment=(1/3)*(1/7)*var_enrolment,
    s_var_higheducation=(1/3)*(1/7)*var_high_education_rate,
    s_var_durationSchool=(1/3)*(1/7)*var_duration_school,
    
    s_var_NEET=(1/3)*(1/7)*var_NEET,
    s_var_unemp=(1/3)*(1/7)*var_unemployment_rate,
    s_var_childLabour=(1/3)*(1/7)*var_child_labour,
    
    s_var_female=(1/2)*(1/7)*var_female_ratio,
    s_var_youngdependency=(1/2)*(1/7)*var_young_dependency_ratio,
    
    score2 = s_var_overcrowdingIssue  + s_var_Wall + s_var_Floor +  s_var_Toilet + s_var_DrinkWater + s_var_Sewage + s_var_Property + s_var_agriLand + s_var_chronic + s_var_common + s_var_measels + s_var_totdisab + s_var_otherdisab +
      s_var_durationHospital + s_var_enrolment+ s_var_higheducation + s_var_durationSchool + s_var_NEET +  s_var_unemp + s_var_childLabour +s_var_female +  s_var_youngdependency )%>%
  drop_na()






