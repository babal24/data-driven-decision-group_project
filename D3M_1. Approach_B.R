#===============================================================================
# Assignment 2 - Data-Driven Decision-Making  
# MSc Business Analytics
# Queen´s University Belfast
#-------------------------------------------------------------------------------
# Vasileios Gounaris-Babaletsos - 40314803
# Anurag Karunakaran            - 40318784
# Marcel Heinz Loré             - 40316722
# Elina Sagmanova               - 40307813
# Hongjie Wei                   - 40272516
#-------------------------------------------------------------------------------

# setting the working directory

setwd("C:/Users/marce/OneDrive/Dokumente/Belfast/Data-Driven Decision-Making/Ass. 2/Coding")

# check working directory

getwd()

# setting system language in english (previously german)

Sys.setenv(LANG = "en")

# loading libraries

library(readxl)
library(ompr)
library(ompr.roi)
library(ROI)
library(ROI.plugin.glpk)
library(stats)
library(dplyr)
library(writexl)
library(ggplot2)

#-------------------------------------------------------------------------------
# LOADING THE DATA

# for the solution of this problem, the EXCEL-file from CANVAS was prepared 
# and supplemented with additional columns
# the EXCEL-file will be uploaded in CANVAS too
# load the pre-processed Overveld-EXCEL-File

D3M <- read_xlsx("D3M_Data_Ass_2_NEW.xlsx") 

# Additionally, the lead times were written out in an EXCEL-File
# load the lead times of the process

LEAD<- read_xlsx("LEAD_TIMES.xlsx")


#-------------------------------------------------------------------------------
# Initializing the variables needed for the o- and p- sessions

# 3 patients can be treated in an o-session: 

o_sessions <-3

# 16 patients can be treated in a p-session: 

p_sessions <-16

# Penalty for delayed treatment:

penalty_COST_O <- 20
penalty_COST_P1 <- 3
penalty_COST_P2 <- 2
penalty_COST_P3 <- 1

#-------------------------------------------------------------------------------
# INITIALIZING THE TABLE

# Setting up a Loop for the whole year (52 weeks) --> week = i

for (i in 1:52){

  # Idea for a first approach: Split the capacity for each week in 2/3 o-session 
  # and 1/3 p-sessions.
  # O-sessions are prioritized because the penalty for delayed treatment
  # is much higher compared to P-sessions.
  
  D3M$os_2[i] <- D3M$capacity[i] / 3 * 2
  D3M$ps_2[i] <- D3M$capacity[i] / 3
  
  # Copy the scheduled session of o.s.(-2) and p.s.(-2) in o.s.(0) and p.s.(0)
  # o.s.(-2) = os_2
  # p.s.(-2) = ps_2
  # o.s.(0) = os_0
  # p.s.(0) = ps_0
  
  D3M$os_0[i] <- D3M$os_2[i]
  D3M$ps_0[i] <- D3M$ps_2[i]

  #-----------------------------------------------------------------------------
  # Setting up the function for the table:
  # creating the cost function:
  D3M$Cost[i] <- D3M$os_2[i] * 100 + D3M$ps_2[i] * 50 + (D3M$os_2[i] - D3M$os_0[i])*(-50) + (D3M$ps_2[i] - D3M$ps_0[i])*(-30) 
  + D3M$Penalty_O[i] + D3M$Penalty_P1[i] + D3M$Penalty_P2[i] + D3M$Penalty_P3[i]
  
  # calculating the treated (O) patients for each week: 
  D3M$Treated_O[i] <- min(D3M$O[i], D3M$os_0[i] * o_sessions)
  
  # calculating the Delayed (O) patients for each week:
  D3M$Delayed_O[i] <- D3M$O[i] - D3M$Treated_O[i]
  
  # calculating the Treated (P1) patients for each week:
  D3M$Treated_P1[i] <- min(D3M$P1[i], D3M$ps_0[i] * p_sessions)
  
  # calculating the Delayed (P1) patients for each week:
  D3M$Delayed_P1[i] <- D3M$P1[i] - D3M$Treated_P1[i]
  
  # calculating the Treated (P2) patients for each week:
  D3M$Treated_P2[i] <- min(D3M$P2[i], D3M$ps_0[i] * p_sessions - D3M$Treated_P1[i])
  
  # calculating the Delayed (P2) patients for each week:
  D3M$Delayed_P2[i] <- D3M$P2[i] - D3M$Treated_P2[i]
  
  # calculating the Treated (P3) patients for each week:
  D3M$Treated_P3[i] <- min(D3M$P3[i], D3M$ps_0[i] * p_sessions - D3M$Treated_P1[i] - D3M$Treated_P2[i])
  
  # calculating the Delayed (P3) patients for each week:
  D3M$Delayed_P3[i] <- D3M$P3[i] - D3M$Treated_P3[i]
  }

#===============================================================================
# Creating a second loop to implement the decisions for the current week and lead times 
# updating the columns according to the number of patients and lead times

for (j in 1:52){
  # Updating Treated and Delayed Patients
   D3M$Treated_O[j] <- min(D3M$O[j], D3M$os_0[j] * o_sessions)
   D3M$Delayed_O[j] <- D3M$O[j] - D3M$Treated_O[j]
   D3M$Treated_P1[j] <- min(D3M$P1[j], D3M$ps_0[j] * p_sessions)
   D3M$Delayed_P1[j] <- D3M$P1[j] - D3M$Treated_P1[j]
   D3M$Treated_P2[j] <- min(D3M$P2[j], D3M$ps_0[j] * p_sessions - D3M$Treated_P1[j])
   D3M$Delayed_P2[j] <- D3M$P2[j] - D3M$Treated_P2[j]
   D3M$Treated_P3[j] <- min(D3M$P3[j], D3M$ps_0[j] * p_sessions - D3M$Treated_P1[j] - D3M$Treated_P2[j])
   D3M$Delayed_P3[j] <- D3M$P3[j] - D3M$Treated_P3[j]
   

   # Implementing a decision rule to optimize the session for the current week:
   # Implement "earlier treatment of patients", 
   
   # Execute Operations earlier. That means treating patients of the next week
   # one week earlier.
   
   if(D3M$os_0[j] * o_sessions > D3M$O[j] ){
     redundant_os= (D3M$os_0 [j] * o_sessions) - D3M$O[j]
     D3M$Treated_O[j]<- D3M$Treated_O[j] + redundant_os
     if(redundant_os< D3M$O[j+1]){
       D3M$O[j+1] <- D3M$O[j+1] - redundant_os
     }else{
       D3M$O[j+1] <- D3M$O[j+1] - D3M$O[j+1]
     }
   }
   
   # Execute P1-sessions earlier. That means treating patients of the next week
   # one week earlier
   # only P1-sessions, because P1 sessions are prioritized compared to P2 and P3
   # P1, P2 and P3 use all the same pool of p-sessions

    if(D3M$ps_0[j] * p_sessions > (D3M$P1[j] + D3M$P2[j] + D3M$P3[j]) ){
      redundant_ps = (D3M$ps_0[j] * p_sessions) - (D3M$P1[j] + D3M$P2[j] + D3M$P3[j])
      D3M$Treated_P1[j]<- D3M$Treated_P1[j] + redundant_ps
      if(redundant_ps< D3M$P1[j+1]){
      D3M$P1[j+1] <- D3M$P1[j+1] - redundant_ps
      }else{
        D3M$P1[j+1] <- D3M$P1[j+1] - D3M$P1[j+1]
        
      }

    }
   

   #----------------------------------------------------------------------------
   # Calculate O, P2 and P3 for the next eight weeks using the lead times (EXCEL-FILE "LEAD")
  # LEAD TIME 1: 
  D3M$O[j+1]<- D3M$O[j+1]  + D3M$Delayed_O[j]  + LEAD$P1_O[1]  * D3M$Treated_P1[j] + LEAD$P2_O[1]  * D3M$Treated_P2[j]
  D3M$P2[j+1]<-D3M$P2[j+1] + D3M$Delayed_P2[j] + LEAD$P1_P2[1] * D3M$Treated_P1[j] + LEAD$O_P2[1]  * D3M$Treated_O[j]
  D3M$P3[j+1]<-D3M$P3[j+1] + D3M$Delayed_P3[j] + LEAD$P2_P3[1] * D3M$Treated_P2[j] + LEAD$P3_P3[1] * D3M$Treated_P3[j]
  
  # LEAD TIME 2:
  D3M$O[j+2] <- D3M$O[j+2] + D3M$Delayed_O[j+1]  + LEAD$P1_O[2]  * D3M$Treated_P1[j] + LEAD$P2_O[2]  * D3M$Treated_P2[j]
  D3M$P2[j+2]<-D3M$P2[j+2] + D3M$Delayed_P2[j+1] + LEAD$P1_P2[2] * D3M$Treated_P1[j] + LEAD$O_P2[2]  * D3M$Treated_O[j]
  D3M$P3[j+2]<-D3M$P3[j+2] + D3M$Delayed_P3[j+1] + LEAD$P2_P3[2] * D3M$Treated_P2[j] + LEAD$P3_P3[2] * D3M$Treated_P3[j]

  # LEAD TIME 3:
  D3M$O[j+3] <-D3M$O[j+3]  + D3M$Delayed_O[j+2]  + LEAD$P1_O[3]  * D3M$Treated_P1[j] + LEAD$P2_O[3]  * D3M$Treated_P2[j]
  D3M$P2[j+3]<-D3M$P2[j+3] + D3M$Delayed_P2[j+2] + LEAD$P1_P2[3] * D3M$Treated_P1[j] + LEAD$O_P2[3]  * D3M$Treated_O[j]
  D3M$P3[j+3]<-D3M$P3[j+3] + D3M$Delayed_P3[j+2] + LEAD$P2_P3[3] * D3M$Treated_P2[j] + LEAD$P3_P3[3] * D3M$Treated_P3[j]
  
  # LEAD TIME 4:
  D3M$O[j+4] <-D3M$O[j+4]  + D3M$Delayed_O[j+3]  + LEAD$P1_O[4]  * D3M$Treated_P1[j] + LEAD$P2_O[4]  * D3M$Treated_P2[j]
  D3M$P2[j+4]<-D3M$P2[j+4] + D3M$Delayed_P2[j+3] + LEAD$P1_P2[4] * D3M$Treated_P1[j] + LEAD$O_P2[4]  * D3M$Treated_O[j]
  D3M$P3[j+4]<-D3M$P3[j+4] + D3M$Delayed_P3[j+3] + LEAD$P2_P3[4] * D3M$Treated_P2[j] + LEAD$P3_P3[4] * D3M$Treated_P3[j]
  
  # LEAD TIME 5:
  D3M$O[j+5] <-D3M$O[j+5]  + D3M$Delayed_O[j+4]  + LEAD$P1_O[5]  * D3M$Treated_P1[j] + LEAD$P2_O[5]  * D3M$Treated_P2[j]
  D3M$P2[j+5]<-D3M$P2[j+5] + D3M$Delayed_P2[j+4] + LEAD$P1_P2[5] * D3M$Treated_P1[j] + LEAD$O_P2[5]  * D3M$Treated_O[j]
  D3M$P3[j+5]<-D3M$P3[j+5] + D3M$Delayed_P3[j+4] + LEAD$P2_P3[5] * D3M$Treated_P2[j] + LEAD$P3_P3[5] * D3M$Treated_P3[j]
  
  # LEAD TIME 6:
  D3M$O[j+6] <-D3M$O[j+6]  + D3M$Delayed_O[j+5]  + LEAD$P1_O[6]  * D3M$Treated_P1[j] + LEAD$P2_O[6]  * D3M$Treated_P2[j]
  D3M$P2[j+6]<-D3M$P2[j+6] + D3M$Delayed_P2[j+5] + LEAD$P1_P2[6] * D3M$Treated_P1[j] + LEAD$O_P2[6]  * D3M$Treated_O[j]
  D3M$P3[j+6]<-D3M$P3[j+6] + D3M$Delayed_P3[j+5] + LEAD$P2_P3[6] * D3M$Treated_P2[j] + LEAD$P3_P3[6] * D3M$Treated_P3[j]
  
  # LEAD TIME 7:
  D3M$O[j+7] <-D3M$O[j+7]  + D3M$Delayed_O[j+6]  + LEAD$P1_O[7]  * D3M$Treated_P1[j] + LEAD$P2_O[7]  * D3M$Treated_P2[j]
  D3M$P2[j+7]<-D3M$P2[j+7] + D3M$Delayed_P2[j+6] + LEAD$P1_P2[7] * D3M$Treated_P1[j] + LEAD$O_P2[7]  * D3M$Treated_O[j]
  D3M$P3[j+7]<-D3M$P3[j+7] + D3M$Delayed_P3[j+6] + LEAD$P2_P3[7] * D3M$Treated_P2[j] + LEAD$P3_P3[7] * D3M$Treated_P3[j]
  
  # LEAD TIME 8:
  D3M$O[j+8] <-D3M$O[j+8]  + D3M$Delayed_O[j+7]  + LEAD$P1_O[8]  * D3M$Treated_P1[j] + LEAD$P2_O[8]  * D3M$Treated_P2[j]
  D3M$P2[j+8]<-D3M$P2[j+8] + D3M$Delayed_P2[j+7] + LEAD$P1_P2[8] * D3M$Treated_P1[j] + LEAD$O_P2[8]  * D3M$Treated_O[j]
  D3M$P3[j+8]<-D3M$P3[j+8] + D3M$Delayed_P3[j+7] + LEAD$P2_P3[8] * D3M$Treated_P2[j] + LEAD$P3_P3[8] * D3M$Treated_P3[j]
  

 }


#===============================================================================
# Creating a third loop to implement the delays of O, P1, P2, P3 and the costs 

#-------------------------------------------------------------------------------
# Calculating the penalties for delayed Operations (O) for week 1-3 + REST
for (k in 1:52) {

  D3M$Dif_Delayed_O[k+1] <- D3M$Treated_O[k+1] - D3M$Delayed_O[k]
  
  if (D3M$Dif_Delayed_O[k+1] > 0){
    D3M$Delay_O_1Week[k] <- D3M$Delayed_O[k]
  }else{
    D3M$Delay_O_1Week[k]
    if(D3M$Delayed_O[k] - D3M$Treated_O[k+1] < D3M$Treated_O[k+2]){
      D3M$Delay_O_1Week[k]<- D3M$Treated_O[k+1]
      D3M$Delay_O_2Week[k]<- D3M$Delayed_O[k] - D3M$Treated_O[k+1]
    }else{
      D3M$Delay_O_1Week[k]
      D3M$Delay_O_2Week[k] 
      if(D3M$Delayed_O[k] - D3M$Treated_O[k+1] - D3M$Treated_O[k+2] < D3M$Treated_O[k+3]){
        D3M$Delay_O_1Week[k]<- D3M$Treated_O[k+1]
        D3M$Delay_O_2Week[k]<- D3M$Treated_O[k+2]
        D3M$Delay_O_3Week[k]<- D3M$Delayed_O[k] - D3M$Treated_O[k+1] - D3M$Treated_O[k+2]
      }else{
        D3M$Delay_O_1Week[k]
        D3M$Delay_O_2Week[k]
        D3M$Delay_O_3Week[k]
        if(D3M$Delayed_O[k] - D3M$Treated_O[k+1] - D3M$Treated_O[k+2] - D3M$Treated_O[k+3] < D3M$Treated_O[k+4]){
          D3M$Delay_O_1Week[k]<- D3M$Treated_O[k+1]
          D3M$Delay_O_2Week[k]<- D3M$Treated_O[k+2]
          D3M$Delay_O_3Week[k]<- D3M$Treated_O[k+3]
          D3M$Delay_O_REST[k]<- D3M$Delayed_O[k] - D3M$Treated_O[k+1] - D3M$Treated_O[k+2] - D3M$Treated_O[k+3]
        }else{
          D3M$Delay_O_1Week[k]<- D3M$Treated_O[k+1]
          D3M$Delay_O_2Week[k]<- D3M$Treated_O[k+2]
          D3M$Delay_O_3Week[k]<- D3M$Treated_O[k+3]
          D3M$Delay_O_REST[k]<- D3M$Delayed_O[k] - D3M$Treated_O[k+1] - D3M$Treated_O[k+2] - D3M$Treated_O[k+3]
          }
      }
    }
    
  }
 
  #-------------------------------------------------------------------------------
  # Calculating the penalties for delayed P1 for week 1-3 + REST
  
  D3M$Dif_Delayed_P1[k+1] <- D3M$Treated_P1[k+1] - D3M$Delayed_P1[k]
  
  if (D3M$Dif_Delayed_P1[k+1] > 0){
    D3M$Delay_P1_1Week[k] <- D3M$Delayed_P1[k]
  }else{
    D3M$Delay_P1_1Week[k]
    if(D3M$Delayed_P1[k] - D3M$Treated_P1[k+1] < D3M$Treated_P1[k+2]){
      D3M$Delay_P1_1Week[k]<- D3M$Treated_P1[k+1]
      D3M$Delay_P1_2Week[k]<- D3M$Delayed_P1[k] - D3M$Treated_P1[k+1]
    }else{
      D3M$Delay_P1_1Week[k]
      D3M$Delay_P1_2Week[k] 
      if(D3M$Delayed_P1[k] - D3M$Treated_P1[k+1] - D3M$Treated_P1[k+2] < D3M$Treated_P1[k+3]){
        D3M$Delay_P1_1Week[k]<- D3M$Treated_P1[k+1]
        D3M$Delay_P1_2Week[k]<- D3M$Treated_P1[k+2]
        D3M$Delay_P1_3Week[k]<- D3M$Delayed_P1[k] - D3M$Treated_P1[k+1] - D3M$Treated_P1[k+2]
      }else{
        D3M$Delay_P1_1Week[k]
        D3M$Delay_P1_2Week[k]
        D3M$Delay_P1_3Week[k]
        if(D3M$Delayed_P1[k] - D3M$Treated_P1[k+1] - D3M$Treated_P1[k+2] - D3M$Treated_P1[k+3] < D3M$Treated_P1[k+4]){
          D3M$Delay_P1_1Week[k]<- D3M$Treated_P1[k+1]
          D3M$Delay_P1_2Week[k]<- D3M$Treated_P1[k+2]
          D3M$Delay_P1_3Week[k]<- D3M$Treated_P1[k+3]
          D3M$Delay_P1_REST[k]<- D3M$Delayed_P1[k] - D3M$Treated_P1[k+1] - D3M$Treated_P1[k+2] - D3M$Treated_P1[k+3]
        }else{
          D3M$Delay_P1_1Week[k]<- D3M$Treated_P1[k+1]
          D3M$Delay_P1_2Week[k]<- D3M$Treated_P1[k+2]
          D3M$Delay_P1_3Week[k]<- D3M$Treated_P1[k+3]
          D3M$Delay_P1_REST[k]<- D3M$Delayed_P1[k] - D3M$Treated_P1[k+1] - D3M$Treated_P1[k+2] - D3M$Treated_P1[k+3]
        }
      }
    }
    
  }
  
  #-------------------------------------------------------------------------------
  # Calculating the penalties for delayed P2 for week 1-3 + REST
  
  D3M$Dif_Delayed_P2[k+1] <- D3M$Treated_P2[k+1] - D3M$Delayed_P2[k]
  
  if (D3M$Dif_Delayed_P2[k+1] > 0){
    D3M$Delay_P2_1Week[k] <- D3M$Delayed_P2[k]
  }else{
    D3M$Delay_P2_1Week[k]
    if(D3M$Delayed_P2[k] - D3M$Treated_P2[k+1] < D3M$Treated_P2[k+2]){
      D3M$Delay_P2_1Week[k]<- D3M$Treated_P2[k+1]
      D3M$Delay_P2_2Week[k]<- D3M$Delayed_P2[k] - D3M$Treated_P2[k+1]
    }else{
      D3M$Delay_P2_1Week[k]
      D3M$Delay_P2_2Week[k] 
      if(D3M$Delayed_P2[k] - D3M$Treated_P2[k+1] - D3M$Treated_P2[k+2] < D3M$Treated_P2[k+3]){
        D3M$Delay_P2_1Week[k]<- D3M$Treated_P2[k+1]
        D3M$Delay_P2_2Week[k]<- D3M$Treated_P2[k+2]
        D3M$Delay_P2_3Week[k]<- D3M$Delayed_P2[k] - D3M$Treated_P2[k+1] - D3M$Treated_P2[k+2]
      }else{
        D3M$Delay_P2_1Week[k]
        D3M$Delay_P2_2Week[k]
        D3M$Delay_P2_3Week[k]
        if(D3M$Delayed_P2[k] - D3M$Treated_P2[k+1] - D3M$Treated_P2[k+2] - D3M$Treated_P2[k+3] < D3M$Treated_P2[k+4]){
          D3M$Delay_P2_1Week[k]<- D3M$Treated_P2[k+1]
          D3M$Delay_P2_2Week[k]<- D3M$Treated_P2[k+2]
          D3M$Delay_P2_3Week[k]<- D3M$Treated_P2[k+3]
          D3M$Delay_P2_REST[k]<- D3M$Delayed_P2[k] - D3M$Treated_P2[k+1] - D3M$Treated_P2[k+2] - D3M$Treated_P2[k+3]
        }else{
          D3M$Delay_P2_1Week[k]<- D3M$Treated_P2[k+1]
          D3M$Delay_P2_2Week[k]<- D3M$Treated_P2[k+2]
          D3M$Delay_P2_3Week[k]<- D3M$Treated_P2[k+3]
          D3M$Delay_P2_REST[k]<- D3M$Delayed_P2[k] - D3M$Treated_P2[k+1] - D3M$Treated_P2[k+2] - D3M$Treated_P2[k+3]
        }
      }
    }
    
  }
  
  #-------------------------------------------------------------------------------
  # Calculating the penalties for delayed P3 for week 1-3 + REST
  
  D3M$Dif_Delayed_P3[k+1] <- D3M$Treated_P3[k+1] - D3M$Delayed_P3[k]
  
  if (D3M$Dif_Delayed_P3[k+1] > 0){
    D3M$Delay_P3_1Week[k] <- D3M$Delayed_P3[k]
  }else{
    D3M$Delay_P3_1Week[k]
    if(D3M$Delayed_P3[k] - D3M$Treated_P3[k+1] < D3M$Treated_P3[k+2]){
      D3M$Delay_P3_1Week[k]<- D3M$Treated_P3[k+1]
      D3M$Delay_P3_2Week[k]<- D3M$Delayed_P3[k] - D3M$Treated_P3[k+1]
    }else{
      D3M$Delay_P3_1Week[k]
      D3M$Delay_P3_2Week[k] 
      if(D3M$Delayed_P3[k] - D3M$Treated_P3[k+1] - D3M$Treated_P3[k+2] < D3M$Treated_P3[k+3]){
        D3M$Delay_P3_1Week[k]<- D3M$Treated_P3[k+1]
        D3M$Delay_P3_2Week[k]<- D3M$Treated_P3[k+2]
        D3M$Delay_P3_3Week[k]<- D3M$Delayed_P3[k] - D3M$Treated_P3[k+1] - D3M$Treated_P3[k+2]
      }else{
        D3M$Delay_P3_1Week[k]
        D3M$Delay_P3_2Week[k]
        D3M$Delay_P3_3Week[k]
        if(D3M$Delayed_P3[k] - D3M$Treated_P3[k+1] - D3M$Treated_P3[k+2] - D3M$Treated_P3[k+3] < D3M$Treated_P3[k+4]){
          D3M$Delay_P3_1Week[k]<- D3M$Treated_P3[k+1]
          D3M$Delay_P3_2Week[k]<- D3M$Treated_P3[k+2]
          D3M$Delay_P3_3Week[k]<- D3M$Treated_P3[k+3]
          D3M$Delay_P3_REST[k]<- D3M$Delayed_P3[k] - D3M$Treated_P3[k+1] - D3M$Treated_P3[k+2] - D3M$Treated_P3[k+3]
        }else{
          D3M$Delay_P3_1Week[k]<- D3M$Treated_P3[k+1]
          D3M$Delay_P3_2Week[k]<- D3M$Treated_P3[k+2]
          D3M$Delay_P3_3Week[k]<- D3M$Treated_P3[k+3]
          D3M$Delay_P3_REST[k]<- D3M$Delayed_P3[k] - D3M$Treated_P3[k+1] - D3M$Treated_P3[k+2] - D3M$Treated_P3[k+3]
        }
      }
    }
    
  }
  #-----------------------------------------------------------------------------
  # Calculating the sum of penalties for O, P1, P2 and P3:
  
  # Penalty O:
  D3M$Penalty_O[k] <- D3M$Delay_O_1Week[k] * penalty_COST_O * 1^2 + D3M$Delay_O_2Week[k] * penalty_COST_O  * 2^2 + D3M$Delay_O_3Week[k]* penalty_COST_O  * 3^2 +D3M$Delay_O_REST[k]* penalty_COST_O  * 4^2
  
  # Penalty P1:
  D3M$Penalty_P1[k] <- D3M$Delay_P1_1Week[k] * penalty_COST_P1 * 1^2 + D3M$Delay_P1_2Week[k] * penalty_COST_P1  * 2^2 + D3M$Delay_P1_3Week[k]* penalty_COST_P1  * 3^2 +D3M$Delay_P1_REST[k]* penalty_COST_P1  * 4^2
  
  # Penalty P2:
  D3M$Penalty_P2[k] <- D3M$Delay_P2_1Week[k] * penalty_COST_P2 * 1^2 + D3M$Delay_P2_2Week[k] * penalty_COST_P2  * 2^2 + D3M$Delay_P2_3Week[k]* penalty_COST_P2  * 3^2 +D3M$Delay_P2_REST[k]* penalty_COST_P2  * 4^2
  
  # Penalty P3:
  D3M$Penalty_P3[k] <- D3M$Delay_P3_1Week[k] * penalty_COST_P3 * 1^2 + D3M$Delay_P3_2Week[k] * penalty_COST_P3  * 2^2 + D3M$Delay_P3_3Week[k]* penalty_COST_P3  * 3^2 +D3M$Delay_P3_REST[k]* penalty_COST_P3  * 4^2
  
  #-----------------------------------------------------------------------------
  # Calculating the total cost for each week
  
  D3M$Cost[k] <- D3M$os_2[k] * 100 + D3M$ps_2[k] * 50 + (D3M$os_2[k] - D3M$os_0[k])*(-50) + (D3M$ps_2[k] - D3M$ps_0[k])*(-30) + 
    D3M$Penalty_O[k] + D3M$Penalty_P1[k] + D3M$Penalty_P2[k] + D3M$Penalty_P3[k]
  }# Closing loop


#===============================================================================
# Calculating the overall results:

# Calculating the Total Cost for the week 1-52: 

TotalCost <- sum(D3M$Cost, na.rm = TRUE)
TotalCost # Result: 526,210.10

# Calculating the penalty for O for the weeks 1-52: 

TotalPenalty_O <- sum(D3M$Penalty_O, na.rm = TRUE)
TotalPenalty_O # Result: 0

# Calculating the penalty for P1 for the weeks 1-52: 

TotalPenalty_P1 <- sum(D3M$Penalty_P1, na.rm = TRUE)
TotalPenalty_P1 # Result: 0

# Calculating the penalty for P2 for the weeks 1-52: 

TotalPenalty_P2 <- sum(D3M$Penalty_P2, na.rm = TRUE)
TotalPenalty_P2 # Result: 10,707.54

# Calculating the penalty for P3 for the weeks 1-52: 

TotalPenalty_P3 <- sum(D3M$Penalty_P3, na.rm = TRUE)
TotalPenalty_P3 # Result:412,502.5

# Calculating the conducted operations for the weeks 1-52: 

Total_O <- sum(D3M$Treated_O, na.rm = TRUE)
Total_O # Result: 2472

# Calculating the conducted P1 for the weeks 1-52: 

Total_P1 <- sum(D3M$Treated_P1, na.rm = TRUE)
Total_P1 # Result: 2502.65

# Calculating the conducted P2 for the weeks 1-52: 

Total_P2 <- sum(D3M$Treated_P2, na.rm = TRUE)
Total_P2 # Result: 3325.943

# Calculating the conducted P3 for the weeks 1-52: 

Total_P3 <- sum(D3M$Treated_P3, na.rm = TRUE)
Total_P3 # 763.4075

#-------------------------------------------------------------------------------
# Displaying the remaining delayed patients at the end of the year that need to
# be moved to the next year:

# Delayed_O, Delayed_P1, Delayed_P2, Delayed_3

print(D3M$Delayed_O[52]) # Result: 0
print(D3M$Delayed_P1[52]) # Result: 0
print(D3M$Delayed_P2[52]) # Result: 126
print(D3M$Delayed_P3[52]) # Result: 1246,018

#===============================================================================
# CREATE DIAGRAMS

# Create a subset for week 1 to 52

YEAR <- D3M[1:52,]

# Create Diagram for Number of Patients (O, P1, P2, P3)
# using geom_smooth because geom_line looks messy and is hard to read

NumberPatients <- ggplot(YEAR,aes(Time (week), Patient)) + ggtitle("Plot Number of patients O, P1, P2 and P3") +
  geom_smooth(mapping = aes(x = week, y = O,color = "Operation"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = P1,color = "P1"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = P2,color = "P2"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = P3,color = "P3"),se=FALSE)
NumberPatients

# Create Diagram for the development of capacity and scheduled sessions

NumberCapacity <- ggplot(YEAR,aes(Time (week), Sessions)) + ggtitle("Plot Number of capacity, o.s.(0) and p.s.(0)") +
  geom_smooth(mapping = aes(x = week, y = capacity,color = "Capacity"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = os_0,color = "o.s.(0)"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = ps_0,color = "p.s.(0)"),se=FALSE) 
NumberCapacity

# Create Diagram for the cost

cost <- ggplot(YEAR,aes(Time (week), MonetaryUnits)) + ggtitle("Plot cost") +
  geom_smooth(mapping = aes(x = week, y = Cost,color = "Cost"),se=FALSE) 
cost

# Create Diagram for the penalties costs of O, P1, P2 and P3

penalties <- ggplot(YEAR,aes(Time (week), MonetaryUnits)) + ggtitle("Plot of penalties costs for O, P1, P2 and P3") +
  geom_smooth(mapping = aes(x = week, y = Penalty_O,color = "penalty_COST_O"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = Penalty_P1,color = "penalty_COST_P1"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = Penalty_P2,color = "penalty_COST_P2"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = Penalty_P3,color = "penalty_COST_P3"),se=FALSE)
penalties

# Create Diagram for the delayed O, P1, P2 and P3

DelayedPatients <- ggplot(YEAR,aes(Time (week), Patients)) + ggtitle("Plot of delayed patients for O, P1, P2 and P3") +
  geom_smooth(mapping = aes(x = week, y = Delayed_O,color = "Delayed_O"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = Delayed_P1,color = "Delayed_P1"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = Delayed_P2,color = "Delayed_P2"),se=FALSE) +
  geom_smooth(mapping = aes(x = week, y = Delayed_P3,color = "Delayed_P3"),se=FALSE)
DelayedPatients


#-------------------------------------------------------------------------------
# WRITE OUT EXCEL FILE

write_xlsx(D3M, "D3M_Solution_1st_Approach_B.xlsx")
