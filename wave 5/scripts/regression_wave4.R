library(haven)
refugees <- read_sav("901737 (2).sav", encoding = 'CP1251')
library(margins)
library(sjPlot)
library(tidyverse)
library(expss)
library(MASS)
library(ggeffects) # for plotting marginal effects
library(ipumsr)
write_csv(refugees, 'refugees_wave4.csv')
#### Regression analysis
refugees$J2
ref_no_2024_leave <- refugees |> filter(S4a<26)
table(ref_no_2024_leave$Q16)
count(ref_no_2024_leave, Q16, wt = weight)
sum(ref_no_2024_leave$weight)
ref_no_2024_leave |> group_by(Q16) |> summarise(weighted = sum(weight), 
                                                perc =weighted/900.7367 )

##Recoding the intention tor return in a different order + include the difficult to say in the middle
##Higher number = more likely to return
refugees <- refugees |> mutate(int_return = case_when(J2 == 1 ~5,
                                                      J2 == 2 ~ 4,
                                                      J2 == 5 ~ 3,
                                                      J2 == 3 ~ 2,
                                                      J2 == 4 ~ 1))
## Coding missing data
assign_missing <- function(x){ 
  lbl_na_if(x, ~.lbl %in% c("Відмова відповідати",
                            'Відсутня відповідь',
                            'Важко сказати',
                            'Відмова від відповіді',
                            'немає відповіді'))
}
##Recoding missing data in our dataset
refugees <- map_df(refugees, assign_missing)
##Applying labels
refugees <-  add_labelled_class(refugees)

#### Cleaning
#Maybe recode oblast to macroregion?
west <- c(3,
          7,
          9,
          14,
          18,
          20,
          23,
          25
)

centre <- c(2,
            12,
            17,
            24
)

south <- c(1,
           8
           , 15
           ,16
)

east <- c(4,
          5,
          13,
          21
)

north <- c(6,
           10,
           11,
           19,
           26
)

##Separate variable for regions with military activities in them

warzone <- c(5,
             8,
             10,
             13,
             15,
             19,
             21,
             22,
             26
)
##Recoding
refugees <- refugees %>% mutate(region = case_when(Z1 %in% west ~ 1,
                                                   Z1 %in% centre ~ 2,
                                                   Z1 %in% north ~ 3,
                                                   Z1 %in% south ~4,
                                                   Z1 %in% east~5),
                                war_zone = case_when(Z1 %in% warzone~1,
                                                     !Z1 %in% warzone~0
                                ))


### Recode working categories

##Recode employment to have less categories (before war)
refugees <- refugees %>% mutate(student = ifelse(E1_04 == 1,1,0),
                                working = ifelse(E1_01 == 1 | E1_02 == 1 |E1_10 ==1,1,0),
                                business = ifelse(E1_03 == 1,1,0),
                                unemployed = ifelse(E1_08 == 1,1,0),
                                #work_abroad = ifelse(E1_10 == 1,1,0),
                                out_of_labor = ifelse(E1_05 == 1 | E1_06 == 1 |E1_07 == 1 ,1,0))

## Employment now
refugees<- refugees %>% mutate(student_here_now = ifelse(E2_08 == 1,1,0),
                               attend_integr_courses = ifelse(E2_09 == 1,1,0),
                               student_remote_now = ifelse(E2_07 == 1,1,0),
                               working_now = ifelse(E2_05 == 1 | E2_04 == 1,1,0),
                               business_here_now = ifelse(E2_11 == 1,1,0),
                               business_remote_now = ifelse(E2_10 == 1,1,0),
                               unemployed_now = ifelse(E2_06 == 1,1,0),
                               work_remotely_now = ifelse(E2_01 == 1 |E2_02 == 1,1,0),
                               out_of_labor_now = ifelse(E2_03 == 1 | E2_12 == 1 |E2_13 == 1 |E2_14 == 1  |E2_15 == 1, 1,0))


### Recoding labels of the factors 
refugees <- refugees |> rename(Студент = student_here_now,
                               Працює = working_now,
                               Безробітний = unemployed_now
) |> 
  mutate(region = base::factor(region,
                               levels = c(1,2,3,4,5),
                               labels = c('Захід', 'Центр', 'Північ', 'Південь', 'Схід'
                               )),
         E3.1 = base::factor(unclass(E3), levels = c(1,2,3,4,5,6),
                             labels = c("Найнижчий дохід до війни",
                                        'Дохід до війни:2','Дохід до війни:3','Дохід до війни:4','Дохід до війни:5', "Дохід до війни:6")),
         E3.2 = base::factor(unclass(E4), levels = c(1,2,3,4,5,6),
                             labels = c("Найнижчий дохід зараз",
                                        'Дохід зараз:2','Дохід зараз:3','Дохід зараз:4','Дохід зараз:5', "Дохід зараз:6")))

##Recoding country
refugees <- refugees |> mutate(Country = ifelse(S5 == 7,1, #Germany
                                                ifelse(S5 == 8,2, #Poland
                                                       ifelse(S5 == 17,3, #Czechia
                                                              ifelse(S5 == 3,4, #UK
                                                                     ifelse(S5 == 12,5, #US
                                                                            ifelse(S5 == 18, 6,7 #6 - Canada, 7 -Other
                                                                            )))))) )
##Recoding the variables which are to be included in the plot
refugees <- refugees |> mutate(Country = base::factor(Country, levels = c(1,2,3,4,5,6,7), labels = c('Німеччина', "Польща", "Чехія", "Велика Британія", "США", "Канада", "Інші")) 
) |> rename(Бізнес = business_here_now)

##Recoding gender to binary
refugees <- refugees |> mutate(male = ifelse(S1 == 1,1,0))

## Adding labels to education
attr(refugees$Z4, "label") <- "Освіта"

attr(refugees$S2, "label") <- "Вік"
attr(refugees$male, "label") <- "Чоловік"
attr(refugees$work_remotely_now, "label") <- "Працює дистанційно"


##Recoding education as factor
refugees <- refugees |> mutate(Z4 = base::factor(unclass(Z4), levels = c(1,2,3,4), labels = c('Незакінчена середня', 'Загальна середня', 'Середня спеціальна', 'Вища або незакінчена вища')))

refugees <- refugees %>% mutate(Z2 = base::factor(unclass(Z2), levels = c(1,2,3), labels = c('Село', ' Селище міського типу', 'Місто')))

attr(refugees$Z2, "label") <- 'До війни'


m1 <- polr(factor(int_return)~
          male #Gender
          + S2 #Age
           #+S2^2
           +Country 
           # +factor(Lang)
           +factor(Q6) #marital status
           +A2.1_2 #Left with children
           #+A3_01 #in my city war
           #+A3_02 #near my city war
           #+A3_04 #my city occupied
           #+A3_05 #My house ruined
           #+A3_06 #No job
           #+A3_08 #Far from the frontline, but feel danger
           #+A3_09 #The living conditions is bad
           #+A3_11 #No electricity
           #+A3_14 #Left befre war
           #+A7_01 #Lost income due to war
           #+A7_02 #Lost some income due to war
           #+A7_03 #Family separated
           #+A7_04 #Job loss
           #+A7_05 #house destroyed
           #+A7_12 #psychological issues
           #+A7_18 #no negative consequences
           # +factor(V2) #learning language
           # +factor(V5) #attitude of locals
           # +factor(D1) #type of housing
           # +G1_03 #Recieve social benefits
           # + G1_12 do not recieve any aid
           ####Work before war (out of labor force default category)
           +student
           +working
           +business
           +unemployed
           #+work_abroad
           ####Work now war (out of labor force default category)
           +Студент
           +student_remote_now
           +Працює
           +Бізнес
           +business_remote_now
           +Безробітний
           +work_remotely_now
           +region
           +war_zone
           +Z2 #City type
           +Z4 #education
           +Z5_1 #husband left in Ukraine
           +E3.1#Fin status before war
           +E3.2 #Fin status during war
           #+Q19
           #+Q20
           ,data = refugees
           , weights = weight
)
summary(m1)

av_mar_ef <- margins(m1)
summary(av_mar_ef)

##Include: income before, age, education, country
#regions, 


##Tidying model output with broom
library(broom)
library(readxl)
library(writexl)
coefs <- tidy(m1, exponentiate = TRUE, conf.int = TRUE, conf.level = 0.90)

var_names <- c('male', 'S2', 'work_remotely_now', 'CountryПольща', 'CountryЧехія', 'CountryВелика Британія', 'CountryСША', 'CountryКанада', 'CountryІнші', 'Z2 Селище міського типу', 'Z2Місто', 'E3.1Дохід до війни:2', 'E3.1Дохід до війни:3', 'E3.1Дохід до війни:4', 'E3.1Дохід до війни:5', 'E3.1Дохід до війни:6', 'Z4Загальна середня', 'Z4Середня спеціальна', 'Z4Вища або незакінчена вища') 

coefs <- coefs |> filter(term %in% var_names)
writexl::write_xlsx(coefs, 'coefs.xlsx')

coef_raw <- m1[["coefficients"]][["Z4Середня спеціальна"]]
odds_ratio_manual <- exp(coef_raw)
odds_ratio_manual


# Extract the coefficient and standard errors
est <- coef(m1)
se  <- sqrt(diag(vcov(m1)))

# Compute Wald-based confidence intervals on the log scale
wald_ci <- cbind(
  lower = est - 1.96 * se,
  upper = est + 1.96 * se
)

# Exponentiate both the point estimates and the Wald CIs
results <- exp(cbind(estimate = est, wald_ci))
results

results <- as.data.frame(results)

results <- results %>%
  rownames_to_column(var = "term") 

results <- results |> filter(term %in% var_names)

writexl::write_xlsx(results, 'results.xlsx')


plot_odds <- plot_model(m1,
                        # order.terms = NULL
                        ,rm.terms = c('1|2', '2|3',
                                      '3|4', '4|5')
                        ,terms = c('male','S2', 'work_remotely_now', 'Country [Польща, Чехія, Велика Британія, США, Канада, Інші]', 'Z2 Селище міського типу', 'Z2Місто', 'E3.1 [Дохід до війни:2,Дохід до війни:3,Дохід до війни:4,Дохід до війни:5,Дохід до війни:6]', 'Z4 [ Загальна середня,Середня спеціальна,Вища або незакінчена вища]')
                        ,colors = c("#73B932", "#00509b")
                        ,show.values = TRUE
                        ,title = ''
                        , vline.color = 'black'
                        , value.offset = 0.3
)+theme_bw()
plot_odds


ggsave('plot_odds.png', plot_odds, width = 2000, height = 1700, units = 'px')


#####Notes:
##Include in the plot: Gender, age, country, student now, work now, business here now, regions,current income

#Include in the text: marital status, children, employment status before the war, education, city type, whether they have husband or wife left in Ukraine, pre-war income not statistically significant,


##need to remove the category names for income variables in the datam, so that subsetting works


### For the Eng version
##Recoding the labels names 
refugees_eng <- refugees |> rename(Student = Студент,
                                   Working = Працює,
                                   Unemployed = Безробітний,
                                   student_before = student,
                                   working_before = working,
                                   unemployed_before = unemployed
) |> 
  mutate(region = base::factor(region,
                               levels = c("Захід", "Центр","Північ","Південь","Схід"),
                               labels = c('West', 'Centre', 'North', 'South', 'East'
                               )),
         E3.1 = base::factor(E3.1, levels = c("Найнижчий дохід до війни","Дохід до війни:2","Дохід до війни:3","Дохід до війни:4","Дохід до війни:5","Дохід до війни:6"),
                             labels = c("Lowest pre-war income",
                                        'Pre-war income:2','Pre-war income:3','Pre-war income:4','Pre-war income:5', "Pre-war income:6")),
         E3.2 = base::factor(E3.2, levels = c("Найнижчий дохід зараз",
                                              'Дохід зараз:2','Дохід зараз:3','Дохід зараз:4','Дохід зараз:5', "Дохід зараз:6"),
                             labels = c("Lowest income now",
                                        'Income now:2','Income now:3','Income now:4','Income now:5', "Income now:6")))

##Model

m1_eng <- polr(factor(int_return)~
                 factor(S1) #Gender
               +S2 #Age
               #+S2^2
               +factor(Country) 
               +A2_2 #Left with children
               #+A3_01 #in my city war
               #+A3_02 #near my city war
               #+A3_04 #my city occupied
               #+A3_05 #My house ruined
               # +factor(V2) #learning language
               # +factor(V5) #attitude of locals
               # +factor(D1) #type of housing
               ####Work before war (out of labor force default category)
               +student_before
               +working_before
               +business
               +unemployed_before
               #+work_abroad
               ####Work now war (out of labor force default category)
               +Student
               +student_remote_now
               +Working
               +business_here_now
               +business_remote_now
               +Unemployed
               +work_remotely_now
               +region
               +war_zone
               +factor(Z2) #City type
               +factor(Z4) #education
               +Z5_1 #husband left in Ukraine
               +E3.1#Fin status before war
               +E3.2 #Fin status during war
               ,data = refugees_eng
)

plot_odds <- plot_model(m1_eng 
                        # order.terms = NULL
                        ,rm.terms = c('1|2', '2|3',
                                      '3|4', '4|5'),
                        terms = c('region [Centre,North,South,East]', 'Student', 'Unemployed', 'Working', 'E3.1 [Pre-war income:2,Pre-war income:3,Pre-war income:4,Pre-war income:5,Pre-war income:6]', 'E3.2 [Income now:2,Income now:3,Income now:4,Income now:5,Income now:6]')
                        ,colors = c("#73B932", "#00509b")
                        ,show.values = TRUE
                        ,title = ''
                        , vline.color = 'black'
                        , value.offset = 0.3
)+theme_bw()
plot_odds

ggsave('plot_odds_eng.png', plot_odds, width = 2000, height = 1700, units = 'px')





#############test
m1 <- polr(factor(int_return)~
             factor(S1) #Gender
           +S2 #Age
           #+S2^2
           +factor(Country) 
           +factor(Lang)
           +factor(Q6) #marital status
           +factor(A2.1_2) #Left with children
           #+A3_01 #in my city war
           #+A3_02 #near my city war
           #+A3_04 #my city occupied
           #+A3_05 #My house ruined
           #+A3_06 #No job
           #+A3_08 #Far from the frontline, but feel danger
           #+A3_09 #The living conditions is bad
           #+A3_11 #No electricity
           #+A3_14 #Left befre war
           #+A7_01 #Lost income due to war
           #+A7_02 #Lost some income due to war
           #+A7_03 #Family separated
           #+A7_04 #Job loss
           #+A7_05 #house destroyed
           #+A7_12 #psychological issues
           #+A7_18 #no negative consequences
           # +factor(V2) #learning language
           # +factor(V5) #attitude of locals
           # +factor(D1) #type of housing
           # +G1_03 #Recieve social benefits
           # + G1_12 do not recieve any aid
           
           ####Work before war (out of labor force default category)
           +student
           +working
           +business
           +unemployed
           #+work_abroad
           ####Work now war (out of labor force default category)
           +Студент
           +student_remote_now
           +Працює
           +Бізнес
           +business_remote_now
           +Безробітний
           +work_remotely_now
           +region
           +war_zone
           +factor(Z2) #City type
           +factor(Z4) #education
           +Z5_1 #husband left in Ukraine
           +factor(E3)#Fin status before war
           +factor(E4) #Fin status during war
           #+Q19
           #+Q20
           ,data = refugees
           , weights = Weight
)
summary(m1)

plot_odds <- plot_model(m1 
                        # order.terms = NULL
                        ,rm.terms = c('1|2', '2|3',
                                      '3|4', '4|5')
                        ,terms = c("Lang['Російська']")
                        ,colors = c("#73B932", "#00509b")
                        ,show.values = TRUE
                        ,title = ''
                        , vline.color = 'black'
                        , value.offset = 0.3
                        ,transform = NULL
)+theme_bw()
#plot(ggeffect(m1, terms = c('S2', 'Country',  'E3.1', 'Z4')), rawterms = TRUE) + theme_bw()


table(refugees$region, refugees$Lang)
