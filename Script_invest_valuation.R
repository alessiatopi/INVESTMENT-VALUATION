rm(list=ls())
library(rstudioapi)
current_path<-getActiveDocumentContext()$path
setwd(dirname(current_path))
print(getwd())
library(readxl)
library(tidyverse)
library(lubridate)
library(ggplot2)

# Caricamento dati
data <- read_excel("Russell_3000_Fundamentals_Enlarged_With_README.xlsx", sheet = "Features")
targets <- read_excel("Russell_3000_Fundamentals_Enlarged_With_README.xlsx", sheet = "Targets")
data <- data.frame(data,targets)
data <- data[,-13]
######## DATA QUALITY ###########
# somma per colonna tutti gli na di data:
colSums(is.na(data))

# Vediamo se ci sono righe duplicate
duplicated(data)
# andiamo a vedere quale riga è duplicata
which(duplicated(data))
# Non ci sono righe duplicate

duplicated(data$Isin)
which(duplicated(data$Record.Name))

data_cleaned <- data %>%
  group_by(Record.Name) %>%
  summarise(
    NET_SALES = mean(NET_SALES, na.rm = TRUE),
    across(-NET_SALES, first)
  ) %>%
  ungroup()

summary(data_cleaned)

# BOXPLOT --- vediamo se ci sono outliers
dev.new()
par(mfrow=c(1,3))
boxplot(data_cleaned$FREE_CASH_FLOW, col="lightgreen", main="FREE_CASH_FLOW")
boxplot(data_cleaned$EBITDA, col="lightgreen", main="EBITDA")
boxplot(data_cleaned$ENTERPRISE_VALUE, col="lightgreen", main="ENTERPRISE_VALUE") 
dev.new()
par(mfrow=c(1,3))
boxplot(data_cleaned$NET_SALES, col="lightgreen", main="NET_SALES")
boxplot(data_cleaned$RETURN_ON_.EQUITY, col="lightgreen", main="RETURN_ON_.EQUITY")
boxplot(data_cleaned$RETURN_ON_INVESTED_CAPITAl, col="lightgreen", main="RETURN_ON_INVESTED_CAPITAl")
dev.new()
par(mfrow=c(1,2))
boxplot(data_cleaned$RETURN_ON_ASSET, col="lightgreen", main="RETURN_ON_ASSET")
boxplot(data_cleaned$EPS_12M_FORWARD, col="lightgreen", main="EPS_12M_FORWARD")

###### RIMOZIONE OUTLIER

# Calcola l'IQR per la variabile FREE CASH FLOW
Q1 <- quantile(data_cleaned$FREE_CASH_FLOW, 0.25)
Q3 <- quantile(data_cleaned$FREE_CASH_FLOW, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 25 * IQR_value
upper_bound <- Q3 + 25 * IQR_value

# Filtra il dataset rimuovendo gli outlier
new_data <- subset(data_cleaned, FREE_CASH_FLOW > lower_bound & FREE_CASH_FLOW < upper_bound)
View(new_data)
boxplot(new_data$FREE_CASH_FLOW, col="lightgreen", main="FREE_CASH_FLOW")


# Calcola l'IQR per la variabile EBITDA
Q1 <- quantile(new_data$EBITDA, 0.25)
Q3 <- quantile(new_data$EBITDA, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 25 * IQR_value
upper_bound <- Q3 + 25 * IQR_value

# Filtra il dataset rimuovendo gli outlier
new_data2 <- subset(new_data, EBITDA > lower_bound & EBITDA < upper_bound)
View(new_data2)
boxplot(new_data2$EBITDA, col="lightgreen", main="EBITDA")


# Calcola l'IQR per la variabile ENTERPRISE_VALUE
Q1 <- quantile(new_data2$ENTERPRISE_VALUE, 0.25)
Q3 <- quantile(new_data2$ENTERPRISE_VALUE, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 25 * IQR_value
upper_bound <- Q3 + 25 * IQR_value

# Filtra il dataset rimuovendo gli outlier
new_data3 <- subset(new_data2, ENTERPRISE_VALUE > lower_bound & ENTERPRISE_VALUE < upper_bound)
View(new_data3)
boxplot(new_data3$ENTERPRISE_VALUE, col="lightgreen", main="ENTERPRISE_VALUE")


# Calcola l'IQR per la variabile NET_SALES
Q1 <- quantile(new_data3$NET_SALES, 0.25)
Q3 <- quantile(new_data3$NET_SALES, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 25 * IQR_value
upper_bound <- Q3 + 25 * IQR_value

# Filtra il dataset rimuovendo gli outlier
new_data4 <- subset(new_data3, NET_SALES > lower_bound & NET_SALES < upper_bound)
View(new_data4)
boxplot(new_data4$NET_SALES, col="lightgreen", main="NET_SALES")


# Calcola l'IQR per la variabile RETURN_ON_EQUITY
Q1 <- quantile(new_data4$RETURN_ON_.EQUITY, 0.25)
Q3 <- quantile(new_data4$RETURN_ON_.EQUITY, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 25 * IQR_value
upper_bound <- Q3 + 25 * IQR_value

# Filtra il dataset rimuovendo gli outlier
new_data5 <- subset(new_data4, RETURN_ON_.EQUITY > lower_bound & RETURN_ON_.EQUITY < upper_bound)
View(new_data5)
boxplot(new_data5$RETURN_ON_.EQUITY, col="lightgreen", main="RETURN_ON_.EQUITY")


# Calcola l'IQR per la variabile RETURN_ON_INVESTED_CAPITAl
Q1 <- quantile(new_data5$RETURN_ON_INVESTED_CAPITAl, 0.25)
Q3 <- quantile(new_data5$RETURN_ON_INVESTED_CAPITAl, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 25 * IQR_value
upper_bound <- Q3 + 25 * IQR_value

# Filtra il dataset rimuovendo gli outlier
new_data6 <- subset(new_data5, RETURN_ON_INVESTED_CAPITAl > lower_bound & RETURN_ON_INVESTED_CAPITAl < upper_bound)
View(new_data6)
boxplot(new_data6$RETURN_ON_INVESTED_CAPITAl, col="lightgreen", main="RETURN_ON_INVESTED_CAPITAl")


# Calcola l'IQR per la variabile RETURN_ON_ASSET
Q1 <- quantile(new_data6$RETURN_ON_ASSET, 0.25)
Q3 <- quantile(new_data6$RETURN_ON_ASSET, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 25 * IQR_value
upper_bound <- Q3 + 25 * IQR_value

# Filtra il dataset rimuovendo gli outlier
new_data7 <- subset(new_data6, RETURN_ON_ASSET > lower_bound & RETURN_ON_ASSET < upper_bound)
View(new_data7)
boxplot(new_data7$RETURN_ON_ASSET, col="lightgreen", main="RETURN_ON_ASSET")


# Calcola l'IQR per la variabile EPS_12M_FORWARD
Q1 <- quantile(new_data7$EPS_12M_FORWARD, 0.25)
Q3 <- quantile(new_data7$EPS_12M_FORWARD, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 25 * IQR_value
upper_bound <- Q3 + 25 * IQR_value

# Filtra il dataset rimuovendo gli outlier
new_data8 <- subset(new_data7, EPS_12M_FORWARD > lower_bound & EPS_12M_FORWARD < upper_bound)
View(new_data8)
boxplot(new_data8$EPS_12M_FORWARD, col="lightgreen", main="EPS_12M_FORWARD")

data_ok<-new_data8
dev.new()
data_esempio <- data_ok[which(data_ok$INDUSTRY=="Consumer Discretionary"),]
ggplot(data_esempio, aes(x = SUB_INDUSTRY, y = EPS_12M_FORWARD)) + geom_violin() + coord_flip()
dev.new()
ggplot(data_ok %>% filter(SUB_INDUSTRY == "Consumer Discretionary"), 
       aes(x = SUB_INDUSTRY, y = EPS_12M_FORWARD)) +
  geom_boxplot(fill = "lightblue", color = "black", outlier.color = "red") +
  coord_flip() +  # Ruota il grafico per leggibilità
  labs(title = "Distribuzione di EPS per sottosettore (Consumer Discretionary)",
       x = "Sottosettore",
       y = "EPS 12M Forward") +
  theme_minimal()

ggplot(data_ok %>% filter(SUB_INDUSTRY == "Consumer Discretionary"), 
       aes(x = SUB_INDUSTRY, y = EPS_12M_FORWARD)) +
  geom_violin(fill = "lightgreen", alpha = 0.5) +
  geom_boxplot(width = 0.1, fill = "white") +  # Aggiunge la mediana e quartili
  coord_flip() +
  labs(title = "Distribuzione di EPS per sottosettore (Consumer Discretionary)",
       x = "Sottosettore",
       y = "EPS 12M Forward") +
  theme_minimal()

dev.new()
ggplot(data_ok %>% filter(SUB_INDUSTRY == "Consumer Discretionary") %>%
         group_by(SUB_INDUSTRY) %>%
         summarise(mean_EPS = mean(EPS_12M_FORWARD, na.rm = TRUE)),
       aes(x = reorder(SUB_INDUSTRY, mean_EPS), y = mean_EPS)) +
  geom_bar(stat = "identity", fill = "blue", alpha = 0.7) +
  coord_flip() +
  labs(title = "Media di EPS per sottosettore (Consumer Discretionary)",
       x = "Sottosettore",
       y = "Media EPS 12M Forward") +
  theme_minimal()


#### CORRELAZIONE TRA LE VARIABILI
data_aux <- new_data8[,-c(1,3:6)]
dev.new()
matrcorr=cor(data_aux)
#library(ggcorrplot)
#ggcorrplot(matrcorr)
library(corrplot)
corrplot(matrcorr, method = "color", 
         addCoef.col = "black", # colore dei coefficienti
         number.cex = 0.7)

## FORTE correlazione positiva tra: FREE CASH FLOW, EBITDA, ENTERPRISE VALUE, NET_SALES
## FORTE correlazione positiva tra: RETURN ON ASSET, RETURN ON INVESTED CAPITAL, RETURN_ON_EQUITY


## Trasformazione logaritmica dei dati

library(nlme)
data_ok[,c(2,7:13)] <- sign(data_ok[,c(2,7:13)]) * log(abs(data_ok[,c(2,7:13)]) + 1)


###### RLS ####### --> lo usiamo per togliere gli outliers per i residui dei modello (da 2411 osservazioni a 2329 osservazioni)
library(MASS)

model_robust <- rlm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY + RETURN_ON_INVESTED_CAPITAl + RETURN_ON_ASSET, data=data_ok)

summary(model_robust)

### Valuto collinearità
library(car)
vif(model_robust)

model_robust2 <- rlm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY + RETURN_ON_ASSET, data=data_ok)

summary(model_robust2)
vif(model_robust2)

model_robust3 <- rlm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY, data=data_ok)

summary(model_robust3)
vif(model_robust3) # collinearità ok

r2_rlm <- 1 - sum(resid(model_robust3)^2) / sum((data_ok$EPS_12M_FORWARD - mean(data_ok$EPS_12M_FORWARD))^2)
r2_rlm # R quadro = 54.1%

# Grafico valori reali vs valori predetti
ggplot(data_ok, aes(x = EPS_12M_FORWARD, y = model_robust3$fitted.values)) +
  
  geom_point(color = "blue", size = 3) +  # Punti per y vs y_pred
  
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +  # Linea ideale y = y_pred
  
  labs(title = "Confronto tra Valori Reali e Predetti",
       
       x = "Valori Reali (y)",
       
       y = "Valori Predetti (ŷ)") +
  
  theme_minimal()

durbinWatsonTest(model_robust3) # non c'è autocorrelazione degli errori

residui_rlm <- residuals(model_robust)

# Calcola l'IQR per i residui
Q1 <- quantile(residui_rlm, 0.25)
Q3 <- quantile(residui_rlm, 0.75)
IQR_value <- Q3 - Q1

# Identifica gli outlier
lower_bound <- Q1 - 1.3 * IQR_value
upper_bound <- Q3 + 1.3 * IQR_value

data_noresid <- data.frame(data_ok,residui_rlm)
# Filtra il dataset rimuovendo gli outlier
data_noresid <- subset(data_noresid, residui_rlm > lower_bound & residui_rlm < upper_bound)



################# REGRESSIONE OLS  ###############

result=lm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY + RETURN_ON_INVESTED_CAPITAl + RETURN_ON_ASSET, data=data_noresid) 

### statistica descrittiva
print(summary(result)) 

### Valuto collinearità
library(car)
vif(result)

#### Togliamo ROI
result2=lm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY + RETURN_ON_ASSET, data=data_noresid) 

### statistica descrittiva
print(summary(result2))
vif(result2)

#### Togliamo ROA
result3=lm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY, data=data_noresid) 

### statistica descrittiva
print(summary(result3))

vif(result3)
### Ora non c'è multicollinearità

##### NON FUNZIONA PERCHE' NON RISPETTA IPOTESI DI OMOSCHEDASTICITA'



############## WLS - MODELLO FINALE #############

library(nlme)
w <- 1 / lm(abs(resid(result3)) ~ fitted(result3))$fitted.values^2
model_wls <- lm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY, data=data_noresid, weights=w)
summary(model_wls)
# R quadro 61.1 %

# omoschedasticità --> ok
ncvTest(model_wls)

# normalità --> ok
library(olsrr)
ols_plot_resid_qq(model_wls) ## Confronta quantili teorici e osservati (voglio i puntini sulla linea)
ols_test_normality(model_wls) # = nei test, in questi test voglio che i pvalue siano alti perchè Ho è che gli errori sono normali
ols_plot_resid_hist(model_wls) ### grafico verifica normalità residui

# Test per l'autocorrelazione degli errori --> ok
durbinWatsonTest(model_wls)

# Test di Chow --> ok
library(strucchange)
chow_test <- sctest(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY, type = "Chow", data = data_noresid)
chow_test

View(cbind(data_noresid$EPS_12M_FORWARD,model_wls$fitted.values))



# Grafico valori reali e predetti
ggplot(data_noresid, aes(x = data_noresid$EPS_12M_FORWARD, y = model_wls$fitted.values)) +
  
  geom_point(color = "blue", size = 3) +  # Punti per y vs y_pred
  
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +  # Linea ideale y = y_pred
  
  labs(title = "Confronto tra Valori Reali e Predetti",
       
       x = "Valori Reali (y)",
       
       y = "Valori Predetti (ŷ)") +
  
  theme_minimal()


## Confronto WLS e RLS 
AIC(model_robust3,model_wls) # più basso WLS
BIC(model_robust3,model_wls) # più basso WLS

# Vediamo dove i residui sono più distribuiti --> WLS
dev.new()
par(mfrow=c(1,3))
# Grafici dei residui vs fitted values
plot(fitted(result3), resid(result3))
plot(fitted(model_robust3), resid(model_robust3))
plot(fitted(model_wls), resid(model_wls))


### K-FOLD CROSS VALIDATION
library(caret)
set.seed(123)  # Per riproducibilità
data_noresid <- data_noresid[,-14]

# Creiamo un indice casuale per il test set (30% dei dati)
trainIndex <- createDataPartition(data_noresid$EPS_12M_FORWARD, p = 0.7, list = FALSE)

# Train e Test Set
train_data <- data_noresid[trainIndex, ]
test_data  <- data_noresid[-trainIndex, ]


# Definire i parametri della Cross-Validation (8-fold)
train_control <- trainControl(method = "cv", number = 8)  

# Modello con cross-validation

## Faccio OLS per calcolare i pesi
result=lm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY + RETURN_ON_INVESTED_CAPITAl + RETURN_ON_ASSET, data=train_data) 

### statistica descrittiva
print(summary(result)) 

### Valuto collinearità
library(car)
vif(result)

#### Togliamo ROI
result2=lm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY + RETURN_ON_ASSET, data=train_data) 

### statistica descrittiva
print(summary(result2))
vif(result2)

#### Togliamo ROA
result3=lm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + EBITDA + ENTERPRISE_VALUE + RETURN_ON_.EQUITY, data=train_data) 

### statistica descrittiva
print(summary(result3))

vif(result3)
### Ora non c'è multicollinearità

#### Togliamo EBITDA
result4=lm(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + ENTERPRISE_VALUE + RETURN_ON_.EQUITY, data=train_data) 

### statistica descrittiva
print(summary(result4))

vif(result4)


### Addestramento modello sul training

w <- 1 / lm(abs(resid(result4)) ~ fitted(result4))$fitted.values^2

library(caret)
model_cv <- train(EPS_12M_FORWARD ~ NET_SALES + FREE_CASH_FLOW + ENTERPRISE_VALUE + RETURN_ON_.EQUITY, weights=w, 
                  data = train_data, 
                  method = "lm", 
                  trControl = train_control)

# Visualizzare i risultati della cross-validation
print(model_cv)
print(model_cv$finalModel)
varImp(model_cv)
summary(model_cv$finalModel)
vif(model_cv$finalModel)

###### VERIFICA OVERFITTING e PREVISIONI SU TRAIN E TEST ######

# Previsioni sul training set
predictions_train <- predict(model_cv, newdata = train_data)

# Valutazione delle performance
mse_train <- mean((train_data$EPS_12M_FORWARD - predictions_train)^2)
rmse_train <- sqrt(mse_train)
r2_train <- cor(train_data$EPS_12M_FORWARD, predictions_train)^2

# Stampare i risultati
cat("MSE:", mse_train, "\nRMSE:", rmse_train, "\nR-squared:", r2_train, "\n")


# Previsioni sul test set
predictions <- predict(model_cv, newdata = test_data)

# Valutazione delle performance
mse <- mean((test_data$EPS_12M_FORWARD - predictions)^2)
rmse <- sqrt(mse)
r2 <- cor(test_data$EPS_12M_FORWARD, predictions)^2

# Stampare i risultati
cat("MSE:", mse, "\nRMSE:", rmse, "\nR-squared:", r2, "\n")

# Non c'è overfitting

library(lmtest)
resettest(model_cv$finalModel, power = 2, type = "fitted")

step(model_cv$finalModel, direction = "both")

########## INTERVALLI DI CONFIDENZA SUL TEST ##########


data_media <- data_noresid[which(data_noresid$SUB_INDUSTRY=="Media"),-14]

# Previsione sul subset del settore Media
predictions_media <- predict(model_cv, newdata = data_media)
predictions_media_interval <- predict(model_cv$finalModel, newdata = data_media, interval = "confidence", level = 0.90)

# Converti predictions_media in un dataframe
predictions_media_interval <- as.data.frame(predictions_media_interval)


# Applica la trasformazione inversa per riportare EPS nella scala originale
predictions_media_interval$fit <- sign(predictions_media_interval$fit) * (exp(abs(predictions_media_interval$fit)) - 1)
predictions_media_interval$lwr <- sign(predictions_media_interval$lwr) * (exp(abs(predictions_media_interval$lwr)) - 1)
predictions_media_interval$upr <- sign(predictions_media_interval$upr) * (exp(abs(predictions_media_interval$upr)) - 1)

mean(predictions_media_interval$fit)
mean(predictions_media_interval$lwr)
mean(predictions_media_interval$upr)

mean(data_media$EPS_upr - data_media$EPS_lwr)

summary(predictions_media_interval$fit)
summary(predictions_media_interval$lwr)
summary(predictions_media_interval$upr)

# Aggiungi le previsioni trasformate al dataset originale
data_media$EPS_pred <- predictions_media_interval$fit
data_media$EPS_lwr <- predictions_media_interval$lwr
data_media$EPS_upr <- predictions_media_interval$upr


# Creare il grafico con intervalli di confidenza
ggplot(data_media, aes(x = 1:nrow(data_media), y = EPS_pred)) +
  geom_line(color = "blue", size = 1) +  # Linea delle previsioni
  geom_ribbon(aes(ymin = EPS_lwr, ymax = EPS_upr), fill = "gray70", alpha = 0.5) +  # Area dell'intervallo di confidenza
  geom_point(aes(y = EPS_12M_FORWARD), color = "red", size = 1.5) +  # Punti dei valori reali
  labs(title = "Intervallo di Confidenza 90% per EPS (Settore Media)",
       x = "Osservazioni",
       y = "Valore EPS") +
  theme_minimal()


# Previsione sul subset del settore Telecommunications
data_telecomm <- data_noresid[which(data_noresid$SUB_INDUSTRY=="Telecommunications Service Providers"),-14]
predictions_telecomm <- predict(model_cv, newdata = data_telecomm)
mean(predictions_telecomm)

# Previsione sul subset del settore Retailers
data_retail <- data_noresid[which(data_noresid$SUB_INDUSTRY=="Retailers"),-14]
predictions_retail <- predict(model_cv, newdata = data_retail)
mean(predictions_retail)

