library(fpp3)
library(imputeTS)
library(dplyr)
library(tidyverse)
library(anytime)
library(ggplot2)
library("recipes")
library("timetk")
library(plotly)

#air = as_tsibble(AirPassengers)
#air = AirPassengers

cheburek <- read.csv("C:/Users/RomanBevz/Documents/detectors_Python/pythonProject/cheburek.csv")
cheburek = cheburek %>% 
  mutate(date_time = anytime(Дата)) %>% 
  #mutate(id = interaction(Количество)) %>% 
  subset(select = c(date_time, Количество)) %>% 
  as_tsibble(index = date_time, key = Количество) %>% 
  arrange(date_time)

cheburek %>% 
  plot_time_series(date_time, Количество, 
                   .plotly_slider = TRUE)

cheburek = cheburek[950:1500, c(1,2)]

ggplot_na_distribution(cheburek)
cheburek_new_interpolation = na_interpolation(cheburek)
cheburek_new_interpolation %>% 
  plot_time_series(date_time, Количество, 
                 .plotly_slider = TRUE)

ggplot_na_imputations(cheburek, cheburek_new_interpolation)

model_arima = arima(cheburek$Количество, 
                    order = c(1, 0, 1),
                    seasonal = list(order = c(1, 0, 0)))$model

cheburek_arima = na_kalman(cheburek, 
                           model = model_arima)

ggplot_na_imputations(cheburek, cheburek_arima)
