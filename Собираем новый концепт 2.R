# setwd("C:/R/Projects/Oper_com_NC")

setwd("~/Desktop/R projects/Newconcept")

# Итак для работы нам нужно:
# 1. файл со списком магазинов НК (Дашборд КВ с портала)
# 2. Выгрузка их табло для сцепки плант-цфо
# 3. Файл актуальный файл "собираем оперком"
#


# Вводим номер недели
weeknumber <- 47

# Вводим константы названий
filename <- paste0("Собираем Оперком ", weeknumber, ".xlsx")
fileresultname <- paste0("Собираем оперком NewConc ", weeknumber, ".xlsx")
filenewconcept <- paste0("NEWC", weeknumber, ".xlsx")

# Загружаем справочник
library(readxl)
plant_cfo <- read_excel("plant-cfo.xlsx")

# Загружаем список нового концепта
df_nc <- read_excel(filenewconcept, skip = 1)
df_nc <- df_nc[c(1:5)] # Оставляем только нужные колонки

# Генерим ДФ со списком магазинов НК
df_nc <- merge(df_nc, plant_cfo, by.x = "Завод", by.y = "Завод Код", all.x = TRUE)

df_nc$'Количество Магазинов' <-c(1) # Для расчёта кол-ва магазинов

# Собираем перечень магазинов нового концепта в один ДФ

df_nc_count_ter <- aggregate(cbind(`Количество Магазинов`)~`Территория`, df_nc, sum)
colnames(df_nc_count_ter) <- c("Признак","Количество Магазинов")

library(janitor) 

df_nc_count_ter <-adorn_totals(df_nc_count_ter, "row", name = "ТС Пятёрочка")

df_nc_count_GC <- aggregate(cbind(`Количество Магазинов`)~`ЦФО Дивизион/гр.кластеров`, df_nc, sum)
colnames(df_nc_count_GC) <- c("Признак","Количество Магазинов")

df_nc_count_MR <- aggregate(cbind(`Количество Магазинов`)~`Макрорегион.y`, df_nc, sum)
colnames(df_nc_count_MR) <- c("Признак","Количество Магазинов")

library(plyr)

df_nc_final <- rbind.fill(df_nc_count_ter, df_nc_count_MR, df_nc_count_GC)

rm(df_nc_count_GC, df_nc_count_MR, df_nc_count_ter) #, plant_cfo) - пока непонятно пригодится или нет

# Приступаем к сбору файла для формирования отчёта

OK_IS <- read_excel(filename, sheet = "ИС")
OK_IS <- OK_IS[1:8]
OK_IS <- merge(OK_IS, df_nc, by.x = "ЦФО ID", by.y = "ЦФО ID", all.x = TRUE) # - надо проверить корректность мерджа

# Заводим функцию для расчёта текучести

library(dplyr)
# Функция для расчёта текучести
turnover <- function(x) {
  x %>%
    mutate(staff_turnover = `Уволенные для текучести (актуальная)` / `ССЧ (усред. по дням) along Table (Down)`)
}

# Функция для переименования столбцов на удобные
name_corr <- function(x) {
  x %>%
    rename(.,
           Признак = c(1),
           Уволенные = c(2),
           ССЧ = c(3),
           Текучесть = c(4)
           
    )
  
}

# Функции для группировки расчёов (аналог сводной)
columns_ter <- function(x) {
  x %>%
  select(`Уволенные для текучести (актуальная)`, 
                  `ССЧ (усред. по дням) along Table (Down)`,
                  `Территория.x`) %>%
                    aggregate(.~`Территория.x`, .,sum) %>%
                    turnover %>%
                    name_corr
}

columns_MR <- function(x) {
  x %>%
    select(`Уволенные для текучести (актуальная)`, 
           `ССЧ (усред. по дням) along Table (Down)`,
           `Макро`) %>%
              aggregate(.~`Макро`, .,sum) %>%
              turnover %>%
              name_corr
}

columns_GC <- function(x) {
  x %>%
    select(`Уволенные для текучести (актуальная)`, 
           `ССЧ (усред. по дням) along Table (Down)`,
           `Группа кластеров`) %>%
            aggregate(.~`Группа кластеров`, .,sum) %>%
            turnover %>%
            name_corr
}



# Собираем текучесть по ИС

OK_IS_ter <- columns_ter(OK_IS)
              
OK_IS_ter <-adorn_totals(OK_IS_ter, "row", name = "ТС Пятёрочка") # Подводим гранд тотал 
OK_IS_ter$Текучесть <- OK_IS_ter$Уволенные / OK_IS_ter$ССЧ # Пересчитываем текучесть для гранд тотал

OK_IS_MR <- columns_MR(OK_IS) # Считаем текучесть по МР
OK_IS_GC <- columns_GC(OK_IS) # Считаем текучесть по Группам кластеров

svod_IS <- rbind.fill(OK_IS_ter, OK_IS_MR, OK_IS_GC) # Сводим текучесть в одну таблицу

rm(OK_IS, OK_IS_GC, OK_IS_MR, OK_IS_ter)

# Собираем текучесть по Новым

OK_NEW <- read_excel(filename, sheet = "Новые")
OK_NEW <- OK_NEW[1:11]
OK_NEW <- merge(OK_NEW, df_nc, by.x = "ЦФО ID", by.y = "ЦФО ID", all.x = TRUE) # - надо проверить корректность мерджа
OK_NEW <- subset(OK_NEW, OK_NEW$Разница...11 == "FALSE")
OK_NEW <- OK_NEW[!is.na(OK_NEW$ЦФО),]

OK_NEW <- rename(OK_NEW, "ССЧ (усред. по дням) along Table (Down)" = "ССЧ (усред. по дням) along ЦФО ID",
                 Территория.x = Территория,
                 Макро = Макрорегион.x)

OK_NEW_ter <- columns_ter(OK_NEW)
OK_NEW_ter <- adorn_totals(OK_NEW_ter, "row", name = "ТС Пятёрочка")
OK_NEW_ter$Текучесть <- OK_NEW_ter$Уволенные / OK_NEW_ter$ССЧ

OK_NEW_MR <- columns_MR(OK_NEW)
OK_NEW_GC <- columns_GC(OK_NEW)

svod_NEW <- rbind.fill(OK_NEW_ter, OK_NEW_MR, OK_NEW_GC)

rm(OK_NEW, OK_NEW_GC, OK_NEW_MR, OK_NEW_ter)

# Собираем общую текучесть

OK_ALL <- read_excel(filename, sheet = "Общая текучесть")
OK_ALL <- OK_ALL[1:8]
OK_ALL <- merge(OK_ALL, df_nc, by.x = "ЦФО ID", by.y = "ЦФО ID", all.x = TRUE) # - надо проверить корректность мерджа
OK_ALL <- OK_ALL[!is.na(OK_ALL$`Количество Магазинов`),]
#OK_ALL <- subset(OK_ALL, OK_ALL$Лайк...6 != "NA")
OK_ALL <- rename(OK_ALL, "ССЧ (усред. по дням) along Table (Down)" = "ССЧ (усред. по дням) along ЦФО ID")

OK_ALL_ter <- columns_ter(OK_ALL)
OK_ALL_ter <- adorn_totals(OK_ALL_ter, "row", name = "ТС Пятёрочка")
OK_ALL_ter$Текучесть <- OK_ALL_ter$Уволенные / OK_ALL_ter$ССЧ

OK_ALL_MR <- columns_MR(OK_ALL)
OK_ALL_GC <- columns_GC(OK_ALL)

svod_ALL <- rbind.fill(OK_ALL_ter, OK_ALL_MR, OK_ALL_GC)


library(zip)
library(openxlsx)

df_excel <- createWorkbook('OK_result')

addWorksheet(df_excel, 'Список НК')
addWorksheet(df_excel, 'Список магазинов нового концепт')
addWorksheet(df_excel, 'Свод ИС')
addWorksheet(df_excel, 'Свод Новые')
addWorksheet(df_excel, 'Свод Общая')

writeData(df_excel, sheet = 1, df_nc, colNames = TRUE)
writeData(df_excel, sheet = 2, df_nc_final, colNames = TRUE)
writeData(df_excel, sheet = 3, svod_IS, colNames = TRUE)
writeData(df_excel, sheet = 4, svod_NEW, colNames = TRUE)
writeData(df_excel, sheet = 5, svod_ALL, colNames = TRUE)

saveWorkbook(df_excel, fileresultname, overwrite = TRUE)

rm(list = ls())

gc()
