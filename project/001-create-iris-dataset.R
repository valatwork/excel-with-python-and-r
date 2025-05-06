library(writexl)
library(janitor)
library(dplyr)

df <- iris |> clean_names()

l <- df |>
  clean_names() |>
  split(f = df$species)

lt <- c(l, iris = list(df))

write_xlsx(lt, path = "./project/data/iris_data.xlsx")