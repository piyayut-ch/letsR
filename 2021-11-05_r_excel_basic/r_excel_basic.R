## -----------------------------------------------------------------------------
# install.packages(xfun) # package ช่วยในการ install และ load library
pkgs_office <- c(
  "tidyverse",
  "readxl", # นำเข้าไฟล์ excel
  "writexl" # สร้างไฟล์ excel จาก data frame
)

xfun::pkg_attach2(pkgs_office, message = FALSE)


## -----------------------------------------------------------------------------
# official document
# help(read_excel)


## -----------------------------------------------------------------------------
# ex1: อ่านไฟล์ excel
read_excel("data/mtcars.xlsx")


## -----------------------------------------------------------------------------
# ex2: อ่านไฟล์ excel กำหนด sheet
read_excel("data/gapminder.xlsx")


## -----------------------------------------------------------------------------
read_excel("data/gapminder.xlsx", sheet = 2)


## -----------------------------------------------------------------------------
read_excel("data/gapminder.xlsx", sheet = "2007")


## -----------------------------------------------------------------------------
# ex3: อ่านไฟล์ excelหลาย sheet
# ขั้นตอนแรก คือ ต้องรู้ว่า sheet ชื่ออะไรบ้าง
(gapminder_sheets <- readxl::excel_sheets("data/gapminder.xlsx"))


## -----------------------------------------------------------------------------
# ขั้นตอนสอง คือ การ loop ให้อ่านทีละ sheet แล้วรวมกันเป็น data.frame เดียวกัน
df <- data.frame() 
for (i in gapminder_sheets) {
  df_ <- read_excel("data/gapminder.xlsx", sheet = i)
  df <- bind_rows(df, df_)
}
df


## -----------------------------------------------------------------------------
# ทางเลือกการทำ loop ใช้ function map และ reduce จาก package purrrr
map(gapminder_sheets, function(i) {
   read_excel("data/gapminder.xlsx", sheet = i)
}) |> reduce(bind_rows)
# |> คือ pipe operator ทำหน้าที่เหมือน %>%
# ต่างกันตรงที่ |> มาพร้อมกับ base R version 4.1 สามารถใช้ได้เลย
# ส่วน %>% ต้อง library(tidyverse) ก่อน ถึงจะใช้ได้


## -----------------------------------------------------------------------------
# ex4: การอ่านไฟล์ excelที่เป็น multiple heading
read_excel("data/ข้าวนาปี62.xlsx")


## -----------------------------------------------------------------------------
# ใช้ skip
read_excel("data/ข้าวนาปี62.xlsx", skip = 2)


## -----------------------------------------------------------------------------
# ใช้ col_names แนะนำให้ตั้งชื่อ column โดยใช้ snake_case
# https://raw.githubusercontent.com/allisonhorst/stats-illustrations/master/other-stats-artwork/coding_cases.png
col_names <- c("province", "area_plant", "area_harvest",
              "production", "yield_plant", "yield_harvest")
read_excel("data/ข้าวนาปี62.xlsx", skip = 3, col_names = col_names)


## -----------------------------------------------------------------------------
# ex5: การอ่านไฟล์ excelที่มีค่า na
col_names <- c("province", "area_plant", "area_harvest",
              "production", "yield_plant", "yield_harvest")
read_excel("data/ข้าวนาปี62_fakena.xlsx", skip = 3, col_names = col_names) |> 
  tail(10)
# ปัญหา 1 area_harvest production type เป็น character
# ปัญหา 2 yield_plant รหัส -999 ไม่ได้ แปลงให้ เป็น na


## -----------------------------------------------------------------------------
na_values <- c("", "-", "NA", "-999")
read_excel("data/ข้าวนาปี62_fakena.xlsx", skip = 3, col_names = col_names,
           na = na_values) |> tail(10)


## -----------------------------------------------------------------------------
# official document
# help(write_xlsx)


## -----------------------------------------------------------------------------
write_xlsx(mtcars, "output/mtcars.xlsx")


## -----------------------------------------------------------------------------
write_xlsx(mtcars |> rownames_to_column(), "output/mtcars_with_rownames.xlsx")

