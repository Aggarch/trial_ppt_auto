

library(tidyverse)
library(officer)
library(quantmod)
library(flextable)


# flextable only one compatible:
# https://ardata-fr.github.io/flextable-gallery/gallery/


# Simulated  --------------------------------------------------------------


read_pptx("trial.pptx") %>% 
   officer::layout_properties()


pres <- read_pptx("trial.pptx")



company <- "MTCARS"
summary <- mtcars %>% head(5) %>% flextable::flextable( )
ticker  <- "SWK"
data    <- "123456..."

company2 <- "MTCARS"
summary2 <- mtcars %>% head(3)%>% flextable::flextable()
ticker2  <- "SWK"
data2    <- "123456..."


first_title = "Growth Investments   North America              Aug YTD & F7 FY Outlook" 


my_pres <- pres %>%
   # remove_slide(index = 1) %>%
   
   add_slide(layout = "cover", master = "tema") %>%
   ph_with(value = first_title, location = ph_location_label(ph_label = "first_title")) %>%
   
   
   add_slide(layout = "temp", master = "tema") %>%
   ph_with(value = company, location = ph_location_label(ph_label = "company")) %>%
   ph_with(value = summary, location = ph_location_label(ph_label = "summary")) %>%
   ph_with(value = ticker, location = ph_location_label(ph_label = "ticker")) %>%
   ph_with(value = data, location = ph_location_label(ph_label = "data")) %>%
   
 
   add_slide(layout = "temp2", master = "tema") %>%
   ph_with(value = company2, location = ph_location_label(ph_label = "company2")) %>%
   ph_with(value = summary2, location = ph_location_label(ph_label = "summary2")) %>%
   ph_with(value = ticker2, location = ph_location_label(ph_label = "ticker2")) %>%
   ph_with(value = data2, location = ph_location_label(ph_label = "data2"))


   
print(my_pres, glue::glue("example.dummy {Sys.Date()}.pptx"))



# Real Example ------------------------------------------------------------

# To make it work, Create templates as "investment_review_NA.pptx" for each scenario
# template was created, using existing original file copy, try to tackle the slide master 
# creation to make it work successfully  
# first step its to create the Master Sliders, 
# the new file will be  pull up in a diferent file, same format, so we can save 
# the original and only one unique Master Slider for diferents scenarios
# Slider master can be explore in the output file, but its not a recomendation 
# changes in the master slide should be produce directly in the source of the master slider. 
# Layouts can be copied from the original files to the new Sliders Masters,
# SBD Sliders masters, contains the Covers and the blacks sheets. 
# additional UI's structures are accessible thru the existing examples.
# only way to modify Master Slides It's adding new elements on the structure,
# to add new elements go to the MSlide and 'Insert Placeholder', then; go to home,
# select the option Arrange//Select Panel & and modify the name of the new element by 
# clicking on it, rename elements with a very short and descriptive name, 
# rename master and each slide to the corresponding number. 
# templates order is related to the sequence of the algorythm. 
# logo details for the top and the bottom are easy to copy from Slide master 2
# also covers are accessible thru Master View. 
# if at opening example file looks bad, "Close Master View"

# Pending >>> Find the way to better reproduce the tables that the PBI Produce.





library(tidyverse)
library(officer)
library(quantmod)
library(gt)


read_pptx("investment_review_NA.pptx") %>% 
   officer::layout_properties()


# read_pptx("investment_review_NA.pptx") %>%
#    officer::layout_properties() %>% as_tibble() %>% filter(master_name == "master_1") %>%
#    filter(grepl("first|second", ph_label))


pres <- officer::read_pptx("investment_review_NA.pptx")


first_title = "Growth Investments North America Aug YTD & F7 FY Outlook" 
second_title = "Executive Summary" 

invest_pres <- pres %>%
   # remove_slide(index = 1) %>%
   add_slide(layout = "1", master = "master_1") %>%
   ph_with(value = first_title, location = ph_location_label(ph_label = "first_title")) %>% 
   
   add_slide(layout = "2", master = "master_1") %>%
   ph_with(value = second_title, location = ph_location_label(ph_label = "second_title"))
   
 

print(invest_pres, glue::glue("Example Investment Review {Sys.Date()}.pptx"))
