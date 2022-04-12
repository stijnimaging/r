library(officer)
library(readxl)

rm(list=ls())
## Inladen Excel bestand met uitslag; let op format! 
## Enkel totaal scores
setwd("~/Documents/KempenOptocht_2020/UITSLAG")
df <- read_excel(file.choose())
head(df)

#df$Nr. <- stri_pad_left(df$Nr., 4, 0)
## Toevoegen logische nummering
df$ID <- seq.int(nrow(df))
## Omgooien laag naar hoog (nummer 1 als laatste!)
df <- df[order(-df$ID),] 
df$Uitsl <- df$WAGENS
df$Plaats <- df$...5
df$Onderwerp <- df$...4
df$Totaal <- df$...6
df$Nr. <- df$...2
df$Naam <- df$...3

#df$Naam <- df$X__1 
#df$start_nr <- df$`Wagens`
#df$Plaats <- df$X__3
#df$Onderwerp <- df$X__2
#df$Uitslag <- df$X__5
#df$Totaal <- df$X__4

## Inladen template (via Master view te bewerken!)
doc <- read_pptx("template_KO_geel5.pptx")
## Wat zit er in het template Uitslag_1?
layout_summary(doc)
layout_properties ( x = doc, layout = "Uitslag_1", master = "KO_template" )
# Label en ID in Uitslag_2:
# Title 9 = 31e kempenoptocht
# TextBox 10 = tweet mee ko2019
# Index 6 = punten
# Index 5? = plaats
# Index 3 = onderwerp
# Title = naam
# Textbox 20, ID 21 = plek

# Label en ID in Uitslag_tussen:
# TextBox 1, ID 2 = plek

#sprintf("This is where a %s goes.", a)
apply(df, 1, function(row) {
  naam <- row["Naam"]
  onderwerp <- row["Onderwerp"]
  plaats <- row["Plaats"]
  punten <- row["Totaal"]
  nummer <- row["Uitsl"]
  start_nr <- row["Nr."]
  # Inladen foto vanuit map /fotos/no_##.jpg
  img.file <- file.path( getwd(),"fotos", sprintf("no_%s.jpg",start_nr))
  message("Now targetting ",img.file)
  if( file.exists(img.file) ){
  ## Nieuwe tussen slide toevoegen in een loop
  doc <- add_slide(doc, layout = "Uitslag_tussen", master = "KO_template")
  ## Insert NAR plaatje
  doc <- ph_with_img_at(x = doc, src = file.path( getwd(), "NAR_correct.png"), left = 0.1, top = 0.5, height = 1.844037, width = 1.5)
  ## Insert positie
  ## Plek, ID = 2
  doc <- ph_with_text(x = doc, type = "body", index = 5, str = nummer)  # 3/4 correct?
    
  ## Nieuwe slide opbwouen en toevoegen in een loop
  doc <- add_slide(doc, layout = "Uitslag_1", master = "KO_template")
  ## Insert NAR plaatje
  doc <- ph_with_img_at(x = doc, src = file.path( getwd(), "NAR_correct.png"), left = 0.1, top = 0.5, height = 1.844037, width = 1.5)
  ## Insert plaatje
  doc <- ph_with_img_at(x = doc, src = img.file, left = 2, top = 1, height = 4.5, width = 6)
  ## Insert names, thema, dorp, resultaat
  doc <- ph_with_text(x = doc, type = "title", str = naam) # OK
  ## Plaats
  doc <- ph_with_text(x = doc, type = "body", index = 1, str = plaats) # Geen 2/5/6/7/9
  ## Punten
  doc <- ph_with_text(x = doc, type = "body", index = 4, str = nummer) # OK
  ## Plaats
  doc <- ph_with_text(x = doc, type = "body", index = 3, str = punten) # OK
  ## Onderwerp
  doc <- ph_with_text(x = doc, type = "body", index = 8, str = onderwerp) # OK 

  }
})

print(doc, target = "KO_PPTX_KO20.pptx" )

