library("openxlsx")
library("tidytable")
library("haven")
library("lubridate")
library("hms")
library("stringr")
library("data.table")
library("tidyr")

print(getwd())
setwd('Excel')
EFILE <- read.xlsx("EFILE.xlsx", sheet = "Munka1", detectDates = TRUE)
dim(EFILE) # 109550 sor és 153 oszlop
setwd('../')
setwd('SPSS')
Input <- read_sav("Idomerleg_Kutatoszoba_20260313_imputalt_2.sav") # "Idomerleg_naplok_egyutt_plusz_mutatokkal_20260119_v3.sav"
Input <- as.data.frame(Input)
dim(Input) # 10955 sor és 4395 oszlop
setwd('../')
print(getwd())

Input$ISZAK <- as.integer(Input$ISZAK)

# values <- c("MACT", "SACTN", "ICTUSE", "WHERE", "ALONE", "WPARTNER", "WPARENT", "WCHILD", "WOTHERH", "WOTHERP")
# EFILE$DCOLUMN <- rep(values, length.out = nrow(EFILE))


round_classic <- function(x, n = 0) {

  posneg = sign(x)
  z = abs(x) * 10 ^ n
  z = z + 0.5 + sqrt(.Machine$double.eps)
  z = trunc(z)
  z = z / 10 ^ n

  return(z * posneg)
}

EFILE <- EFILE %>% mutate(across(10:153, as.character))
setDT(EFILE)
# setkey(EFILE, LAKAZON, ISZAK, DCOLUMN)
setDT(Input)


Input[3540, c("LAKAZON", "ISZAK")]
View(EFILE[is.na(EFILE$TS_145) == FALSE, c("LAKAZON", "ISZAK")])
View(Input[Input$LAKAZON == "L01535467977", ]) # 3383. utolsó előtti tevékenység 22:18-kor befejeződött, de az utolsó már 22:14-kor megkezdődött!!!
View(Input[Input$LAKAZON == "L01535469890", ])
# L01535467977

# Előre definiáljuk a keresett oszlopokat, hogy ne a ciklusban kelljen stringet fűzni
val_cols <- c("FOTEV_KOD_", "where_", "ALONE", "FIBC113_", "FIBC114_", "FIBC115_", "FIBC116_", "FIBC117_")
d_cols <- c("MACT", "WHERE", "ALONE", "WPARTNER", "WPARENT", "WCHILD", "WOTHERH", "WOTHERP")
for (i in c(1:nrow(Input))) { # 34:34 # nrow(Input) # c(1, 34) # 3540 nem jó!!!  # 3541:nrow(Input)  #3383 # 1:nrow(Input)
  if(i %% 100 == 0) cat("Feldolgozás:", i, "/", nrow(Input), "\n")
  
  TS_START <- 1
  previous <- 0
  curr_lakazon <- Input$LAKAZON[i]
  # print(curr_lakazon)
  curr_iszak <- Input$ISZAK[i]
  
  for (j in 1:65) {
    col_from <- paste0("FIBC126_", j)
    col_to   <- paste0("FIBC111_", j)
    
    if (is.na(Input[[col_from]][i]) | Input[[col_from]][i] == "") break
    if (is.na(Input[[col_to]][i]) | Input[[col_to]][i] == "") break
    
    from <- round_hms(as_hms(as.numeric(hm(Input[[col_from]][i]))), 600)
    to   <- round_hms(as_hms(as.numeric(hm(Input[[col_to]][i]))), 600)
    
    diff_min <- as.numeric(as.duration(to - from), "minutes") / 10
    if (!is.na(diff_min) && diff_min < 0) diff_min <- diff_min + 144
    
    rounded_diff <- round_classic(diff_min)
    
    if (rounded_diff == 0){ 
      # print("Különbség 0:")
      # print(curr_lakazon)
      # print(curr_iszak)
      # print(j)
      next 
      
    }
    
    # print(paste(from, to, diff_min, rounded_diff, Input[[paste0(val_cols[1], j)]][i], sep = ", "))
    
    if (rounded_diff > 0) {
      # Cél oszlopok nevei (TS_001, TS_002...)
      target_ts <- paste0("TS_", str_pad(TS_START:(TS_START + rounded_diff - 1), 3, pad = "0"))
      
      # Vektorizált kitöltés: egyszerre az összes DCOLUMN típusra
      for (idx in seq_along(d_cols)) {
        val <- Input[[paste0(val_cols[idx], j)]][i] # as.character()
        # data.table szűrés és in-place módosítás (:=)
        EFILE[LAKAZON == curr_lakazon & ISZAK == curr_iszak & DCOLUMN == d_cols[idx], 
              (target_ts) := val]
      }
      
      TS_START <- TS_START + rounded_diff
    }
    
    if (Input[[col_to]][i] == "04:00") break
  }
}


write.xlsx(EFILE, "EFILE_v1_fast_20260407.xlsx", overwrite = TRUE)

# SAVE_EFILE <- EFILE

library(data.table)
setDT(Input)

# 1. FŐTEVÉKENYSÉGEK (1-64)
# Kigyűjtjük a FIBC128_1...64 oszlopokat
cols_128 <- grep("^FIBC128_[0-9]+$", names(Input), value = TRUE)
fotev <- melt(Input, 
              id.vars = c("LAKAZON", "TEV", "ISZAK"), 
              measure.vars = cols_128,
              variable.name = "id_fotev",
              value.name = "FIBC128", 
              na.rm = TRUE)
# Sorszám kinyerése (pl. "FIBC128_1" -> 1)
fotev[, sorszam := as.numeric(gsub("FIBC128_", "", id_fotev))]

# 2. ALTEVÉKENYSÉGEK (1A...64H)
# Kigyűjtjük az oszlopokat típusonként
c_altev <- grep("^ALTEV_KOD_", names(Input), value = TRUE)
c_125   <- grep("^FIBC125_",   names(Input), value = TRUE)
c_127   <- grep("^FIBC127_",   names(Input), value = TRUE)
c_124   <- grep("^FIBC124_",   names(Input), value = TRUE)

altev <- melt(Input, 
              id.vars = c("LAKAZON", "TEV", "ISZAK"), 
              measure.vars = list(c_altev, c_125, c_127, c_124),
              value.name = c("ALTEV", "FIBC125", "FIBC127", "FIBC124"),
              na.rm = TRUE)

# Kinyerjük a sorszámot az ALTEV kód nevéből (pl. "ALTEV_KOD_1A" -> 1)
# Ehhez a 'variable' oszlopot használjuk, amit a melt automatikusan létrehoz
altev[, sorszam := as.numeric(gsub("\\D", "", c_altev[variable]))]

# 3. ÖSSZEKAPCSOLÁS
# Itt történik a 1:8-as másolás: a sorszám (1-64) alapján
Input_Values <- merge(altev, fotev[, .(LAKAZON, TEV, ISZAK, sorszam, FIBC128)], 
                      by = c("LAKAZON", "TEV", "ISZAK", "sorszam"), 
                      all.x = TRUE)

# 4. TISZTÍTÁS
Input_Values <- Input_Values[!is.na(ALTEV) & ALTEV != ""]
setorder(Input_Values, LAKAZON, TEV, ISZAK, sorszam)
Input_Values <- Input_Values[, .(LAKAZON, ISZAK, ALTEV, FIBC128, FIBC125, FIBC127, FIBC124)]
dim(Input_Values) # 77314 sor és 7 oszlop


# 1. A kerekítő függvény (HH:MM formátumhoz)
round_time_text <- function(x) {
  # Ha alapból NA vagy nem tartalmaz kettőspontot, rögtön NA-t adunk vissza
  if (is.na(x) || x == "" || !grepl(":", x)) return(NA)
  
  # Szétválasztás
  parts <- as.numeric(unlist(strsplit(as.character(x), ":")))
  
  # Ellenőrizzük, hogy mindkét rész (óra, perc) szám-e
  if (any(is.na(parts)) || length(parts) < 2) return(NA)
  
  ora <- parts[1]
  perc <- parts[2]
  
  # 10 perces kerekítés
  perc_ker <- round(perc / 10) * 10
  
  if (perc_ker == 60) {
    ora <- ora + 1
    perc_ker <- 0
  }
  
  # Formázás vezető nullával
  sprintf("%02d:%02d", ora %% 24, perc_ker)
}


round_end_time <- function(x) {
  # Alapvető ellenőrzések
  if (is.na(x) || x == "" || !grepl(":", x)) return(NA)
  
  # Idő szétbontása és számmá alakítása
  parts <- as.numeric(unlist(strsplit(as.character(x), ":")))
  
  # Ha nem sikerült számot kinyerni (pl. " : "), NA-t adunk vissza
  if (any(is.na(parts)) || length(parts) < 2) return(NA)
  
  ora <- parts[1]
  perc <- parts[2]
  
  # 5 perctől felfelé kerekítés a következő 10-esre
  perc_ker <- ifelse(perc %% 10 >= 5, (floor(perc / 10) + 1) * 10, floor(perc / 10) * 10)
  
  # Ha 60 perc lett, órát váltunk
  if (!is.na(perc_ker) && perc_ker == 60) {
    ora <- ora + 1
    perc_ker <- 0
  }
  
  # HH:MM formátum vezető nullákkal
  sprintf("%02d:%02d", ora %% 24, perc_ker)
}

# 2. Új oszlopok létrehozása
Input_Values[, FIBC127_RND := sapply(FIBC127, round_time_text)]
Input_Values[, FIBC124_RND := sapply(FIBC124, round_end_time)]

# View(Input[LAKAZON == "L01535428079" & ISZAK == "7", ])

# View(Input[LAKAZON == "L01535428675" & ISZAK == "7", ])



get_10min_index <- function(time_str) {
  if (is.na(time_str) || time_str == "") return(NA)
  
  parts <- as.numeric(unlist(strsplit(time_str, ":")))
  # Összes perc éjféltől számítva
  total_mins <- parts[1] * 60 + parts[2]
  
  # A nap 04:00-kor (240. perc) kezdődik
  start_offset <- 240
  
  diff_mins <- total_mins - start_offset
  
  # Ha negatív (00:00 és 03:59 között), adjunk hozzá egy teljes napot (1440 perc)
  if (diff_mins < 0) {
    diff_mins <- diff_mins + 1440
  }
  
  # 10 perces blokk sorszáma (1-től 144-ig)
  return(floor(diff_mins / 10) + 1)
}

Input_Values[, FIBC127_INT := sapply(FIBC127_RND, get_10min_index)]
Input_Values[, FIBC124_INT_NYERS := sapply(FIBC124_RND, get_10min_index)]

# 2. Intelligens korrekció: csak akkor vonunk le, ha a végpont későbbi
Input_Values[, FIBC124_INT := ifelse(FIBC124_INT_NYERS > FIBC127_INT, 
                                     FIBC124_INT_NYERS - 1, 
                                     FIBC127_INT)]

# 3. Éjféli korrekció (speciális eset, ha átnyúlik a napon: pl. 144-ről 1-re)
Input_Values[FIBC124_INT_NYERS == 1 & FIBC127_INT > 1, FIBC124_INT := 144]

# 4. BLOKK_DB frissítése (+1 kell, mert ha 27-től 27-ig tart, az 1 blokk)
Input_Values[, BLOKK_DB := (FIBC124_INT - FIBC127_INT)]
Input_Values[BLOKK_DB < 0, BLOKK_DB := BLOKK_DB + 144]
Input_Values[, BLOKK_DB := BLOKK_DB + 1]

# 5. DURATION_BLOCKS szinkronizálása
Input_Values[, DURATION_BLOCKS := BLOKK_DB]

# 2. Új oszlop létrehozása a kiválasztáshoz
# Sorrendbe tesszük: Lakáson és időszeleten belül a leghosszabb legyen elöl
Input_Values[, ISZAK := factor(ISZAK, levels = c(4, 7, 10, 1))]
setorder(Input_Values, LAKAZON, ISZAK, FIBC127_INT, -DURATION_BLOCKS)

# 3. Megjelöljük azokat a sorokat, amik a leghosszabbak (holtversenynél az elsőt)
Input_Values[, KEEP_SACTN := FALSE]
Input_Values[, KEEP_SACTN := .I == .I[1], by = .(LAKAZON, ISZAK, FIBC127_INT)]

# 4. ICTUSE jelző (ezt minden sorra ki tudjuk számolni)
Input_Values[, IS_ICT := as.numeric(FIBC128 == 1 | FIBC125 == 1)]


write.xlsx(Input_Values, "Altevékenységek_duration.xlsx", overwrite = TRUE)
