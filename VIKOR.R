#Load Kriteria
  kr <- xlsx::read.xlsx2(file ="./VIKOR.xlsx", sheetName = "Kriteria")
  kr

# Mengubah Tipe Data
  kr$id_kriteria <- as.numeric(kr$id_kriteria)
  kr$bobot <- as.numeric(kr$bobot)/100

  str(kr)

# Load Data
  library(xlsx)
  df <- xlsx::read.xlsx2(file ="./VIKOR.xlsx", sheetName = "Data")

  df
# lihat dataframe
  class(df)
  str(df)

# Mengubah tipe data
  kolom_char <- c("Kode","Nama")
  kolom_numeric <- c("C1", "C2", "C3", "C4")

  df[,kolom_char] <- sapply(df[,kolom_char], as.character)
  df[,kolom_numeric] <- sapply(df[,kolom_numeric], as.numeric)

  str(df)

#1. Membuat Matrix Nilai Penilaian

  library(dplyr)
  vik_sample <- select(df, kolom_numeric)
  str(vik_sample)

  row.names(vik_sample) <- df$Kode
  colnames(vik_sample) <- kr$kriteria
  str(vik_sample)

#2. Normalisasi Data

# Mencari nilai minimum & maksimum dari setiap kolom
  min_values <- apply(vik_sample, 2, min)
  max_values <- apply(vik_sample, 2, max)

# Membuat dataframe baru untuk menyimpan hasil normalisasi
  vik_normalized <- data.frame(matrix(ncol = ncol(vik_sample), nrow = nrow(vik_sample)))
  colnames(vik_normalized) <- colnames(vik_sample)
  row.names(vik_normalized) <- row.names(vik_sample)

# Melakukan normalisasi dengan iterasi pada setiap kolom min max
  # Melakukan normalisasi dengan iterasi pada setiap kolom min max
  for (i in 1:ncol(vik_sample)) {
    if (kr$tipe[i] == "max") {
       vik_normalized[,i] <- (max_values[i] - vik_sample[,i]) / (max_values[i] - min_values[i])
    if(any(is.nan(vik_normalized[,i]))){
      vik_normalized[,i] <- 0
    }
  }else{
    vik_normalized[,i] <- 1-((max_values[i] - vik_sample[,i]) / (max_values[i] - min_values[i]))
    if(any(is.nan(vik_normalized[,i]))){
      vik_normalized[,i] <- 0
    }}
  }
    
  vik_normalized <- vik_normalized %>% mutate_all(funs(round(.,3)))

# Perkalian nilai yang telah dinormalisasi dengan bobot
  vik_weighted <- data.frame(matrix(ncol = ncol(vik_normalized), nrow = nrow(vik_normalized)))
  colnames(vik_weighted) <- colnames(vik_sample)
  row.names(vik_weighted) <- row.names(vik_sample)

    for (i in 1:ncol(vik_normalized)) {
      vik_weighted[,i] <- (vik_normalized[,i] * kr$bobot[i])
    }

  vik_weighted <- vik_weighted %>% mutate_all(funs(round(.,3)))

#3.Menghitung Utility Measure S dan R
require(Hmisc)
  S <- R <- as.array(nrow(vik_weighted))
  
  S <- rowSums(vik_weighted)
  R <- apply(vik_weighted[], 1,max)

#4. Menghitung Nilai Indeks VIKOR
  s_min <- min(S)
  s_max <- max(S)
  r_max <- max(R)
  r_min <- min(R)
  
  Q <- as.array(nrow(vik_weighted))
  v <- 0.5

  SQ <- function(S,R,v,Q)
  {
  
    for (i in nrow(Q)) {
    if (v==0){
      Q <- (R-min(R))/(max(R)-min(R))
      }else if (v==1)
      {Q <- (S-min(S))/(max(S)-min(S))
      }else{
        Q<- v*(S-min(S))/(max(S)-min(S))+(1-v)*(R-min(R))/(max(R)-min(R))
      }
    }
    return(Q)
  }
  
  Q_utama <- round(SQ(S,R,v,Q),4)
  print(Q_utama)
  
  
  RQ <- function(Q){
      if( (Q == "NaN") || (Q == "Inf")){
        RankingQ <- rep("-",nrow(decision))
      }else{
        RankingQ <- rank(Q, ties.method= "first") 
      }
    return(RankingQ)
  }
  
  RankingQ <- RQ(Q_utama)
  print(RankingQ)
  
#5. Ranking the alternatives
  vik_hasil <- data.frame(Kode = df$Kode, Nama = df$Nama, S = S, R = R, Q = Q_utama, Ranking = RankingQ)
  
  print(vik_hasil)

# Menentukan Solusi Kompromi
  DQ = 1/(nrow(df)- 1)
  str(DQ)
  
  selisih <- round(vik_hasil$Q[vik_hasil$Ranking==2] - vik_hasil$Q[vik_hasil$Ranking==1],3)
  print(selisih)
  kondisi_1 <- 0
    if (selisih<DQ) {
      print('Kondisi 1 Tidak Terpenuhi')
    }else{
      kondisi_1 <- 1
      print('Kondisi 1 Terpenuhi')
    }

# Pembuktian kondisi Acceptable Stability in Decision Making
  sQ <- vector()
  v <- c(0.4, 0.5, 0.6)
  #menghitung nilai Q untuk V=[0.4] (by Vote)  
  Q_vote <- round(SQ(S,R,v[1],Q),4)
  QV_ranked <- data.frame(Kode = df$Kode, Nama = df$Nama, S = S, R = R, Q = Q_vote, Ranking = RQ(Q_vote))
  print(QV_ranked)

  # Menghitung nilai Q untuk V=[0.6] (by majority rule)  
  Q_major <- round(SQ(S,R,v[2],Q),4)
  QM_ranked <- data.frame(Kode = df$Kode, Nama3 =df$Nama, S = S, R = R, Q = Q_major, Ranking = RQ(Q_major))
  print(QM_ranked)  
  
 # Membuat tabel hasil perbandingan
  QV_ranked <- arrange(QV_ranked,Ranking)
  QM_ranked <- arrange(QM_ranked,Ranking)
  vik_hasil <- arrange(vik_hasil,Ranking)
  result <- cbind(QV_ranked, vik_hasil, QM_ranked)
  row.names(result) <- NULL
  print(result)
  
  # Pembuktian lebih lanjut
  kondisi_2 <- 0
  if((vik_hasil$Nama[1]== QV_ranked$Nama[1]) && (vik_hasil$Nama[1] == QM_ranked$Nama[1])) {
    kondisi_2 <- 1
    print('Kondisi 2 Terpenuhi')
  } else {
    print('Kondisi 2 Tidak Terpenuhi')
  }
  
  cat(' Berdasarkan Hasil Pembuktian kedua kondisi dapat diketahui bahwa :',"\n",'Kedua kondisi tersebut : ',"\n", if(kondisi_1 != 1){
    'Kondisi satu tidak terpenuhi'
  }else if(kondisi_2 != 1){
    'Kondisi dua tidak terpenuhi'
  }else{
    'Terpenuhi'
  })