# Kural Tabanlı Sınıflandırma ile Potansiyel Müşteri Getirisi Hesaplama
# Gezinomi yaptığı satışların bazı özelliklerini kullanarak seviye tabanlı (level based) yeni satış tanımları
# oluşturmak ve bu yeni satış tanımlarına göre segmentler oluşturup bu segmentlere göre yeni gelebilecek müşterilerin şirkete
# ortalama ne kadar kazandırabileceğini tahmin etmek istemektedir.
# Örneğin: Antalya’dan Herşey Dahil bir otele yoğun bir dönemde gitmek isteyen bir müşterinin ortalama ne kadar kazandırabileceği belirlenmek isteniyor.
# GÖREV 1: Aşağıdaki soruları yanıtlayınız.

# Soru 1: miuul_gezinomi.xlsx dosyasını okutunuz ve veri seti ile ilgili genel bilgileri gösteriniz.
import pandas as pd
pd.set_option("display.max_columns", None)
df = pd.read_excel('C:/Users/asus/Desktop/miuul/miuul_gezinomi.xlsx')
print(df.head())
print(df.shape)
print(df.info())

# Soru 2: Kaç unique şehir vardır? Frekansları (tekrar etme sıklığı) nedir?
print(df["SaleCityName"].nunique())
print(df["SaleCityName"].value_counts())

# Soru 3: Kaç unique Concept vardır?
print(df["ConceptName"].nunique())

# Soru 4: Hangi Concept'dan kaçar tane satış gerçekleşmiş?
print(df["ConceptName"].value_counts())

#Soru 5: Şehirlere göre satışlardan toplam ne kadar kazanılmış?
df.groupby('SaleCityName').agg({'Price': 'sum'})

# Soru 6: Concept türlerine göre göre ne kadar kazanılmış?
df.groupby('ConceptName').agg({'Price': 'sum'})

# Soru 7: Şehirlere göre PRICE ortalamaları nedir?
df.groupby('SaleCityName').agg({'Price': 'mean'})

# Soru 8: Conceptlere  göre PRICE ortalamaları nedir?
df.groupby('ConceptName').agg({'Price':'mean'})

# Soru 9: Şehir-Concept kırılımında PRICE ortalamaları nedir?
df.groupby(['SaleCityName','ConceptName']).agg({'Price': 'mean'})
#veya
df.groupby(by=["SaleCityName", 'ConceptName']).agg({"Price": "mean"})

# GÖREV 2: sale_checkin_day_diff değişkenini EB_Score adında yeni bir kategorik değişkene çeviriniz. Aralıkları ikna edici şekilde oluşturunuz.
# Örneğin: ‘0_7’, ‘7_30', ‘30_90', ‘90_max’ aralıklarını kullanabilirsiniz.
# Bu aralıklar için "Last Minuters", "Potential Planners", "Planners", "Early Bookers“ isimlerini kullanabilirsiniz

bins = [-1, 7, 30, 90, df["SaleCheckInDayDiff"].max()]
labels = ["Last Minuters", "Potential Planners", "Planners", "Early Bookers"]

df["EB_Score"] = pd.cut(df["SaleCheckInDayDiff"], bins, labels=labels)
df.head(50).to_excel("eb_scorew.xlsx", index=False)
df.head()

# GÖREV 3: Şehir,Concept, [EB_Score,Sezon,CInday] kırılımında ücret ortalamalarına ve frekanslarına bakınız
# Şehir-Concept-EB Score kırılımında ücret ortalamaları
df.groupby(by=["SaleCityName", 'ConceptName', "EB_Score" ]).agg({"Price": ["mean", "count"]})

# Şehir-Concept-Sezon kırılımında ücret ortalamaları
df.groupby(by=["SaleCityName", "ConceptName", "Seasons"]).agg({"Price": ["mean", "count"]})

# Şehir-Concept-CInday kırılımında ücret ortalamaları
df.groupby(by=["SaleCityName", "ConceptName", "CInDay"]).agg({"Price": ["mean", "count"]})


#############################################
# GÖREV 4: City-Concept-Season kırılımın çıktısını PRICE'a göre sıralayınız.
#############################################
# Önceki sorudaki çıktıyı daha iyi görebilmek için sort_values metodunu azalan olacak şekilde PRICE'a uygulayınız.
# Çıktıyı agg_df olarak kaydediniz.
agg_df = df.groupby(["SaleCityName", "ConceptName", "Seasons"]).agg({"Price": "mean"}).sort_values("Price", ascending=False)
agg_df.head(20)

# GÖREV 5: Indekste yer alan isimleri değişken ismine çeviriniz.
#############################################
# Üçüncü sorunun çıktısında yer alan PRICE dışındaki tüm değişkenler index isimleridir.
# Bu isimleri değişken isimlerine çeviriniz.
# İpucu: reset_index()
agg_df.reset_index(inplace=True)

agg_df.head()

# GÖREV 6: Yeni level based satışları tanımlayınız ve veri setine değişken olarak ekleyiniz.
#############################################
# sales_level_based adında bir değişken tanımlayınız ve veri setine bu değişkeni ekleyiniz.
agg_df['sales_level_based'] = agg_df[["SaleCityName", "ConceptName", "Seasons"]].agg(lambda x: '_'.join(x).upper(), axis=1)
agg_df['sales_level_based'] = agg_df[['SaleCityName', 'ConceptName', 'Seasons']]. agg(lambda x: '_'.join(x).upper(), axis=1)
agg_df.head()

# GÖREV 7: Personaları segmentlere ayırınız.
#############################################
# PRICE'a göre segmentlere ayırınız,
# segmentleri "SEGMENT" isimlendirmesi ile agg_df'e ekleyiniz
# segmentleri betimleyiniz
agg_df["SEGMENT"] = pd.qcut(agg_df["Price"], 4, labels=["D", "C", "B", "A"])
agg_df.head(30)
agg_df.groupby("SEGMENT").agg({"Price": ["mean", "max", "sum"]})


# GÖREV 8: Oluşan son df'i price değişkenine göre sıralayınız.
# "ANTALYA_HERŞEY DAHIL_HIGH" hangi segmenttedir ve ne kadar ücret beklenmektedir?
#############################################
agg_df.sort_values(by="Price")


new_user = "ANTALYA_HERŞEY DAHIL_HIGH"
agg_df[agg_df["sales_level_based"] == new_user]