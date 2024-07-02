import pandas as pd
from gtts import gTTS
import os
from tkinter import Tk, filedialog

# Tkinter arayüzünü kapalı tutmak için başlat
root = Tk()
root.withdraw()

# Kullanıcıdan Excel dosyasını seçmesini iste
file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
if not file_path:
    print("No file selected. Exiting.")
    exit()

# Kullanıcıdan kaydedilecek dizini seçmesini iste
output_directory = filedialog.askdirectory(title="Select Output Directory")
if not output_directory:
    print("No output directory selected. Exiting.")
    exit()

# Kaydedilecek dizin yoksa oluştur
os.makedirs(output_directory, exist_ok=True)

# Excel dosyasını oku
try:
    df = pd.read_excel(file_path)
    print("Excel dosyası başarıyla okundu.")
except Exception as e:
    print(f"Excel dosyasını okurken hata oluştu: {e}")

# Her bir metni kontrol et ve yazdır
for index, row in df.iterrows():
    english_text = row['English'] if 'English' in row else None
    turkish_text = row['Turkish'] if 'Turkish' in row else None
    
    # Dosya adları için geçersiz karakterleri kaldır
    file_safe_english_text = ''.join(e for e in english_text if e.isalnum()) if english_text else None
    
    print(f"Index: {index}, English: {english_text}, Turkish: {turkish_text}")
    
    # İngilizce metni sesli dosyaya çevir
    if pd.notna(english_text):
        try:
            tts_en = gTTS(text=english_text, lang='en')
            english_file_path = os.path.join(output_directory, f'{file_safe_english_text}.mp3')
            tts_en.save(english_file_path)
            print(f"English text saved to: {english_file_path}")
        except Exception as e:
            print(f"Error saving English text: {e}")
    
    # Türkçe metni sesli dosyaya çevir
    if pd.notna(turkish_text):
        try:
            tts_tr = gTTS(text=turkish_text, lang='tr')
            turkish_file_path = os.path.join(output_directory, f'{file_safe_english_text}_turkish.mp3')
            tts_tr.save(turkish_file_path)
            print(f"Turkish text saved to: {turkish_file_path}")
        except Exception as e:
            print(f"Error saving Turkish text: {e}")

print(f"Sesli dosyalar bu dizinde oluşturuldu: {output_directory}")

