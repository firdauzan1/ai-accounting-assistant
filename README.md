# AI Accounting Assistant

Sistem Akuntansi berbasis AI untuk prediksi jurnal otomatis.

## Fitur Utama
- Prediksi akun debit/kredit otomatis 
- Generator laporan keuangan (Income Statement, Balance Sheet, Retained Earnings)
- Buku besar
- Jurnal penyesuaian
- Multi-currency support (Rupiah, USD, Euro)
- Import/Export Excel
- Machine Learning dengan feedback learning

## Instalasi

### Prerequisites
- Python 3.8+
- pip

### Langkah Instalasi
# Clone repository:

git clone https://github.com/firdauzan/ai-accounting-assistant.git

cd ai-accounting-assistant

# Install dependencies:

pip install -r requirements.txt

# Download spaCy model:

python -m spacy download en_core_web_sm

# Jalankan aplikasi:

python machine_learning.py

Login Credentials:

username: admin

password: admin

# Setup Data Files
Pastikan file berikut ada di folder yang sesuai:

logo.png (untuk icon aplikasi)

data/chart_of_account.xlsx (daftar akun)

data/normal_account_position.xlsx (posisi normal akun)

# Cara Menjalankan:
python machine_learning.py

