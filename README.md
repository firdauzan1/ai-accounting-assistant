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
1. Clone repository:
```bash
git clone https://github.com/firdauzan/ai-accounting-assistant.git
cd ai-accounting-assistant

2. Install dependencies:
```bash
pip install -r requirements.txt

3. Download spaCy model:
```bash
python -m spacy download en_core_web_sm

4. Jalankan aplikasi:
```bash
python machine_learning.py

Login Credentials:

username: admin
password: admin

5. Setup Data Files
Pastikan file berikut ada di folder yang sesuai:

logo.png (untuk icon aplikasi)
data/chart_of_account.xlsx (daftar akun)
data/normal_account_position.xlsx (posisi normal akun)

### Cara Menjalankan:
```bash
python machine_learning.py

