import math2docx
from docx import Document
import re
import os
import sys

def normalize_sqrt(latex_str):
    """
    Normalisasi perintah \sqrt untuk kompatibilitas dengan Word MathML.
    - Ubah \sqrt[2]{...} -> \sqrt{...}
    - Ubah \sqrt abc (tanpa kurung) -> \sqrt{abc}
    - Perbaiki kurung kurawal bertingkat: \sqrt{...} tetap, \sqrt[2]{...} diubah.
    """
    # 1. Ubah \sqrt[2]{...} menjadi \sqrt{...}
    #    Pola: \\sqrt\[2\]\{([^}]*)\}  -- ini hanya tangkap sampai } pertama
    #    Untuk menangani kurung bersarang, kita gunakan regex non-greedy dan balance
    #    Namun untuk sederhana, kita tangkap hingga } yang seimbang (gunakan rekursif?)
    #    Kita gunakan fungsi untuk mencari kurung kurawal seimbang
    def replace_sqrt2(match):
        # match grup 1 adalah isi di dalam {...}
        return r'\sqrt{' + match.group(1) + r'}'

    # Dengan asumsi tidak ada kurung kurawal bersarang di dalam argumen \sqrt[2]
    latex_str = re.sub(r'\\sqrt\[2\]\{([^}]*)\}', replace_sqrt2, latex_str)

    # 2. Ubah bentuk singkat: \sqrt abc -> \sqrt{abc}
    #    Pola: \\sqrt([^{])  -> menangkap satu karakter setelah \sqrt yang bukan '{'
    latex_str = re.sub(r'(\\sqrt)([^{])', r'\1{\2}', latex_str)

    # 3. (Opsional) Ubah \sqrt{} dengan argumen yang mengandung spasi menjadi \sqrt{...}
    #    Sudah ditangani oleh pola di atas jika ada spasi? Tidak, karena spasi bukan karakter.
    #    Biarkan saja.

    return latex_str

def clean_latex(latex_str):
    """Bersihkan delimiter LaTeX (\[ \], \( \), $, $$)."""
    latex_str = re.sub(r'\\\[', '', latex_str)
    latex_str = re.sub(r'\\\]', '', latex_str)
    latex_str = re.sub(r'\\\(', '', latex_str)
    latex_str = re.sub(r'\\\)', '', latex_str)
    latex_str = re.sub(r'\$\$', '', latex_str)
    latex_str = re.sub(r'\$', '', latex_str)
    return latex_str.strip()

def split_blocks(content):
    """Pisahkan blok berdasarkan baris kosong."""
    lines = content.splitlines()
    blocks = []
    current = []
    for line in lines:
        if line.strip() == '':
            if current:
                blocks.append('\n'.join(current))
                current = []
        else:
            current.append(line)
    if current:
        blocks.append('\n'.join(current))
    return blocks

def main():
    input_file = "rumus.txt"
    output_file = "output.docx"

    if not os.path.isfile(input_file):
        print(f"❌ File '{input_file}' tidak ditemukan.")
        sys.exit(1)

    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()

    if not content.strip():
        print("❌ File kosong.")
        sys.exit(1)

    blocks = split_blocks(content)
    if not blocks:
        print("❌ Tidak ada blok LaTeX ditemukan.")
        sys.exit(1)

    print(f"🔍 Ditemukan {len(blocks)} blok rumus.")

    doc = Document()
    for idx, block in enumerate(blocks, 1):
        # Normalisasi sqrt terlebih dahulu
        normalized = normalize_sqrt(block)
        cleaned = clean_latex(normalized)

        if not cleaned:
            print(f"⚠️ Blok {idx} kosong, dilewati.")
            continue

        try:
            p = doc.add_paragraph()
            math2docx.add_math(p, cleaned)
        except Exception as e:
            print(f"❌ Gagal konversi blok {idx}: {e}")
            # Jika gagal, coba fallback: ubah \sqrt menjadi pangkat
            fallback = re.sub(r'\\sqrt\{([^}]*)\}', r'(\1)^{1/2}', cleaned)
            try:
                p = doc.add_paragraph()
                math2docx.add_math(p, fallback)
                print(f"   ↳ Fallback pangkat berhasil untuk blok {idx}.")
            except Exception as e2:
                print(f"   ✖ Fallback juga gagal: {e2}")

    # Simpan
    try:
        doc.save(output_file)
        print(f"✅ Berhasil menyimpan {output_file} dengan {len(doc.paragraphs)} persamaan.")
    except PermissionError:
        base, ext = os.path.splitext(output_file)
        new_name = f"{base}_baru{ext}"
        doc.save(new_name)
        print(f"⚠️ File '{output_file}' sedang terbuka. Disimpan sebagai '{new_name}'")
        output_file = new_name

    # Buka otomatis
    try:
        os.startfile(output_file)
        print("📂 File dibuka otomatis.")
    except Exception as e:
        print(f"⚠️ Tidak bisa membuka file: {e}")

if __name__ == "__main__":
    main()