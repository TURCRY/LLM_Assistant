import os
import csv
from PIL import Image
import piexif

# --- Paramètres à ajuster ---
SOURCE_DIR = r"C:\\Users\\Utilisateur\\Documents\\NTU-Consult\\4 - Dossiers\\4 - Judiciaire\\4.5 - Préventifs avec suivi\\J04 - OPH Vallée Sud Clamart TA95\\C - Photos Audios\\Photos du 02 juillet 2024\\JPG"
DEST_DIR = os.path.join(os.path.dirname(SOURCE_DIR), "JPG reduits")
CSV_PATH = os.path.join(os.path.dirname(SOURCE_DIR), "photos_reduits.csv")

# Qualité JPEG (valeur estimée de vos réductions actuelles)
JPEG_QUALITY = 63  # Ajustable selon votre estimation
SUBSAMPLING = "4:2:0"  # Options: "4:4:4", "4:2:2", "4:2:0"
PROGRESSIVE = True
# ----------------------------

def get_subsampling_code(ss):
    mapping = {"4:4:4": 0, "4:2:2": 1, "4:2:0": 2}
    return mapping.get(ss, 2)

def compress_and_save(src_path, dest_path):
    try:
        img = Image.open(src_path)
        exif_bytes = None
        if "exif" in img.info:
            exif_bytes = img.info["exif"]
        elif hasattr(img, "info") and "exif" in img.info:
            exif_bytes = img.info["exif"]
        # Compression avec conservation EXIF
        img.save(
            dest_path,
            format="JPEG",
            quality=JPEG_QUALITY,
            subsampling=get_subsampling_code(SUBSAMPLING),
            progressive=PROGRESSIVE,
            optimize=True,
            exif=exif_bytes
        )
        return True
    except Exception as e:
        print(f"Erreur sur {src_path}: {e}")
        return False

def main():
    if not os.path.exists(DEST_DIR):
        os.makedirs(DEST_DIR)
    with open(CSV_PATH, mode="w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["fichier_original", "fichier_reduit", "qualite", "subsampling", "progressif"])
        for fname in os.listdir(SOURCE_DIR):
            if fname.lower().endswith((".jpg", ".jpeg")):
                src_path = os.path.join(SOURCE_DIR, fname)
                dest_path = os.path.join(DEST_DIR, fname)
                ok = compress_and_save(src_path, dest_path)
                if ok:
                    writer.writerow([fname, fname, JPEG_QUALITY, SUBSAMPLING, PROGRESSIVE])
    print(f"Compression terminée. Fichiers enregistrés dans: {DEST_DIR}")
    print(f"CSV généré: {CSV_PATH}")

if __name__ == "__main__":
    main()
