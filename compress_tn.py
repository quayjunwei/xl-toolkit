import os
import shutil
import zipfile
import tempfile
from PIL import Image

EXCEL_FILE = r"PATH//TO//EXCEL"
COMPRESSION_QUALITY = 70


def compress_excel_images(excel_file: str, compression_quality: int = 70) -> str:
    """
    Compress embedded images inside an Excel (.xlsx) file by recompressing
    images in the internal `xl/media` directory as JPEG.

    The original file is overwritten with the compressed version.
    A temporary file is used during processing to ensure the original
    is not corrupted if an error occurs.

    Args:
        excel_file: Path to the .xlsx file to compress.
        compression_quality: JPEG quality (1-95). Lower = smaller file, lower quality.

    Returns:
        Path to the compressed file (same as input).
    """

    if not os.path.exists(excel_file):
        raise FileNotFoundError(f"File not found: {excel_file}")

    if not (1 <= compression_quality <= 95):
        raise ValueError("compression_quality must be between 1 and 95")

    print("=" * 60)
    print("Excel Image Compression Script")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp(prefix="excel_compress_")

    try:
        print("\nExtracting Excel contents...")
        with zipfile.ZipFile(excel_file, "r") as zip_ref:
            zip_ref.extractall(temp_dir)

        media_folder = os.path.join(temp_dir, "xl", "media")

        total_before = 0
        total_after = 0

        if os.path.exists(media_folder):
            image_files = [
                f
                for f in os.listdir(media_folder)
                if f.lower().endswith((".png", ".jpg", ".jpeg"))
            ]

            print(f"Found {len(image_files)} images")

            for filename in image_files:
                img_path = os.path.join(media_folder, filename)

                size_before = os.path.getsize(img_path)
                total_before += size_before

                img = Image.open(img_path)
                if img.mode in ("RGBA", "P", "LA"):
                    img = img.convert("RGB")

                temp_file = os.path.join(media_folder, f"_temp_{filename}")
                img.save(temp_file, "JPEG", quality=compression_quality, optimize=True)
                img.close()

                os.remove(img_path)
                os.rename(temp_file, img_path)

                size_after = os.path.getsize(img_path)
                total_after += size_after

                reduction = ((size_before - size_after) / size_before) * 100
                print(
                    f"{filename}: {size_before/1024:.1f}KB → {size_after/1024:.1f}KB (-{reduction:.1f}%)"
                )

            if total_before > 0:
                print(
                    f"\nTotal images: {total_before/1024/1024:.2f}MB → {total_after/1024/1024:.2f}MB"
                )
                print(
                    f"Image reduction: {((total_before - total_after) / total_before) * 100:.1f}%"
                )
        else:
            print("No images found in Excel file.")

        print("\nRebuilding Excel file...")

        # Write to temp first to avoid corrupting original if something goes wrong
        temp_output = excel_file + ".tmp"
        if os.path.exists(temp_output):
            os.remove(temp_output)

        with zipfile.ZipFile(temp_output, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)

        shutil.rmtree(temp_dir)

        original_size = os.path.getsize(excel_file)
        compressed_size = os.path.getsize(temp_output)
        reduction = ((original_size - compressed_size) / original_size) * 100

        # Atomically replace original with compressed version
        os.replace(temp_output, excel_file)

        print("\nDone.")
        print(f"Original:   {original_size/1024/1024:.2f} MB")
        print(f"Compressed: {compressed_size/1024/1024:.2f} MB")
        if reduction > 0:
            print(f"Reduction:  {reduction:.1f}%")
        else:
            print(f"Size change: +{abs(reduction):.1f}%")
        print(f"\nSaved to: {excel_file}")

        return excel_file

    finally:
        if os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass


if __name__ == "__main__":
    compress_excel_images(
        excel_file=EXCEL_FILE, compression_quality=COMPRESSION_QUALITY
    )
